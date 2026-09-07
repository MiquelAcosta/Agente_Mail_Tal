# -*- coding: utf-8 -*-
"""FASE 1 EN OMBRA — llegeix mails, els classifica i APUNTA. No mou, no respon, no toca res.

Dos lectors intercanviables:
  python src\\demo\\sombra.py                  -> lector de CARPETA (fitxers .txt de prova)
  python src\\demo\\sombra.py --outlook "NOM"  -> lector d'OUTLOOK real (via COM; cal Outlook
                                                 obert en aquesta maquina i pywin32)
  python src\\demo\\sombra.py --informe        -> nomes mostra l'estadistica acumulada

Format dels fitxers de prova (carpeta src\\demo\\bandeja_prueba\\, un .txt per mail):
  linia 1: remitent
  linia 2: assumpte
  linia 3: ADJUNTO o NADA
  linia 4 en endavant: cos del mail

Usa el mateix escenario.json (historial, BBDD, model). El registre es
src\\demo\\log_sombra.sqlite. Un mail ja processat no es reprocessa.
"""
import json, os, sys, sqlite3, urllib.request, argparse, hashlib
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
from reglas import triaje

RUTA_ESC = os.path.join(AQUI, "escenario.json")
CARPETA_PRUEBA = os.path.join(AQUI, "bandeja_prueba")

PROMPT_CLASIFICADOR = """Eres el clasificador del buzon de atencion de una empresa de reclamaciones de vehiculos. Devuelve EXCLUSIVAMENTE un JSON valido:
{"categoria": "<una>", "sospecha_sucesion": true|false, "cancelacion": true|false, "complejo": true|false, "repregunta_insatisfecha": true|false, "motivo": "<una frase>"}
Categorias: estado_reclamacion, falta_factura_precio, envio_documentacion, confirmacion_documentacion, problema_web_subida, elegibilidad_vehiculo, informacion_general, coste_comision, poderes_pleitos, titularidad_caso_especial, cancelacion_desistimiento, cortesia_breve, contacto_llamada, ambiguo
Reglas: sospecha_sucesion=true ante CUALQUIER mencion a fallecimiento/herencia/viudedad/"era cliente" (ante la duda, true). cancelacion=true si expresa voluntad de desistir. complejo=true si varias peticiones, enojo, excepciones o dudas. cortesia_breve = agradecimientos/acuses SIN peticion nueva. Nada fuera del JSON."""


# ---------------------------------------------------------------- utilidades
def cargar_escenario():
    if not os.path.exists(RUTA_ESC):
        print("No existe escenario.json — ejecuta antes probar_mail.py una vez."); sys.exit(1)
    esc = json.load(open(RUTA_ESC, encoding="utf-8-sig"))
    esc["historial"] = [h.strip().lower() for h in esc["historial"]]
    esc["bbdd"] = {k.strip().lower(): v for k, v in esc["bbdd"].items()}
    return esc


def clasificar(ia, asunto, cuerpo):
    payload = json.dumps({"model": ia["modelo"], "temperature": 0,
        "messages": [{"role": "system", "content": PROMPT_CLASIFICADOR},
                     {"role": "user", "content": f"HILO: (no disponible en sombra)\n\nMENSAJE:\nAsunto: {asunto}\nCuerpo: {cuerpo[:3000]}"}]}).encode("utf-8")
    req = urllib.request.Request(ia["base_url"].rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {ia['api_key']}"})
    with urllib.request.urlopen(req, timeout=300) as r:
        data = json.loads(r.read().decode("utf-8"))
    t = data["choices"][0]["message"]["content"].replace("```json", "").replace("```", "").strip()
    return json.loads(t[t.find("{"):t.rfind("}") + 1])


def log_init():
    con = sqlite3.connect(os.path.join(AQUI, "log_sombra.sqlite"))
    con.execute("""CREATE TABLE IF NOT EXISTS sombra(
        id INTEGER PRIMARY KEY, ts TEXT, mail_id TEXT UNIQUE, remitente TEXT,
        asunto TEXT, destino TEXT, motivo TEXT, categoria TEXT, flags TEXT)""")
    return con


def ya_procesado(con, mail_id):
    return con.execute("SELECT 1 FROM sombra WHERE mail_id=?", (mail_id,)).fetchone() is not None


# ---------------------------------------------------------------- lectores
def lector_carpeta():
    """Lector de proves: un .txt per mail a bandeja_prueba/."""
    if not os.path.isdir(CARPETA_PRUEBA):
        os.makedirs(CARPETA_PRUEBA)
        ejemplo = ("cliente.uno@mail.com\nEstado de mi reclamacion\nNADA\n"
                   "Hola, queria saber como va mi expediente. Gracias.")
        open(os.path.join(CARPETA_PRUEBA, "ejemplo1.txt"), "w", encoding="utf-8").write(ejemplo)
        print(f"(Creada {CARPETA_PRUEBA} amb un mail d'exemple — afegeix-hi .txt de prova)")
    for nombre in sorted(os.listdir(CARPETA_PRUEBA)):
        if not nombre.endswith(".txt"):
            continue
        lineas = open(os.path.join(CARPETA_PRUEBA, nombre), encoding="utf-8-sig").read().splitlines()
        if len(lineas) < 4:
            continue
        yield {"id": "fichero:" + nombre,
               "remitente": lineas[0].strip().lower(),
               "asunto": lineas[1].strip(),
               "cuerpo": "\n".join(lineas[3:]).strip(),
               "adjuntos": ([{"nombre": "doc.pdf", "is_inline": False, "bytes": 200000}]
                            if lineas[2].strip().upper() == "ADJUNTO" else [])}


def lector_outlook(nombre_buzon, max_mails=50, nombre_carpeta=None):
    """Lector real: mails del buzon indicat (NOMES LLEGEIX).
    Sense --carpeta llegeix la safata d'entrada; amb --carpeta, aquella carpeta
    (la busca recursivament pel nom, p. ex. 'PruebasAgente')."""
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = next((s for s in stores if nombre_buzon.lower() in s.Name.lower()), None)
    if store is None:
        print(f"Cap buzon conte '{nombre_buzon}'. Visibles: " + ", ".join(s.Name for s in stores))
        return

    def buscar(raiz, nombre):
        for i in range(raiz.Folders.Count):
            f = raiz.Folders.Item(i + 1)
            if f.Name.lower() == nombre.lower():
                return f
            sub = buscar(f, nombre)
            if sub:
                return sub
        return None

    if nombre_carpeta:
        inbox = buscar(store, nombre_carpeta)
        if inbox is None:
            print(f"No trobo la carpeta '{nombre_carpeta}' dins de '{store.Name}'."); return
    else:
        inbox = next((store.Folders.Item(i + 1) for i in range(store.Folders.Count)
                      if store.Folders.Item(i + 1).Name.lower() in
                      ("bandeja de entrada", "inbox", "safata d'entrada")), None)
        if inbox is None:
            print(f"No trobo la safata d'entrada a '{store.Name}'."); return
    print(f"Llegint: {store.Name} > {inbox.Name}\n")
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    n = 0
    for msg in items:
        try:
            if msg.Class != 43:
                continue
            if msg.SenderEmailType == "EX":
                ex = msg.Sender.GetExchangeUser()
                rem = ex.PrimarySmtpAddress if ex else "desconocido@exchange"
            else:
                rem = msg.SenderEmailAddress or "desconocido"
            adjuntos = []
            for a in msg.Attachments:
                fn = str(a.FileName)
                adjuntos.append({"nombre": fn, "is_inline": fn.lower().startswith("image00"),
                                 "bytes": getattr(a, "Size", 0)})
            yield {"id": "outlook:" + str(msg.EntryID),
                   "remitente": rem.lower(),
                   "asunto": str(msg.Subject or ""),
                   "cuerpo": str(msg.Body or "")[:5000],
                   "adjuntos": adjuntos}
            n += 1
            if n >= max_mails:
                return
        except Exception as e:
            print(f"  (mail il·legible, es salta: {e})")


# ---------------------------------------------------------------- principal
def procesar(mails, esc, con):
    ia = esc["ia"]
    nuevos = clasif = 0
    for m in mails:
        if ya_procesado(con, m["id"]):
            continue
        nuevos += 1
        destino, motivo, _ = triaje(m, set(esc["historial"]), esc["bbdd"])
        categoria = flags = ""
        if destino == "CIRCUITO":
            try:
                c = clasificar(ia, m["asunto"], m["cuerpo"])
                categoria = c.get("categoria", "ambiguo")
                flags = ", ".join(k for k in ("sospecha_sucesion", "cancelacion",
                                              "complejo", "repregunta_insatisfecha") if c.get(k))
                clasif += 1
            except Exception as e:
                categoria, flags = "ERROR_IA", str(e)[:80]
        print(f"  {m['remitente'][:32]:32} -> {destino:16} {categoria} {('['+flags+']') if flags else ''}")
        con.execute("INSERT OR IGNORE INTO sombra(ts,mail_id,remitente,asunto,destino,motivo,categoria,flags)"
                    " VALUES(?,?,?,?,?,?,?,?)",
                    (datetime.now(timezone.utc).isoformat(), m["id"], m["remitente"],
                     m["asunto"], destino, motivo, categoria, flags))
        con.commit()
    print(f"\nProcessats {nuevos} mails nous ({clasif} classificats amb IA). Res s'ha mogut ni respost.")


def informe(con):
    total = con.execute("SELECT COUNT(*) FROM sombra").fetchone()[0]
    if not total:
        print("Registre buit encara."); return
    print(f"OMBRA — {total} mails processats en total")
    print("\nPer desti:")
    for d, n in con.execute("SELECT destino, COUNT(*) FROM sombra GROUP BY destino ORDER BY 2 DESC"):
        print(f"  {d:18} {n:5}  ({n/total*100:.0f}%)")
    print("\nCategories del circuit:")
    for c, n in con.execute("SELECT categoria, COUNT(*) FROM sombra WHERE destino='CIRCUITO'"
                            " GROUP BY categoria ORDER BY 2 DESC"):
        print(f"  {c:28} {n:5}")
    fl = con.execute("SELECT COUNT(*) FROM sombra WHERE flags<>''").fetchone()[0]
    print(f"\nAmb flags encesos (anirien a huma tot i la categoria): {fl}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="NOM", help="llegir del buzon d'Outlook que contingui NOM")
    ap.add_argument("--carpeta", metavar="CARPETA", default=None,
                    help="llegir aquesta carpeta del buzon en lloc de la safata d'entrada")
    ap.add_argument("--max", type=int, default=50, help="max mails a llegir d'Outlook (def. 50)")
    ap.add_argument("--informe", action="store_true", help="nomes mostrar estadistica")
    args = ap.parse_args()
    esc = cargar_escenario()
    con = log_init()
    print("=" * 64)
    print(" FASE 1 EN OMBRA — nomes llegir i apuntar | model:", esc["ia"]["modelo"])
    print("=" * 64)
    if args.informe:
        informe(con)
    elif args.outlook:
        procesar(lector_outlook(args.outlook, args.max, args.carpeta), esc, con)
        print(); informe(con)
    else:
        print(f"Lector de proves: {CARPETA_PRUEBA}\n")
        procesar(lector_carpeta(), esc, con)
        print(); informe(con)