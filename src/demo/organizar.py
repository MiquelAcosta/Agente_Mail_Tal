# -*- coding: utf-8 -*-
"""FASE 1 REAL — ORGANITZADOR DE SAFATA: classifica i MOU els mails a subcarpetes.

Per cada mail de la carpeta indicada decideix el seu calaix i EL MOU:
  0 DESCARTES  -> correu de sistema, proveidors, partners, newsletters: res de clients
  1 FACIL      -> client identificat + categoria automatitzable + sense banderes
  2 MEDIO      -> client identificat + categoria de mode esborrany
  3 DIFICIL    -> tot el que demana persona: sense fitxa, successions, cancelacions,
                  banderes (enfado, repregunta, multi-tema), adjunts, ambigus

Les subcarpetes es creen soles (dins de la carpeta processada) la primera vegada.
MOURE es reversible: els mails no s'esborren mai, nomes canvien de calaix.

Us (Outlook obert):
  python src\\demo\\organizar.py --outlook "NOM" --carpeta "Bandeja de entrada" --dry
  python src\\demo\\organizar.py --outlook "NOM" --carpeta "Tests Auto Cartel"
  python src\\demo\\organizar.py --informe

--dry: mostra on aniria cada mail SENSE moure res (fes-ho sempre primer).
Registre: src\\demo\\log_organizar.sqlite
"""
import json, os, sys, sqlite3, urllib.request, argparse
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)
from reglas import triaje, REGEX_MATRICULA

RUTA_ESC = os.path.join(AQUI, "escenario.json")

CARPETAS = {"DESCARTES": "0 DESCARTES", "FACIL": "1 FACIL", "DIFICIL": "2 DIFICIL"}

CATEGORIAS_FACIL = {"estado_reclamacion", "confirmacion_documentacion",
                    "problema_web_subida", "informacion_general",
                    "coste_comision", "cortesia_breve"}
CATEGORIAS_MEDIO = {"elegibilidad_vehiculo", "falta_factura_precio",
                    "poderes_pleitos", "envio_documentacion"}

# Dominis que mai son clients (ampliable a escenario.json amb "dominios_descartes")
DOMINIOS_DESCARTES_BASE = [
    "microsoft.com", "microsoftonline.com", "office.com", "windows.com",
    "salesforce.com", "aircall.io", "google.com", "github.com",
    "linkedin.com", "mailchimp.com", "sendgrid.net", "amazonaws.com",
    "atlassian.com", "zoom.us", "docusign.com", "godaddy.com",
]

PROMPT_CLASIFICADOR = """Eres el clasificador del buzon de atencion de una empresa de reclamaciones del cartel de fabricantes de coches. Devuelve EXCLUSIVAMENTE un JSON valido:
{"categoria": "<una>", "sospecha_sucesion": true|false, "cancelacion": true|false, "complejo": true|false, "repregunta_insatisfecha": true|false, "motivo": "<una frase>"}
Categorias: estado_reclamacion, falta_factura_precio, envio_documentacion, confirmacion_documentacion, problema_web_subida, elegibilidad_vehiculo, informacion_general, coste_comision, poderes_pleitos, titularidad_caso_especial, cancelacion_desistimiento, cortesia_breve, contacto_llamada, fuera_de_contexto, ambiguo
Reglas: fuera_de_contexto = el mail NO es de un cliente sobre su reclamacion (proveedores, partners comerciales, notificaciones de servicios, publicidad, temas internos de empresa). sospecha_sucesion=true ante CUALQUIER mencion a fallecimiento/herencia/viudedad (ante la duda, true). cancelacion=true si expresa voluntad de desistir. complejo=true si varias peticiones, enojo, excepciones o dudas. cortesia_breve = agradecimientos SIN peticion nueva. Nada fuera del JSON."""


import re
RE_DE = re.compile(r"^\s*>?\s*(?:De|From|Von|A):?\s*(.{0,120}?)([\w.+-]+@[\w-]+(?:\.[\w-]+)+)",
                   re.IGNORECASE | re.MULTILINE)


def es_reenviador(rem, lista):
    rem = (rem or "").strip().lower()
    for entrada in lista:
        e = entrada.strip().lower()
        if not e:
            continue
        if e.startswith("@"):
            if rem.endswith(e):
                return True
        elif rem == e:
            return True
    return False


def extraer_cliente_de_reenvio(cuerpo, internos=()):
    for m in RE_DE.finditer(cuerpo or ""):
        etiqueta = m.group(0).strip().lower()
        if etiqueta.startswith(("para", "to", "a:")):
            continue
        email = m.group(2).lower()
        if es_reenviador(email, internos):
            continue
        return email
    return None


def cargar_escenario():
    esc = json.load(open(RUTA_ESC, encoding="utf-8-sig"))
    esc["historial"] = [h.strip().lower() for h in esc["historial"]]
    esc["bbdd"] = {k.strip().lower(): v for k, v in esc["bbdd"].items()}
    return esc


class BBDDConectada:
    def __init__(self, esc_bbdd):
        self.mock = esc_bbdd
        self.modo = "escenario.json"
        self._real = None
        self._real_mat = None
        if os.path.exists(os.path.join(AQUI, "credenciales.json")):
            try:
                from conector_bbdd import ficha_por_email, ficha_por_matricula
                self._real = ficha_por_email
                self._real_mat = ficha_por_matricula
                self.modo = "BBDD REAL"
            except Exception as e:
                print(f"(Connector BBDD no carregat: {e})")

    def get(self, email, default=None):
        if self._real is not None:
            try:
                f = self._real(email)
                if f:
                    return f
            except Exception as e:
                print(f"  (avis BBDD: {str(e)[:60]})")
        return self.mock.get(email, default)

    def por_matricula(self, matricula):
        if self._real_mat is None:
            return None
        try:
            return self._real_mat(matricula)
        except Exception:
            return None


class HistorialAbierto:
    def __contains__(self, x):
        return True


def llamar(ia, system, user, timeout=240):
    payload = json.dumps({"model": ia["modelo"], "temperature": 0, "max_tokens": 300,
        "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}).encode("utf-8")
    req = urllib.request.Request(ia["base_url"].rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {ia['api_key']}"})
    for intento in (1, 2):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = json.loads(r.read().decode("utf-8"))
            t = data["choices"][0]["message"]["content"].replace("```json", "").replace("```", "").strip()
            return json.loads(t[t.find("{"):t.rfind("}") + 1], strict=False)
        except Exception as e:
            if intento == 2:
                raise
            print(f"    (crida fallida: {str(e)[:50]} — reintentant...)")


def es_dominio_descartes(rem, dominios):
    rem = (rem or "").lower()
    return any(rem.endswith("@" + d) or rem.endswith("." + d) for d in dominios)


def calaix_de(destino, categoria, flags, motivo=""):
    """Regla de repartiment AMPLIADA: FACIL = tot el que te resposta clara de
    plantilla (candidats al futur enviament automatic). DIFICIL = el que demana
    ull huma de veritat: adjunts, successions, cancelacions, enfados, ambigus."""
    if categoria == "fuera_de_contexto" or destino in ("SISTEMA", "DESCARTE"):
        return "DESCARTES"
    if motivo == "adjunto_real" or destino == "HUMANO_SUCESION":
        return "DIFICIL"
    if flags:
        return "DIFICIL"          # sucesion/cancelacio/enfado/repregunta: persona
    if motivo == "sin_ficha_bbdd":
        return "FACIL"            # client nou o peticio d'identificacio: plantilla fixa
    if categoria in CATEGORIAS_FACIL or categoria in CATEGORIAS_MEDIO:
        return "FACIL"
    return "DIFICIL"              # titularidad, contacto_llamada, ambiguo...


def subcarpeta(carpeta, nombre):
    for i in range(carpeta.Folders.Count):
        f = carpeta.Folders.Item(i + 1)
        if f.Name.lower() == nombre.lower():
            return f
    return carpeta.Folders.Add(nombre)


def log_init():
    con = sqlite3.connect(os.path.join(AQUI, "log_organizar.sqlite"))
    con.execute("""CREATE TABLE IF NOT EXISTS organizar(
        id INTEGER PRIMARY KEY, ts TEXT, remitente TEXT, asunto TEXT,
        destino TEXT, categoria TEXT, flags TEXT, calaix TEXT, movido INTEGER)""")
    return con


def procesar(nombre_buzon, nombre_carpeta, esc, con, bbdd, dry, max_mails):
    import win32com.client
    ia = esc["ia"]
    dominios = DOMINIOS_DESCARTES_BASE + [d.strip().lower().lstrip("@")
                                          for d in esc.get("dominios_descartes", [])]
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = next((s for s in stores if nombre_buzon.lower() in s.Name.lower()), None)
    if store is None:
        print(f"Cap buzon conte '{nombre_buzon}'."); return

    def buscar(raiz, nombre):
        for i in range(raiz.Folders.Count):
            f = raiz.Folders.Item(i + 1)
            if f.Name.lower() == nombre.lower():
                return f
            sub = buscar(f, nombre)
            if sub:
                return sub
        return None

    carpeta = buscar(store, nombre_carpeta) if nombre_carpeta else None
    if carpeta is None:
        print(f"No trobo la carpeta '{nombre_carpeta}' a '{store.Name}'."); return
    print(f"Organitzant: {store.Name} > {carpeta.Name}"
          + ("   [DRY: nomes mostrar, res es mou]" if dry else ""))

    destinos = {} if dry else {k: subcarpeta(carpeta, v) for k, v in CARPETAS.items()}
    recuento = {k: 0 for k in CARPETAS}
    # Congelar la llista abans de moure (moure mentre s'itera trenca l'index COM)
    mails = [m for m in list(carpeta.Items) if getattr(m, "Class", 0) == 43]
    n = 0
    for msg in mails:
        try:
            n += 1
            if n > max_mails:
                break
            if msg.SenderEmailType == "EX":
                ex = msg.Sender.GetExchangeUser()
                rem = (ex.PrimarySmtpAddress if ex else "desconocido@exchange").lower()
            else:
                rem = (msg.SenderEmailAddress or "desconocido").lower()
            asunto = str(msg.Subject or "")
            cuerpo = str(msg.Body or "")[:2500]
            adjuntos = [{"nombre": str(a.FileName),
                         "is_inline": str(a.FileName).lower().startswith("image00"),
                         "bytes": getattr(a, "Size", 0)} for a in msg.Attachments]
            mail = {"remitente": rem, "asunto": asunto, "cuerpo": cuerpo, "adjuntos": adjuntos}

            reenviadores = esc.get("reenviadores", [])
            if es_reenviador(rem, reenviadores):
                extraido = extraer_cliente_de_reenvio(cuerpo, internos=reenviadores)
                if extraido:
                    rem = extraido
                    mail["remitente"] = rem
            categoria = flags = ""
            if es_dominio_descartes(rem, dominios):
                destino, motivo = "SISTEMA", "dominio_no_cliente"
            else:
                destino, motivo, ficha = triaje(mail, HistorialAbierto(), bbdd)
                if destino == "HUMANO" and motivo == "sin_ficha_bbdd":
                    m_mat = REGEX_MATRICULA.search((asunto + " " + cuerpo).upper())
                    if m_mat:
                        f_mat = bbdd.por_matricula(m_mat.group(0).replace(" ", ""))
                        if f_mat:
                            ficha = f_mat
                            destino, motivo = "CIRCUITO", "identificado_por_matricula"
                if destino == "CIRCUITO" and ficha and ficha.get("sospecha_sucesion_bbdd"):
                    destino, motivo = "HUMANO_SUCESION", "herencia_en_bbdd"
                if destino == "CIRCUITO" or motivo == "sin_ficha_bbdd":
                    c = llamar(ia, PROMPT_CLASIFICADOR,
                               f"MENSAJE:\nAsunto: {asunto}\nCuerpo: {cuerpo}")
                    categoria = c.get("categoria", "ambiguo")
                    if categoria == "fuera_de_contexto" and ficha:
                        categoria = "ambiguo"  # client amb fitxa mai es descarte
                    flags = ", ".join(k for k in ("sospecha_sucesion", "cancelacion",
                                                  "complejo", "repregunta_insatisfecha") if c.get(k))
            calaix = calaix_de(destino, categoria, flags, motivo)
            recuento[calaix] += 1
            print(f"  {rem[:34]:34} -> {CARPETAS[calaix]:12} ({categoria or motivo}{' ['+flags+']' if flags else ''})")
            movido = 0
            if not dry:
                msg.Move(destinos[calaix])
                movido = 1
            con.execute("INSERT INTO organizar(ts,remitente,asunto,destino,categoria,flags,calaix,movido)"
                        " VALUES(?,?,?,?,?,?,?,?)",
                        (datetime.now(timezone.utc).isoformat(), rem, asunto, destino,
                         categoria, flags, calaix, movido))
            con.commit()
        except Exception as e:
            print(f"  ERROR amb un mail (es continua): {str(e)[:100]}")
    print("\nRepartiment: " + " · ".join(f"{CARPETAS[k]}: {v}" for k, v in recuento.items()))
    if dry:
        print("(DRY: res s'ha mogut. Treu --dry per organitzar de veritat.)")


def informe(con):
    total = con.execute("SELECT COUNT(*) FROM organizar").fetchone()[0]
    if not total:
        print("Registre buit."); return
    print(f"ORGANITZADOR — {total} mails processats")
    for c, nn in con.execute("SELECT calaix, COUNT(*) FROM organizar GROUP BY calaix ORDER BY 2 DESC"):
        print(f"  {CARPETAS.get(c, c):14} {nn:5}  ({nn/total*100:.0f}%)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="NOM")
    ap.add_argument("--carpeta", metavar="CARPETA", default="Tests Auto Cartel")
    ap.add_argument("--max", type=int, default=50)
    ap.add_argument("--dry", action="store_true", help="mostrar el reparto sin mover nada")
    ap.add_argument("--informe", action="store_true")
    args = ap.parse_args()
    esc = cargar_escenario()
    con = log_init()
    bbdd = BBDDConectada(esc["bbdd"])
    print("=" * 64)
    print(" ORGANITZADOR DE SAFATA (fase 1) | model:", esc["ia"]["modelo"], "| fitxes:", bbdd.modo)
    print(" Mou mails a subcarpetes. No respon, no esborra, no envia.")
    print("=" * 64)
    if args.informe:
        informe(con)
    elif args.outlook:
        procesar(args.outlook, args.carpeta, esc, con, bbdd, args.dry, args.max)
        import gc; gc.collect()
        con.close()
        os._exit(0)
    else:
        print('Cal --outlook "NOM" (i opcionalment --carpeta, --dry, --max).')