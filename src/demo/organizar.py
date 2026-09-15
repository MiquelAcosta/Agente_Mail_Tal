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

PACK DE CLIENT (escenario.json -> "pack_cliente": "si"):
  Si un mateix client te DIVERSOS mails a la tanda, no es reparteixen per separat:
  es tracten com un PAQUET i es mouen TOTS al mateix calaix, decidit pel mail MES
  NOU (si l'ultim es desistiment, tot el pack va a 3 DESISTIMIENTO). Aixi el
  borradores els troba junts i la seva agrupacio pot respondre nomes el mes recent
  llegint els anteriors com a context. Sense aixo, dos mails del mateix client
  podien acabar en calaixos diferents i generar dos esborranys cecs.
  Salvaguarda: si un mail PREVI del pack demana persona si o si (adjunt real,
  sospita de successio, cancelacio), el pack escala a aquell calaix encara que el
  mes nou sigui inofensiu. Amb "pack_regla": "restrictivo" mana sempre el mes
  restrictiu del pack (DESISTIMIENTO > DIFICIL > FACIL) en lloc del mes nou.

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

CARPETAS = {"DESCARTES": "0 DESCARTES", "FACIL": "1 FACIL", "DIFICIL": "2 DIFICIL",
            "DESISTIMIENTO": "3 DESISTIMIENTO"}

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
_RE_EMAIL_SOLO = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
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
    # Xarxa: cap "De:" trobat -> primer email EXTERN escrit al cos (simulacions
    # directes i casos on el client menciona la seva adreca a pel)
    for m in _RE_EMAIL_SOLO.finditer(cuerpo or ""):
        email = m.group(0).lower()
        if not es_reenviador(email, internos):
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
    if categoria == "cancelacion_desistimiento" or "cancelacion" in (flags or ""):
        return "DESISTIMIENTO"    # esborrany d'acusament + tramit manual del treballador
    if flags:
        return "DIFICIL"          # sucesion/cancelacio/enfado/repregunta: persona
    if motivo == "sin_ficha_bbdd":
        return "FACIL"            # client nou o peticio d'identificacio: plantilla fixa
    if categoria in CATEGORIAS_FACIL or categoria in CATEGORIAS_MEDIO:
        return "FACIL"
    return "DIFICIL"              # titularidad, contacto_llamada, ambiguo...


# ---------------------------------------------------------------- PACK CLIENT
PRIORIDAD_CALAIX = {"DESCARTES": 0, "FACIL": 1, "DIFICIL": 2, "DESISTIMIENTO": 3}


def demana_persona(info):
    """El mail demana ull huma si o si, encara que no sigui el mes nou del pack:
    adjunt real, sospita de successio o voluntat de cancel.lar. Aquests no poden
    quedar sepultats dins d'un pack marcat com a FACIL pel mail de mes amunt."""
    if info["motivo"] == "adjunto_real" or info["destino"] == "HUMANO_SUCESION":
        return True
    if info["categoria"] == "cancelacion_desistimiento":
        return True
    flags = info["flags"] or ""
    return "sospecha_sucesion" in flags or "cancelacion" in flags


def decidir_pack(infos, regla="ultimo"):
    """infos: analisis dels mails d'UN client, ordenats de mes vell a mes nou.
    Retorna (calaix_del_pack, motiu_llegible)."""
    if regla == "restrictivo":
        final = max((i["calaix"] for i in infos), key=lambda c: PRIORIDAD_CALAIX[c])
        return final, "mana el mes restrictiu del pack"
    final = infos[-1]["calaix"]                       # per defecte: mana el mes NOU
    motiu = "mana el mail mes nou del client"
    durs = [i["calaix"] for i in infos[:-1] if demana_persona(i)]
    if durs:
        alt = max(durs, key=lambda c: PRIORIDAD_CALAIX[c])
        if PRIORIDAD_CALAIX[alt] > PRIORIDAD_CALAIX[final]:
            return alt, "escalat: un mail previ del client demana persona"
    return final, motiu


def clave_orden(info):
    """Data de recepcio comparable (amb l'ordre de la safata com a desempat)."""
    r = info.get("recibido")
    try:
        return (0, r.timestamp(), info["orden"])
    except Exception:
        pass
    try:
        return (0, float(r), info["orden"])
    except Exception:
        return (1, str(r), info["orden"])


def agrupar_por_cliente(analisis, regla="ultimo"):
    """Construeix els packs. Els DESCARTES no s'agrupen mai (no son clients) i
    un mail sense remitent identificable queda sol (millor separat que agrupat
    sota el reenviador equivocat). Retorna {id(msg): (calaix, motiu, pack)}."""
    grupos = {}
    for info in analisis:
        if info["calaix"] == "DESCARTES" or not info["rem"] or "@" not in info["rem"]:
            continue
        grupos.setdefault(info["rem"], []).append(info)
    decisiones = {}
    packs = []
    for cliente, infos in grupos.items():
        if len(infos) < 2:
            continue
        infos.sort(key=clave_orden)
        calaix, motiu = decidir_pack(infos, regla)
        packs.append((cliente, infos, calaix, motiu))
        for info in infos:
            decisiones[id(info["msg"])] = (calaix, motiu, True)
    return decisiones, packs


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
    pack_on = str(esc.get("pack_cliente", "si")).lower() not in ("no", "false", "off", "0")
    pack_regla = str(esc.get("pack_regla", "ultimo")).lower()
    # Congelar la llista abans de moure (moure mentre s'itera trenca l'index COM)
    mails = [m for m in list(carpeta.Items) if getattr(m, "Class", 0) == 43]
    analisis = []
    n = 0
    # ---- FASE A: analitzar-ho TOT sense moure res (cal la foto sencera de la
    # tanda per poder veure quins mails son del mateix client) --------------
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
            print(f"  {rem[:34]:34} -> {CARPETAS[calaix]:12} ({categoria or motivo}{' ['+flags+']' if flags else ''})")
            analisis.append({"msg": msg, "orden": n, "rem": rem, "asunto": asunto,
                             "recibido": getattr(msg, "ReceivedTime", None),
                             "destino": destino, "motivo": motivo, "categoria": categoria,
                             "flags": flags, "calaix": calaix})
        except Exception as e:
            print(f"  ERROR analitzant un mail (es continua): {str(e)[:100]}")

    # ---- FASE B: packs de client (mateix client -> mateix calaix, junts) ----
    decisiones = {}
    if pack_on:
        decisiones, packs = agrupar_por_cliente(analisis, pack_regla)
        if packs:
            print(f"\nPACKS DE CLIENT ({len(packs)}; regla: {pack_regla}) —"
                  " els mails d'un mateix client viatgen junts:")
            for cliente, infos, calaix, motiu in packs:
                print(f"  {cliente[:34]:34} {len(infos)} mails -> {CARPETAS[calaix]} ({motiu})")
                for i, info in enumerate(infos):
                    etiqueta = "MES NOU" if i == len(infos) - 1 else "previ   "
                    canvi = "" if info["calaix"] == calaix else f"  [{CARPETAS[info['calaix']]} -> {CARPETAS[calaix]}]"
                    print(f"      {etiqueta}  {str(info['asunto'])[:44]:44}{canvi}")
        else:
            print("\n(cap client amb diversos mails en aquesta tanda: res a empaquetar)")
    else:
        print("\n(packs de client DESACTIVATS a l'escenario: cada mail va pel seu compte)")

    # ---- FASE C: moure i registrar amb el calaix DEFINITIU ------------------
    for info in analisis:
        try:
            calaix, motiu_pack, en_pack = decisiones.get(id(info["msg"]),
                                                         (info["calaix"], "", False))
            recuento[calaix] += 1
            movido = 0
            if not dry:
                info["msg"].Move(destinos[calaix])
                movido = 1
            con.execute("INSERT INTO organizar(ts,remitente,asunto,destino,categoria,flags,calaix,movido)"
                        " VALUES(?,?,?,?,?,?,?,?)",
                        (datetime.now(timezone.utc).isoformat(), info["rem"], info["asunto"],
                         info["destino"], info["categoria"],
                         (info["flags"] + (" | PACK" if en_pack else "")).strip(" |"),
                         calaix, movido))
            con.commit()
        except Exception as e:
            print(f"  ERROR movent un mail (es continua): {str(e)[:100]}")
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