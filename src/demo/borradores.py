# -*- coding: utf-8 -*-
"""FASE 2 EN MAQUETA — el pipeline complet que acaba en ESBORRANY real a l'Outlook.

Recorregut per mail de la carpeta indicada:
  triatge -> fitxa (BBDD real o escenari) -> classificador -> redactor -> verificador
  -> si tot aprova: crea un ESBORRANY de resposta al fil (Reply + Save).

L'esborrany NO s'envia mai: queda a la carpeta Borradores del compte perque
una persona el revisi. Aquest programa NO conte cap ordre d'enviament.

Us (des de l'arrel del repo, amb Outlook obert):
  python src\\demo\\borradores.py --outlook "NOM" --carpeta "Tests Auto Cartel" --dry
  python src\\demo\\borradores.py --outlook "NOM" --carpeta "Tests Auto Cartel"
  python src\\demo\\borradores.py --informe

--dry: fa tot el proces pero NO crea l'esborrany (assaig complet).
Registre: src\\demo\\log_borradores.sqlite (no reprocessa mails ja vistos).
"""
import json, os, re, sys, sqlite3, urllib.request, argparse

PLANTILLAS = {}
TABLA_CARTEL = {}
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)
from reglas import triaje, REGEX_MATRICULA

RUTA_ESC = os.path.join(AQUI, "escenario.json")

MODO_CATEGORIA = {
    "estado_reclamacion": "AUTO", "confirmacion_documentacion": "AUTO",
    "problema_web_subida": "AUTO", "informacion_general": "AUTO",
    "coste_comision": "AUTO", "cortesia_breve": "AUTO",
    "elegibilidad_vehiculo": "BORRADOR", "falta_factura_precio": "BORRADOR",
    "poderes_pleitos": "BORRADOR", "envio_documentacion": "BORRADOR",
}

DOMINIOS_DESCARTES_BASE = [
    "microsoft.com", "microsoftonline.com", "office.com", "windows.com",
    "salesforce.com", "aircall.io", "google.com", "github.com",
    "linkedin.com", "mailchimp.com", "sendgrid.net", "amazonaws.com",
    "atlassian.com", "zoom.us", "docusign.com", "godaddy.com",
]


def es_dominio_descartes(rem, dominios):
    rem = (rem or "").lower()
    return any(rem.endswith("@" + d) or rem.endswith("." + d) for d in dominios)


PROMPT_CLASIFICADOR = """Eres el clasificador del buzon de atencion de una empresa de reclamaciones de vehiculos. Devuelve EXCLUSIVAMENTE un JSON valido:
{"categoria": "<una>", "sospecha_sucesion": true|false, "cancelacion": true|false, "complejo": true|false, "repregunta_insatisfecha": true|false, "motivo": "<una frase>"}
Categorias: estado_reclamacion, falta_factura_precio, envio_documentacion, confirmacion_documentacion, problema_web_subida, elegibilidad_vehiculo, informacion_general, coste_comision, poderes_pleitos, titularidad_caso_especial, cancelacion_desistimiento, cortesia_breve, contacto_llamada, fuera_de_contexto, ambiguo
Reglas: fuera_de_contexto = SOLO cuando el mensaje claramente NO tiene NINGUNA relacion con reclamaciones, coches, expedientes, documentacion o clientes (ej: publicidad, notificaciones de software, temas internos de oficina). OJO: los mails reenviados llevan cabeceras De:/Para: internas — IGNORALAS, juzga solo el contenido del mensaje del cliente. Si menciona reclamacion, expediente, coche, cartel, factura, demanda o documentacion: NUNCA fuera_de_contexto. Ante la duda: ambiguo, no fuera_de_contexto. sospecha_sucesion=true ante CUALQUIER mencion a fallecimiento/herencia/viudedad/"era cliente" (ante la duda, true). cancelacion=true si expresa voluntad de desistir. complejo=true si varias peticiones, enojo, excepciones o dudas. cortesia_breve = agradecimientos/acuses SIN peticion nueva. Nada fuera del JSON."""

PROMPT_REDACTOR = """Eres el redactor de respuestas del buzon de atencion de una empresa de reclamaciones de vehiculos. Tono cercano y claro, frases cortas, cero jerga juridica, en castellano. Tratamiento SIEMPRE de usted (nunca tutees). Firma SIEMPRE exactamente asi, en dos lineas finales: "Un saludo," y "El equipo de atencion".

Reglas INQUEBRANTABLES:
1. Solo afirmas datos que esten en DATOS VERIFICADOS o en el HILO. Si el dato necesario para responder NO esta, no lo inventes: confianza "baja".
2. Sin promesas de plazos ni resultados que no esten en DATOS VERIFICADOS.
3. Documento de salida: SOLO si el cliente PIDE explicitamente ese contenido en su mensaje. Si el cliente solo pregunta por el estado, agradece o hace una consulta puntual: "NINGUNO". Adjuntar un documento no pedido es un ERROR.
   Catalogo: D1 = Instructivo general del proceso (solo si pregunta como funciona el proceso). D2 = Guia de documentacion alternativa a la factura (solo si dice que no tiene o no encuentra la factura).
4. 50-130 palabras. Saludo con el nombre si consta, respuesta directa, siguiente paso si lo hay, despedida. Un solo tema.
5. Si el cliente dice que adjunta algo pero no consta: pide que lo reenvie, no confirmes recepciones.
6. El nombre del cliente SOLO puede salir de DATOS VERIFICADOS o de la firma de su mensaje. NUNCA lo deduzcas de la direccion de email. Si no lo sabes: saluda sin nombre ("Buenos dias:").

Devuelve EXCLUSIVAMENTE: {"respuesta":"<texto>","documento_salida":"D1|D2|NINGUNO","confianza":"alta|media|baja","motivo_confianza":"<una frase>"}"""

PROMPT_VERIFICADOR = """Eres el verificador final de respuestas de un buzon de atencion. Tu UNICA mision es detectar VIOLACIONES OBJETIVAS. Rechaza SOLO si ocurre alguna de estas cinco:
1. INVENCION: la respuesta afirma un dato concreto (fase, fecha, importe, referencia, recepcion de documentos) que NO aparece en DATOS VERIFICADOS ni en el MENSAJE del cliente.
2. TERCEROS: menciona a otras personas u otros expedientes.
3. PROMESAS: compromete plazos o resultados que no estan en DATOS VERIFICADOS.
4. DOCUMENTO NO PEDIDO: adjunta D1 o D2 sin que el cliente lo haya pedido explicitamente (D1 solo si pregunta como funciona el proceso; D2 solo si dice no tener factura).
5. TUTEO: trata al cliente de tu en lugar de usted.

PROHIBIDO rechazar por: estilo, tono, brevedad, falta de detalle, no mencionar temas adicionales, o porque la respuesta "podria ser mejor". La respuesta la revisara ademas una persona: tu solo filtras violaciones objetivas. Si no hay ninguna de las cinco: APROBADO.
Devuelve EXCLUSIVAMENTE: {"veredicto":"APROBADO|RECHAZADO","problemas":["..."],"riesgo":"bajo|medio|alto"}"""


RUTA_PLANTILLAS = os.path.join(AQUI, "plantillas.json")
RUTA_TABLA_CARTEL = os.path.join(AQUI, "tabla_cartel.json")
RE_ANYO = re.compile(r"\b(19[6-9]\d|20[0-2]\d)\b")
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
         "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
         "noviembre": 11, "diciembre": 12, "gener": 1, "febrer": 2, "marc": 3,
         "maig": 5, "juny": 6, "juliol": 7, "setembre": 9, "novembre": 11, "desembre": 12}


def cargar_tabla_cartel():
    if not os.path.exists(RUTA_TABLA_CARTEL):
        return {}
    try:
        return json.load(open(RUTA_TABLA_CARTEL, encoding="utf-8-sig"))["marcas"]
    except Exception:
        return {}


def comprobar_cartel(texto, tabla):
    """Comprovacio deterministica oficial: marca mencionada + any de matriculacio.
    Retorna una linia per a DATOS VERIFICADOS, o cap si no hi ha prou dades."""
    if not tabla:
        return ""
    t = " " + texto.upper() + " "
    marca = next((m for m in tabla if " " + m + " " in t or t.count(m) > 0), None)
    if not marca:
        return ""
    anyos = [int(a) for a in RE_ANYO.findall(texto)]
    anyo = next((a for a in anyos if 1990 <= a <= 2025), None)
    ini, fin = tabla[marca]
    a_ini, a_fin = int(ini[:4]), int(fin[:4])
    if anyo is None:
        return (f"COMPROBACION CARTEL: la marca {marca} SI esta en la tabla de marcas afectadas "
                f"(periodo {ini[8:]}/{ini[5:7]}/{ini[:4]} a {fin[8:]}/{fin[5:7]}/{fin[:4]}), pero el cliente no indica "
                f"el anyo de matriculacion: pedirselo para confirmar la viabilidad.")
    if a_ini > a_fin:  # rangs impossibles (TWINGO/DAIMLER): mai viable
        return f"COMPROBACION CARTEL: {marca} matriculado en {anyo}: NO viable segun la tabla oficial."
    mes = next((v for k, v in MESES.items() if k in texto.lower()), None)
    if mes and anyo in (a_ini, a_fin):
        dentro = (anyo, mes) >= (a_ini, int(ini[5:7])) and (anyo, mes) <= (a_fin, int(fin[5:7]))
        if dentro:
            return (f"COMPROBACION CARTEL: {marca} matriculado en {mes:02d}/{anyo}: DENTRO del periodo del cartel "
                    f"({ini[8:]}/{ini[5:7]}/{ini[:4]} a {fin[8:]}/{fin[5:7]}/{fin[:4]}): VIABLE por marca y fecha "
                    f"(pendiente de verificar compra nueva en Espana).")
        return (f"COMPROBACION CARTEL: {marca} matriculado en {mes:02d}/{anyo}: FUERA del periodo del cartel "
                f"({ini[8:]}/{ini[5:7]}/{ini[:4]} a {fin[8:]}/{fin[5:7]}/{fin[:4]}): NO viable por fecha.")
    if a_ini < anyo < a_fin:
        return (f"COMPROBACION CARTEL: {marca} matriculado en {anyo}: DENTRO del periodo del cartel "
                f"({ini[:4]}-{fin[:4]}): VIABLE por marca y fecha (pendiente de verificar compra nueva en Espana).")
    if anyo == a_ini or anyo == a_fin:
        return (f"COMPROBACION CARTEL: {marca} matriculado en {anyo}: el anyo esta EN EL LIMITE del periodo "
                f"({ini[8:]}/{ini[5:7]}/{ini[:4]} a {fin[8:]}/{fin[5:7]}/{fin[:4]}): depende del mes exacto; "
                f"pedir mes de matriculacion o documentacion antes de confirmar.")
    return (f"COMPROBACION CARTEL: {marca} matriculado en {anyo}: FUERA del periodo del cartel "
            f"({ini[:4]}-{fin[:4]}): NO viable por fecha.")


def cargar_plantillas():
    if not os.path.exists(RUTA_PLANTILLAS):
        return {}
    try:
        return json.load(open(RUTA_PLANTILLAS, encoding="utf-8-sig"))
    except Exception as e:
        print(f"(plantillas.json no cargado: {e})")
        return {}


def cargar_escenario():
    if not os.path.exists(RUTA_ESC):
        print("No existe escenario.json"); sys.exit(1)
    esc = json.load(open(RUTA_ESC, encoding="utf-8-sig"))
    esc["historial"] = [h.strip().lower() for h in esc["historial"]]
    esc["bbdd"] = {k.strip().lower(): v for k, v in esc["bbdd"].items()}
    return esc


RE_DE = re.compile(r"^\s*>?\s*(?:De|From|Von|A):?\s*(.{0,120}?)([\w.+-]+@[\w-]+(?:\.[\w-]+)+)",
                   re.IGNORECASE | re.MULTILINE)


def es_reenviador(rem, lista):
    """La whitelist accepta adreces exactes ('algu@empresa.com') o dominis
    sencers comencant per @ ('@empresa.com' = tothom del domini)."""
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
    """Busca al cos citat la linia 'De:/From:' del mail original i en treu l'email.
    Salta les adreces internes (els nostres dominis): en un fil reenviat, el primer
    'De:' pot ser un mail nostre de sortida — el client es el primer De: EXTERN."""
    for m in RE_DE.finditer(cuerpo or ""):
        etiqueta = m.group(0).strip().lower()
        if etiqueta.startswith(("para", "to", "a:")):
            continue
        email = m.group(2).lower()
        if es_reenviador(email, internos):
            continue  # adreca nostra citada al fil: seguir buscant el client
        return email
    return None


class HistorialAbierto:
    """Mode --real: el filtre d'historial es desactiva (identitat = fitxa BBDD).
    L'historial real per cerca al buzon arribara amb l'ombra v2."""
    def __contains__(self, x):
        return True


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
                self.modo = "BBDD REAL (+ escenario de reserva)"
            except Exception as e:
                print(f"(Connector BBDD no carregat: {e} — s'usa escenario.json)")

    def get(self, email, default=None):
        if self._real is not None:
            try:
                f = self._real(email)
                if f:
                    return f
            except Exception as e:
                print(f"  (avis: BBDD real ha fallat: {str(e)[:70]} — provant escenari)")
        return self.mock.get(email, default)

    def por_matricula(self, matricula):
        if self._real_mat is None:
            return None
        try:
            return self._real_mat(matricula)
        except Exception as e:
            print(f"  (avis: cerca per matricula ha fallat: {str(e)[:70]})")
            return None


def _llamada_cruda(ia, system, user, timeout, modelo=None):
    payload = json.dumps({"model": modelo or ia["modelo"], "temperature": 0, "max_tokens": 400,
        "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}).encode("utf-8")
    req = urllib.request.Request(ia["base_url"].rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {ia['api_key']}"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def llamar(ia, system, user, timeout=240, rapido=False):
    """Crida amb REINTENT. rapido=True usa el model 'modelo_rapido' si esta configurat
    (per a classificador/verificador: tasques simples, model petit = mes velocitat)."""
    modelo = ia.get("modelo_rapido") if (rapido and ia.get("modelo_rapido")) else None
    for intento in (1, 2):
        try:
            data = _llamada_cruda(ia, system, user, timeout, modelo=modelo)
            t = data["choices"][0]["message"]["content"].replace("```json", "").replace("```", "").strip()
            return json.loads(t[t.find("{"):t.rfind("}") + 1], strict=False)
        except Exception as e:
            if intento == 2:
                raise
            print(f"    (crida fallida: {str(e)[:60]} — reintentant, el model pot estar recarregant-se...)")


def escalfar(ia):
    """Carrega el model abans de comencar, de manera visible i controlada."""
    print(" Escalfant el model (la primera carrega pot trigar 1-3 min)...", flush=True)
    try:
        _llamada_cruda(ia, "Responde solo OK", "ping", timeout=300)
        print(" Model carregat i llest.\n")
    except Exception as e:
        print(f" AVIS: no s'ha pogut escalfar el model: {str(e)[:80]}")
        print(" Esta corrent Ollama? (ollama serve)\n")


PROP_LAST_VERB = "http://schemas.microsoft.com/mapi/proptag/0x10810003"


def ya_respondido(msg):
    """True si aquest mail ja te resposta enviada des d'aquesta bustia d'Outlook.
    Llegeix la propietat interna 'last verb' (102=respost, 103=respost a tots).
    104 (reenviat) NO compta com a respost."""
    try:
        v = msg.PropertyAccessor.GetProperty(PROP_LAST_VERB)
        return v in (102, 103)
    except Exception:
        return False


def log_init():
    con = sqlite3.connect(os.path.join(AQUI, "log_borradores.sqlite"))
    con.execute("""CREATE TABLE IF NOT EXISTS borradores(
        id INTEGER PRIMARY KEY, ts TEXT, mail_id TEXT UNIQUE, remitente TEXT,
        asunto TEXT, categoria TEXT, flags TEXT, confianza TEXT, doc TEXT,
        veredicto_ia TEXT, resultado TEXT, respuesta TEXT)""")
    return con


def ficha_a_datos(ficha):
    partes = [f"Titular: {ficha.get('titular','')}", f"Matricula: {ficha.get('matricula','')}"]
    if ficha.get("vehiculo"):
        partes.append(f"Vehiculo: {ficha['vehiculo']}")
    partes.append(f"Estado del expediente: {ficha.get('estado','')}")
    docs = ficha.get("documentacion")
    if docs:
        partes.append("Documentacion recibida: " + ", ".join(f"{k}: {v}" for k, v in docs.items()))
    return " · ".join(partes)


def procesar_carpeta(nombre_buzon, nombre_carpeta, esc, con, bbdd, dry, max_mails):
    global PLANTILLAS
    import win32com.client
    ia = esc["ia"]
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = next((s for s in stores if nombre_buzon.lower() in s.Name.lower()), None)
    if store is None:
        print(f"Cap buzon conte '{nombre_buzon}'. Visibles: " + ", ".join(s.Name for s in stores)); return

    def buscar(raiz, nombre):
        for i in range(raiz.Folders.Count):
            f = raiz.Folders.Item(i + 1)
            if f.Name.lower() == nombre.lower():
                return f
            sub = buscar(f, nombre)
            if sub:
                return sub
        return None

    carpeta = buscar(store, nombre_carpeta)
    if carpeta is None:
        print(f"No trobo la carpeta '{nombre_carpeta}' dins de '{store.Name}'."); return
    print(f"Llegint: {store.Name} > {carpeta.Name}" + ("   [MODE DRY: no es crearan esborranys]" if dry else ""))

    n = creados = 0
    for msg in list(carpeta.Items):
        try:
            if msg.Class != 43:
                continue
            mail_id = "outlook:" + str(msg.EntryID)
            if con.execute("SELECT 1 FROM borradores WHERE mail_id=?", (mail_id,)).fetchone():
                continue
            n += 1
            if n > max_mails:
                break
            if ya_respondido(msg):
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                print("    OMES: ja te resposta enviada des d'aquesta bustia")
                con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (datetime.now(timezone.utc).isoformat(), mail_id, "", str(msg.Subject or ""),
                             "", "", "", "", "", "OMITIDO (ya respondido)", ""))
                con.commit()
                continue
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
            reenviadores = esc.get("reenviadores", [])
            if es_reenviador(rem, reenviadores):
                extraido = extraer_cliente_de_reenvio(cuerpo, internos=reenviadores)
                if extraido:
                    print(f"\n=== reenviament de {rem}")
                    print(f"    client extret del cos: {extraido}")
                    rem = extraido
            mail = {"remitente": rem, "asunto": asunto, "cuerpo": cuerpo, "adjuntos": adjuntos}

            print(f"=== {rem} | {asunto[:50]}")
            historial = esc["historial"] if isinstance(esc["historial"], HistorialAbierto) else set(esc["historial"])
            destino, motivo, ficha = triaje(mail, historial, bbdd)
            aviso_mat = ""
            if destino == "HUMANO" and motivo == "sin_ficha_bbdd":
                m_mat = REGEX_MATRICULA.search((asunto + " " + cuerpo).upper())
                if m_mat:
                    mat = m_mat.group(0).replace(" ", "").upper()
                    f_mat = bbdd.por_matricula(mat)
                    if f_mat:
                        ficha = f_mat
                        destino, motivo = "CIRCUITO", f"identificado_por_matricula:{mat}"
                        aviso_mat = " [PER MATRICULA: revisar titularitat]"
                        print(f"    sense fitxa per email; matricula {mat} al missatge -> FITXA LOCALITZADA (revisar titularitat)")
            sucesion_bbdd = bool(ficha and ficha.get("sospecha_sucesion_bbdd"))
            categoria = flags = confianza = doc = ver_txt = respuesta = ""
            resultado = destino + " (" + motivo + ")"
            avisos = aviso_mat

            # ESTRATEGIA: BORRADOR PER A TOT, excepte adjunts, sistema i DESCARTES.
            dominios = DOMINIOS_DESCARTES_BASE + [d.strip().lower().lstrip("@")
                                                  for d in esc.get("dominios_descartes", [])]
            if es_dominio_descartes(rem, dominios):
                destino, motivo = "DESCARTE", "dominio_no_cliente"
            sin_borrador = (destino in ("SISTEMA", "DESCARTE")) or (motivo == "adjunto_real")
            if sin_borrador:
                resultado = destino + f" ({motivo}) — SENSE esborrany (per disseny)"
            else:
                # Classificar sempre (amb o sense fitxa): la categoria tria la plantilla
                c = llamar(ia, PROMPT_CLASIFICADOR, f"HILO: (no disponible)\n\nMENSAJE:\nAsunto: {asunto}\nCuerpo: {cuerpo}", rapido=True)
                categoria = c.get("categoria", "ambiguo")
                flags = ", ".join(k for k in ("sospecha_sucesion", "cancelacion", "complejo",
                                              "repregunta_insatisfecha") if c.get(k))
                print(f"    categoria: {categoria} | flags: {flags or 'ninguno'}")
                if categoria == "fuera_de_contexto" and ficha:
                    # XARXA: un client identificat a la BBDD mai es descarte
                    print("    (fuera_de_contexto ignorat: el remitent TE fitxa -> es tracta com ambiguo)")
                    categoria = "ambiguo"
                if categoria == "fuera_de_contexto":
                    resultado = "DESCARTE (fuera_de_contexto per IA) — SENSE esborrany"
                    print(f"    RESULTADO -> {resultado}")
                    con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto,
                                 categoria, flags, "", "", "", resultado, ""))
                    con.commit()
                    continue
                if flags:
                    avisos += " [FLAGS: " + flags + " — revisar amb cura]"
                if sucesion_bbdd or c.get("sospecha_sucesion"):
                    avisos += " [SUCESSIO: to especialment curos, revisar sempre]"

                # Dades: fitxa real o avis de client no identificat
                cartel_info = comprobar_cartel(asunto + " " + cuerpo, TABLA_CARTEL)
                if ficha:
                    datos = ficha_a_datos(ficha)
                    print(f"    ficha: {datos[:100]}...")
                else:
                    datos = ("SIN FICHA EN BBDD: remitente no identificado. Posibles casos: "
                             "cliente nuevo interesado, o cliente que escribe desde un email "
                             "no registrado. NO afirmar nada de ningun expediente.")
                    avisos += " [SENSE FITXA: possible client nou]"
                    print("    sense fitxa: esborrany de client nou / peticio d'identificacio")
                if cartel_info:
                    datos += " · " + cartel_info
                    print(f"    {cartel_info[:90]}...")

                # Plantilla: la de la categoria; sense fitxa, la de clients nous
                pl = PLANTILLAS.get(categoria, {})
                guia = ""
                if pl.get("guia"):
                    modo_pl = ("COPIA CASI LITERAL: usa el texto aprobado que coincida TAL CUAL, "
                               "cambiando unicamente el saludo con el nombre del cliente y los datos "
                               "concretos de su ficha. No reescribas ni resumas."
                               if pl.get("literal") else "PLANTILLA A SEGUIR (adapta con naturalidad):")
                    guia = f"{modo_pl} {pl['guia']}\n"
                    for situacion, texto in pl.get("textos_aprobados", {}).items():
                        guia += f"TEXTO APROBADO ({situacion}): {texto}\n"
                if not ficha:
                    ref = PLANTILLAS.get("_referencia_equipo_no_automatizable", {})
                    guia += ("SITUACION SIN FICHA — elige segun el mensaje: (a) si es un interesado "
                             "nuevo que quiere reclamar, usa este TEXTO APROBADO de alta: "
                             + ref.get("clientes_nuevos", "") +
                             " (b) si pregunta por un expediente existente, responde que no localizamos "
                             "su expediente con este correo y pidele amablemente la matricula del vehiculo "
                             "o el email con el que se registro.\n")
                if sucesion_bbdd or c.get("sospecha_sucesion"):
                    guia += ("SITUACION DE SUCESION: saludo formal (Estimado/a senyor/a o Buenos dias; NUNCA Querido/a), tono sobrio y humano, condolencias breves si procede, "
                             "explicar que una persona del equipo se hara cargo personalmente de su caso "
                             "y le contactara. NO detallar tramites ni datos del expediente.\n")

                print("    REDACTOR -> escribiendo...")
                red = llamar(ia, PROMPT_REDACTOR,
                             f"DATOS VERIFICADOS: {datos}\nHILO: (no disponible)\nCATEGORIA: {categoria}\n{guia}"
                             f"MENSAJE del cliente:\nAsunto: {asunto}\nCuerpo: {cuerpo}")
                confianza, doc, respuesta = red.get("confianza", ""), red.get("documento_salida", ""), red.get("respuesta", "")
                print(f"    confianza: {confianza} | doc: {doc}")
                if confianza == "baja":
                    avisos += " [CONFIANCA BAIXA del redactor]"

                print("    VERIFICADOR -> revisando...")
                try:
                    ver = llamar(ia, PROMPT_VERIFICADOR,
                                 f"DATOS VERIFICADOS: {datos}\nMENSAJE del cliente: {cuerpo}\n"
                                 f"RESPUESTA PROPUESTA (doc: {doc}): {respuesta}", rapido=True)
                    ver_txt = ver.get("veredicto", "")
                    if ver_txt != "APROBADO":
                        avisos += " [VERIFICADOR: " + "; ".join(ver.get("problemas", []))[:120] + "]"
                except Exception as e:
                    avisos += f" [verificador no disponible: {str(e)[:40]}]"

                if dry:
                    resultado = "BORRADOR (dry: no creado)" + avisos
                else:
                    reply = msg.Reply()
                    reply.Body = respuesta  # nomes el missatge generat, sense fil citat
                    destino_seguro = esc.get("borradores_para", "").strip()
                    reply.To = destino_seguro if destino_seguro else rem
                    reply.Save()  # <- ESBORRANY. Mai .Send()
                    creados += 1
                    resultado = ("BORRADOR CREADO (per a: "
                                 + (destino_seguro if destino_seguro else rem) + ")" + avisos)
                print("    ---- RESPUESTA " + "-" * 38)
                for lin in respuesta.split("\n"):
                    print(f"    | {lin}")
                print("    " + "-" * 53)
            print(f"    RESULTADO -> {resultado}")
            con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                        " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                        (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto, categoria,
                         flags, confianza, doc, ver_txt, resultado, respuesta))
            con.commit()
        except Exception as e:
            print(f"    ERROR con este mail (se continua): {str(e)[:120]}")
    print(f"\nFet: {min(n, max_mails)} mails processats, {creados} esborranys creats." if not dry
          else f"\nFet (dry): {min(n, max_mails)} mails processats, 0 esborranys (mode assaig).")


def informe(con):
    total = con.execute("SELECT COUNT(*) FROM borradores").fetchone()[0]
    if not total:
        print("Registre buit."); return
    print(f"BORRADORES — {total} mails processats")
    for r, nn in con.execute("SELECT resultado, COUNT(*) FROM borradores GROUP BY resultado ORDER BY 2 DESC"):
        print(f"  {r[:60]:60} {nn}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="NOM")
    ap.add_argument("--carpeta", metavar="CARPETA", default="Tests Auto Cartel")
    ap.add_argument("--max", type=int, default=10)
    ap.add_argument("--dry", action="store_true", help="assaig: tot el proces sense crear esborranys")
    ap.add_argument("--real", action="store_true",
                    help="mode 100%% BBDD: fitxes nomes de la BBDD real, filtre d'historial desactivat")
    ap.add_argument("--informe", action="store_true")
    args = ap.parse_args()
    esc = cargar_escenario()
    PLANTILLAS.update(cargar_plantillas())
    TABLA_CARTEL.update(cargar_tabla_cartel())
    con = log_init()
    if args.real:
        esc["bbdd"] = {}   # cap fitxa de pont: nomes la BBDD real
        esc["historial"] = HistorialAbierto()
    bbdd = BBDDConectada(esc["bbdd"] if not args.real else {})
    if args.real:
        bbdd.modo = "NOMES BBDD REAL (escenario: sols configuracio)"
    print("=" * 64)
    print(" FASE 2 EN MAQUETA — esborranys | model:", esc["ia"]["modelo"])
    print(" Font de fitxes:", bbdd.modo, "| Aquest programa NO pot enviar res.")
    if args.real:
        print(" MODE --real: nomes BBDD real; filtre d'historial desactivat (pilot)")
    print("=" * 64)
    if args.informe:
        informe(con)
    elif args.outlook:
        escalfar(esc["ia"])
        procesar_carpeta(args.outlook, args.carpeta, esc, con, bbdd, args.dry, args.max)
        print(); informe(con)
        import gc; gc.collect()
        con.close()
        os._exit(0)
    else:
        print("Cal --outlook \"NOM\" (i opcionalment --carpeta, --dry, --max).")