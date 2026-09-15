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
import json, os, re, sys, time, sqlite3, urllib.request, argparse

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

PROMPT_REDACTOR = """Eres el redactor de respuestas del buzon de atencion de una empresa de reclamaciones de vehiculos. Tono cercano y claro, frases cortas, cero jerga juridica, en castellano. Tratamiento SIEMPRE de usted (nunca tutees). CIERRE del mensaje: si pides al cliente que envie o facilite algo (documentos, datos, matricula), la ULTIMA linea es exactamente "Quedamos a la espera."; si solo informas o respondes, la ULTIMA linea es exactamente "Saludos cordiales,". PROHIBIDO firmar como "El equipo de atencion" o con cualquier nombre de equipo o persona: el cierre es la ultima linea del mensaje, sin firma.

Reglas INQUEBRANTABLES:
1. Solo afirmas datos que esten en DATOS VERIFICADOS o en el HILO. Si el dato necesario para responder NO esta, no lo inventes: confianza "baja".
2. Sin promesas de plazos ni resultados que no esten en DATOS VERIFICADOS.
3. Documento de salida: SOLO si el cliente PIDE explicitamente ese contenido en su mensaje. Si el cliente solo pregunta por el estado, agradece o hace una consulta puntual: "NINGUNO". Adjuntar un documento no pedido es un ERROR.
   Catalogo: D1 = Instructivo general del proceso (solo si pregunta como funciona el proceso). D2 = Guia de documentacion alternativa a la factura (solo si dice que no tiene o no encuentra la factura).
4. 50-130 palabras. Saludo con el nombre si consta, respuesta directa, siguiente paso si lo hay, despedida. Un solo tema.
5. Si el cliente dice que adjunta algo pero no consta: pide que lo reenvie, no confirmes recepciones. MATIZ: si el cliente afirma haber enviado documentacion RECIENTEMENTE (hoy, ayer, "acabo de...") y en DATOS figura como no recibida, NO se lo niegues rotundamente: indica que el registro de documentacion puede tardar unas horas en actualizarse y que lo verificaremos; si en unos dias no recibe confirmacion, que nos lo reenvie.
6. El nombre del cliente SOLO puede salir de DATOS VERIFICADOS o de la firma de su mensaje. NUNCA lo deduzcas de la direccion de email. Si no lo sabes: saluda sin nombre ("Buenos dias:").
7. NO REPITAS lo ya dicho: si en el HISTORIAL PREVIO o en el hilo citado del propio mensaje ya se le pidio un documento (p.ej. el modelo 576) o ya se le dio una informacion, NO lo vuelvas a pedir ni a mencionar. Responde SOLO a lo nuevo del mensaje actual.
8. ESTILO: nunca uses la formula "sobreprecio del cartel 2006-2013" ni menciones el rango de anyos al hablar del expediente de un cliente: di "su reclamación del cártel de coches" o simplemente "su expediente". Los anyos solo se mencionan al explicar la elegibilidad a un interesado nuevo.
9. Documentacion pendiente: SOLO pide un documento si de verdad falta y frena el avance. Factura y contrato de compra son EQUIVALENTES: si uno consta "Sí", NUNCA pidas el otro. Si el estado del expediente indica fase de informe pericial, demanda, remitido o cerrado: la documentacion YA esta completa, NO pidas nada. Como maximo UNA linea cordial y solo si procede.
10. FORMATO del correo (usa saltos de linea \n dentro del texto): saludo en su propia linea; linea en blanco; el cuerpo en 1-3 parrafos cortos separados por linea en blanco; linea en blanco; y el CIERRE ("Quedamos a la espera." o "Saludos cordiales,") como ultima linea.

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


_RE_EMAIL_SOLO = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
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
    # Xarxa: cap "De:" trobat -> primer email EXTERN escrit al cos (simulacions
    # directes i casos on el client menciona la seva adreca a pel)
    for m in _RE_EMAIL_SOLO.finditer(cuerpo or ""):
        email = m.group(0).lower()
        if not es_reenviador(email, internos):
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
                try:
                    from conector_bbdd import ficha_por_telefono
                    self._real_tel = ficha_por_telefono
                except Exception:
                    self._real_tel = None
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

    def por_telefono(self, telefono):
        if getattr(self, "_real_tel", None) is None:
            return None
        try:
            return self._real_tel(telefono)
        except Exception as e:
            print(f"  (avis: cerca per telefon ha fallat: {str(e)[:60]})")
            return None

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


def _salvar_json_roto(texto):
    """Ultima xarxa: pescar els camps amb regex quan el model escup JSON trencat."""
    out = {}
    m = _re.search(r'"respuesta"\s*:\s*"(.+?)"\s*(?=[,}]|"[a-z_]+"\s*:)', texto, _re.DOTALL)
    if m:
        out["respuesta"] = m.group(1).replace('\\n', '\n').replace('\\"', '"')
    for campo in ("categoria", "confianza", "documento_salida", "veredicto", "motivo"):
        m = _re.search(rf'"{campo}"\s*:\s*"([^"]*)"', texto)
        if m:
            out[campo] = m.group(1)
    for flag in ("sospecha_sucesion", "cancelacion", "complejo", "repregunta_insatisfecha"):
        m = _re.search(rf'"{flag}"\s*:\s*(true|false)', texto)
        if m:
            out[flag] = (m.group(1) == "true")
    return out if out else None


def llamar(ia, system, user, timeout=480, rapido=False):
    """Crida amb REINTENT. rapido=True usa el model 'modelo_rapido' si esta configurat
    (per a classificador/verificador: tasques simples, model petit = mes velocitat)."""
    modelo = ia.get("modelo_rapido") if (rapido and ia.get("modelo_rapido")) else None
    for intento in (1, 2):
        t = ""
        try:
            data = _llamada_cruda(ia, system, user, timeout, modelo=modelo)
            t = data["choices"][0]["message"]["content"].replace("```json", "").replace("```", "").strip()
            if not t:
                raise ValueError("el model ha retornat contingut BUIT")
            if "{" not in t:
                # prosa sense JSON: per al redactor, la prosa ES la resposta
                print("    (el model ha respost text pla sense JSON: s'aprofita com a resposta)")
                return {"respuesta": t, "confianza": "media", "documento_salida": "NINGUNO",
                        "categoria": "ambiguo", "veredicto": "", "motivo": "text pla del model"}
            return json.loads(t[t.find("{"):t.rfind("}") + 1], strict=False)
        except json.JSONDecodeError as e:
            try:
                salvado = _salvar_json_roto(t)
            except Exception:
                salvado = None
            if salvado:
                print("    (JSON trencat del model: camps rescatats amb la xarxa)")
                return salvado
            if intento == 2:
                print(f"    (contingut cru del model: {repr(t[:160])})")
                raise
            print(f"    (crida fallida: {str(e)[:60]} — pausa i reintent...)")
            time.sleep(3)
        except Exception as e:
            if intento == 2:
                raise
            print(f"    (crida fallida: {str(e)[:60]} — pausa i reintent...)")
            time.sleep(3)


def escalfar(ia):
    """Carrega el model abans de comencar, de manera visible i controlada."""
    print(" Escalfant el model (la primera carrega pot trigar 1-3 min)...", flush=True)
    try:
        _llamada_cruda(ia, "Responde solo OK", "ping", timeout=300)
        print(" Model carregat i llest.\n")
    except Exception as e:
        print(f" AVIS: no s'ha pogut escalfar el model: {str(e)[:80]}")
        print(" Esta corrent Ollama? (ollama serve)\n")


CATEGORIA_AGENTE = "Agente"  # etiqueta d'Outlook: marca indeleble de "ja tractat"


def te_marca_agente(msg):
    try:
        return CATEGORIA_AGENTE.lower() in str(msg.Categories or "").lower()
    except Exception:
        return False


def marcar_agente(msg, marcar_leido=False):
    """Posa l'etiqueta 'Agente' al mail (i opcionalment el marca llegit).
    La marca viu AL MAIL: sobreviu a neteges de registre i a moviments."""
    try:
        cats = str(msg.Categories or "")
        if CATEGORIA_AGENTE.lower() not in cats.lower():
            msg.Categories = (cats + "; " if cats else "") + CATEGORIA_AGENTE
        if marcar_leido:
            msg.UnRead = False
        msg.Save()
    except Exception as e:
        print(f"    (avis: no s'ha pogut marcar el mail: {str(e)[:60]})")


PROP_LAST_VERB = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
PROP_MSG_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"


def id_estable(msg):
    """Identificador que sobreviu a moviments de carpeta: el Message-ID d'internet.
    (L'EntryID canvia en moure el mail; el Message-ID viatja amb ell.)"""
    try:
        mid = msg.PropertyAccessor.GetProperty(PROP_MSG_ID)
        if mid:
            return "msgid:" + str(mid).strip()
    except Exception:
        pass
    return "outlook:" + str(msg.EntryID)


def ya_respondido(msg):
    """True si aquest mail ja te resposta enviada des d'aquesta bustia d'Outlook.
    Llegeix la propietat interna 'last verb' (102=respost, 103=respost a tots).
    104 (reenviat) NO compta com a respost."""
    try:
        v = msg.PropertyAccessor.GetProperty(PROP_LAST_VERB)
        return v in (102, 103)
    except Exception:
        return False


def _subcarpeta_por_nombres(store, nombres):
    try:
        for i in range(store.Folders.Count):
            f = store.Folders.Item(i + 1)
            if f.Name.lower() in nombres:
                return f
    except Exception:
        pass
    return None


def buscar_historial(inbox, enviados, email, max_por_lado=3):
    """Busca al buzon els ultims mails d'aquest client (rebuts) i les nostres
    respostes (enviats), i els retorna com a context cronologic per a la IA."""
    email = (email or "").strip().lower()
    if not email or "@" not in email:
        return ""
    piezas = []

    def _recoger(carpeta, campo_dasl, etiqueta, campo_fecha):
        if carpeta is None:
            return
        try:
            filtro = f"@SQL=\"{campo_dasl}\" LIKE '%{email}%'"
            items = carpeta.Items.Restrict(filtro)
            items.Sort(f"[{campo_fecha}]", True)
            n = 0
            for it in items:
                if getattr(it, "Class", 0) != 43:
                    continue
                fecha = getattr(it, campo_fecha, None)
                cuerpo = str(getattr(it, "Body", "") or "")[:300].replace("\r\n", " ").replace("\n", " ")
                piezas.append((fecha, etiqueta, str(getattr(it, "Subject", "") or "")[:60], cuerpo))
                n += 1
                if n >= max_por_lado:
                    break
        except Exception:
            pass

    _recoger(inbox, "urn:schemas:httpmail:senderemail", "CLIENTE", "ReceivedTime")
    _recoger(enviados, "urn:schemas:httpmail:displayto", "EQUIPO", "SentOn")
    if not piezas:
        return ""
    piezas.sort(key=lambda p: str(p[0]))
    lineas = []
    for fecha, quien, asunto, cuerpo in piezas[-6:]:
        f = str(fecha)[:16] if fecha else "?"
        lineas.append(f"[{f}] {quien} ({asunto}): {cuerpo}")
    return "\n".join(lineas)[:1500]


def log_init():
    con = sqlite3.connect(os.path.join(AQUI, "log_borradores.sqlite"))
    con.execute("""CREATE TABLE IF NOT EXISTS borradores(
        id INTEGER PRIMARY KEY, ts TEXT, mail_id TEXT UNIQUE, remitente TEXT,
        asunto TEXT, categoria TEXT, flags TEXT, confianza TEXT, doc TEXT,
        veredicto_ia TEXT, resultado TEXT, respuesta TEXT)""")
    return con


import re as _re


def formatear_respuesta(texto):
    """Garanteix format de correu cordial encara que el model escrigui un bloc:
    salutacio en linia propia, paragrafs, i firma en dues linies."""
    t = (texto or "").strip()
    if not t:
        return t
    # normalitzar salts existents
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    # fora qualsevol firma d'equip o "Un saludo" que el model hagi posat
    t = _re.sub(r"\s*Un saludo,?\s*(El equipo de atenci[oó]n\.?)?\s*$", "", t, flags=_re.IGNORECASE)
    t = _re.sub(r"\s*El equipo de atenci[oó]n\.?\s*$", "", t, flags=_re.IGNORECASE)
    t = _re.sub(r"\n\s*Un saludo,?\s*\n?", "\n", t, flags=_re.IGNORECASE)
    # tancament garantit: peticio -> "Quedamos a la espera." | informatiu -> "Saludos cordiales,"
    if not _re.search(r"(quedamos a la espera|saludos cordiales)\s*[.,]?\s*$", t, _re.IGNORECASE):
        pide = _re.search(r"(env[ií]e|env[ií]enos|nos mande|m[aá]ndenos|facil[ií]t|adjunte|reenv[ií]e|"
                          r"necesitamos que|puede indicarnos|puede facilitarnos|\?)", t, _re.IGNORECASE)
        t = t.rstrip() + "\n\n" + ("Quedamos a la espera." if pide else "Saludos cordiales,")
    # salutacio en linia propia: tallar despres de la PRIMERA frase (punt o dos punts)
    primera = t.split("\n", 1)[0]
    if _re.match(r"^(Buen|Estimad|Hola|Querid|Apreciad)", primera, _re.IGNORECASE):
        for m in _re.finditer(r"(.{3,70}?[.:])\s+(?=\S)", primera):
            tros = m.group(1)
            if _re.search(r"(^|\s)[A-ZÁÉÍÓÚ]\.$", tros):
                continue  # acaba en inicial ("M.") : no es final de salutacio
            t = t.replace(tros + " ", tros + "\n\n", 1)
            break
    # partir cossos massa llargs sense paragrafs: despres de frase, cada ~2 frases
    bloques = t.split("\n\n")
    nuevos = []
    for b in bloques:
        if "\n" not in b and len(b) > 320:
            frases = _re.split(r"(?<=[.!?]) ", b)
            mitad = len(frases) // 2
            b = " ".join(frases[:mitad]) + "\n\n" + " ".join(frases[mitad:])
        nuevos.append(b)
    t = "\n\n".join(nuevos)
    # neteja de salts triples
    t = _re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


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
    hist_on = str(esc.get("historial_buzon", "si")).lower() not in ("off", "no", "false", "0")
    inbox = _subcarpeta_por_nombres(store, {"bandeja de entrada", "inbox"}) if hist_on else None
    enviados = _subcarpeta_por_nombres(store, {"elementos enviados", "sent items", "enviados"}) if hist_on else None
    print(f"Llegint: {store.Name} > {carpeta.Name}" + ("   [MODE DRY: no es crearan esborranys]" if dry else ""))

    filtro_fuente = str(esc.get("solo_no_leidos", "si")).lower() not in ("no", "false", "off", "0")
    try:
        items = carpeta.Items.Restrict("[UnRead] = true") if filtro_fuente else carpeta.Items
    except Exception:
        items = carpeta.Items
    if filtro_fuente:
        print("    (filtre a la font: nomes mails NO llegits entren a la llista)")
    n = creados = 0
    omesos_registre = 0
    for msg in list(items):
        try:
            if msg.Class != 43:
                continue
            mail_id = id_estable(msg)
            if con.execute("SELECT 1 FROM borradores WHERE mail_id=?", (mail_id,)).fetchone():
                omesos_registre += 1
                continue
            n += 1
            if n > max_mails:
                break
            if te_marca_agente(msg):
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                print("    OMES: porta la marca 'Agente' (ja tractat, marca al propi mail)")
                con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (datetime.now(timezone.utc).isoformat(), mail_id, "", str(msg.Subject or ""),
                             "", "", "", "", "", "OMITIDO (marca Agente)", ""))
                con.commit()
                continue
            if str(esc.get("solo_no_leidos", "si")).lower() not in ("no", "false", "off", "0") and not msg.UnRead:
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                print("    OMES: mail ja llegit per algu de l'equip (solo_no_leidos actiu)")
                con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (datetime.now(timezone.utc).isoformat(), mail_id, "", str(msg.Subject or ""),
                             "", "", "", "", "", "OMITIDO (leido)", ""))
                con.commit()
                continue
            if ya_respondido(msg):
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                print("    OMES: ja te resposta enviada des d'aquesta bustia")
                marcar_agente(msg)
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
            if destino == "HUMANO" and motivo == "sin_ficha_bbdd":
                m_tel = re.search(r"\b[6789]\d{2}[\s.-]?\d{3}[\s.-]?\d{3}\b", asunto + " " + cuerpo)
                if m_tel:
                    f_tel = bbdd.por_telefono(m_tel.group(0))
                    if f_tel:
                        ficha = f_tel
                        destino, motivo = "CIRCUITO", "identificado_por_telefono"
                        aviso_mat = " [PER TELEFON: revisar identitat]"
                        print(f"    sense fitxa per email; telefon al missatge -> FITXA LOCALITZADA (revisar identitat)")
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
                # Historial real del buzon amb aquest client (abans de classificar: el fa servir)
                hilo_txt = buscar_historial(inbox, enviados, rem) or "(sin historial en el buzon)"
                if "(sin historial" not in hilo_txt:
                    print(f"    historial del buzon: {hilo_txt.count(chr(10)) + 1} missatges previs trobats")
                # Classificar sempre (amb o sense fitxa): la categoria tria la plantilla
                c = llamar(ia, PROMPT_CLASIFICADOR, f"HISTORIAL PREVIO con este cliente:\n{hilo_txt}\n\nMENSAJE ACTUAL:\nAsunto: {asunto}\nCuerpo: {cuerpo}", rapido=True)
                categoria = c.get("categoria", "ambiguo")
                t_low = (asunto + " " + cuerpo).lower()
                if categoria.startswith("cancelacion"):
                    categoria = "cancelacion_desistimiento"
                if ficha and categoria in ("informacion_general", "ambiguo") and any(
                        k in t_low for k in ("como va", "cómo va", "estado de", "novedades",
                                             "mi expediente", "mi reclamacion", "mi reclamación",
                                             "mi caso", "mi coche", "que se sabe", "qué se sabe")):
                    categoria = "estado_reclamacion"
                    print("    (categoria ajustada a estado_reclamacion: pregunta per l'estat i te fitxa)")
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
                    textos = pl.get("textos_aprobados", {})
                    # dieta: si alguna variant coincideix amb l'estat de la fitxa, injectar NOMES aquella
                    coincidentes = {s: t for s, t in textos.items()
                                    if s.replace("si estado = ", "")[:28].lower() in datos.lower()}
                    for situacion, texto in (coincidentes or textos).items():
                        guia += f"TEXTO APROBADO ({situacion}): {texto}\n"
                # SENSE FITXA: decisio DETERMINISTA en codi (copy-paste real, sense IA redactant)
                respuesta_directa = None
                if not ficha and "cancelacion" not in flags and categoria not in (
                        "cancelacion_desistimiento", "cortesia_breve", "contacto_llamada"):
                    pide_expediente = any(k in t_low for k in
                        ("mi expediente", "mi reclamacion", "mi reclamación", "mi caso",
                         "como va", "cómo va", "estado de mi", "mi demanda"))
                    if pide_expediente:
                        respuesta_directa = ("Buenos días:\n\nGracias por su mensaje. No localizamos "
                            "ningún expediente asociado a esta dirección de correo. Para poder ayudarle, "
                            "¿puede indicarnos la matrícula del vehículo o el correo electrónico con el "
                            "que se registró?\n\nQuedamos a la espera.")
                        avisos += " [PLANTILLA FIXA: peticio d'identificacio]"
                    elif "VIABLE" in cartel_info and "NO viable" not in cartel_info:
                        # dades completes i dins del periode: ara SI te sentit el text d'alta
                        pl_nuevo = PLANTILLAS.get("clientes_nuevos", {})
                        textos_alta = list(pl_nuevo.get("textos_aprobados", {}).values())
                        if textos_alta:
                            respuesta_directa = textos_alta[0]
                            avisos += " [PLANTILLA FIXA: alta client nou (copy-paste)]"
                    elif any(k in t_low for k in ("modelo 576", "el 576", "modelo576", "no tengo factura",
                                                   "no encuentro la factura", "no conservo la factura",
                                                   "sin factura", "perdido la factura")):
                        pl_f = PLANTILLAS.get("falta_factura_precio", {})
                        textos_f = list(pl_f.get("textos_aprobados", {}).values())
                        if textos_f:
                            respuesta_directa = textos_f[0]
                            avisos += " [PLANTILLA FIXA: pasos modelo 576]"
                    elif "no indica" in cartel_info or "EN EL LIMITE" in cartel_info or not cartel_info:
                        # Falta informacio per resoldre la viabilitat. MAI demanar la data al
                        # client (feina interna): si ja ha donat la matricula -> "ho comprovem"
                        # + avis al treballador; si no -> demanar NOMES la matricula.
                        mat_en_text = REGEX_MATRICULA.search((asunto + " " + cuerpo).upper())
                        if mat_en_text:
                            respuesta_directa = ("Buenos días:\n\nGracias por su mensaje. Estamos "
                                "comprobando los datos de su vehículo para confirmarle si entra dentro "
                                "del periodo de afectación del cártel. Le responderemos en breve con la "
                                "confirmación y los pasos a seguir.\n\nSaludos cordiales,")
                            avisos += (" [TREBALLADOR: comprovar la data de matriculacio de la "
                                       f"matricula {mat_en_text.group(0).replace(' ', '')} (web) abans d'enviar]")
                        else:
                            respuesta_directa = ("Buenos días:\n\nGracias por su mensaje. Para poder "
                                "valorar su caso, ¿puede facilitarnos la matrícula del vehículo? Con ella "
                                "comprobaremos si entra dentro del periodo de afectación del cártel y le "
                                "indicaremos los pasos a seguir.\n\nQuedamos a la espera.")
                            avisos += " [PLANTILLA FIXA: peticio de matricula]"
                    # (si diu NO viable: cap al redactor amb la plantilla d'elegibilitat)
                if sucesion_bbdd or c.get("sospecha_sucesion"):
                    guia += ("SITUACION DE SUCESION: saludo formal (Estimado/a senyor/a o Buenos dias; NUNCA Querido/a), tono sobrio y humano, condolencias breves si procede, "
                             "explicar que una persona del equipo se hara cargo personalmente de su caso "
                             "y le contactara. NO detallar tramites ni datos del expediente.\n")

                if respuesta_directa is not None:
                    respuesta = formatear_respuesta(respuesta_directa)
                    confianza, doc = "plantilla", "NINGUNO"
                    print("    PLANTILLA FIXA -> copy-paste (sense redactor)")
                else:
                  print("    REDACTOR -> escribiendo...")
                  red = llamar(ia, PROMPT_REDACTOR,
                               f"DATOS VERIFICADOS: {datos}\nHISTORIAL PREVIO con este cliente:\n{hilo_txt}\nCATEGORIA: {categoria}\n{guia}"
                               f"MENSAJE del cliente:\nAsunto: {asunto}\nCuerpo: {cuerpo}")
                  confianza, doc, respuesta = red.get("confianza", ""), red.get("documento_salida", ""), red.get("respuesta", "")
                  respuesta = formatear_respuesta(respuesta)
                  print(f"    confianza: {confianza} | doc: {doc}")
                  if confianza == "baja":
                    avisos += " [CONFIANCA BAIXA del redactor]"

                if "adjunt" in respuesta.lower():
                    avisos += " [RECORDA ADJUNTAR els PDF que la resposta menciona abans d'enviar]"
                if respuesta_directa is not None or str(esc.get("verificador", "on")).lower() in ("off", "no", "false", "0"):
                    ver_txt = "OMITIDO"
                else:
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
                    marcar_agente(msg, marcar_leido=str(esc.get("marcar_leido", "si")).lower() in ("si", "sí", "true", "on", "1"))
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
    extra = f" ({omesos_registre} omesos en silenci: ja al registre)" if omesos_registre else ""
    print(f"\nFet: {min(n, max_mails)} mails processats, {creados} esborranys creats.{extra}" if not dry
          else f"\nFet (dry): {min(n, max_mails)} mails processats, 0 esborranys (mode assaig).{extra}")


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
    ap.add_argument("--registrar", action="store_true",
                    help="apuntar els mails de la carpeta al registre SENSE fer res (vacuna contra duplicats)")
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
    filtro_leidos = str(esc.get("solo_no_leidos", "si")).lower() not in ("no", "false", "off", "0")
    marcar_l = str(esc.get("marcar_leido", "si")).lower() in ("si", "sí", "true", "on", "1")
    print(" Filtre nomes-no-llegits:", "ACTIU" if filtro_leidos else "INACTIU",
          "| Marca 'Agente': sempre | Esborrany fet -> marcar llegit:", "SI" if marcar_l else "NO")
    if args.real:
        print(" MODE --real: nomes BBDD real; filtre d'historial desactivat (pilot)")
    print("=" * 64)
    if args.informe:
        informe(con)
    elif args.registrar and args.outlook:
        import win32com.client
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
        store = next((s for s in stores if args.outlook.lower() in s.Name.lower()), None)
        def _buscar(raiz, nombre):
            for i in range(raiz.Folders.Count):
                f = raiz.Folders.Item(i + 1)
                if f.Name.lower() == nombre.lower():
                    return f
                sub = _buscar(f, nombre)
                if sub:
                    return sub
            return None
        carpeta = _buscar(store, args.carpeta) if store else None
        if carpeta is None:
            print(f"No trobo '{args.carpeta}'."); os._exit(1)
        n = 0
        for msg in list(carpeta.Items):
            if getattr(msg, "Class", 0) != 43:
                continue
            mid = id_estable(msg)
            cur = con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                              " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                              (datetime.now(timezone.utc).isoformat(), mid, "", str(msg.Subject or ""),
                               "", "", "", "", "", "REGISTRADO (vacuna, sense esborrany)", ""))
            n += cur.rowcount
            marcar_agente(msg, marcar_leido=True)  # marca indeleble + llegit: veterans segellats
        con.commit()
        print(f"Vacunats {n} mails de '{args.carpeta}': marca 'Agente' + llegits + registre. Mai mes es tocaran.")
        con.close(); os._exit(0)
    elif args.outlook:
        escalfar(esc["ia"])
        procesar_carpeta(args.outlook, args.carpeta, esc, con, bbdd, args.dry, args.max)
        print(); informe(con)
        import gc; gc.collect()
        con.close()
        os._exit(0)
    else:
        print("Cal --outlook \"NOM\" (i opcionalment --carpeta, --dry, --max).")