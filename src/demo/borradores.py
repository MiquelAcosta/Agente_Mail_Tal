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
from datetime import datetime, timezone, timedelta

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


def escoger_buzon(stores, nombre):
    """Tria la bustia. PRIORITAT: nom exacte -> comenca per -> conte.

    Sense aixo, 'info@recuperatudinero.com' tambe coincideix amb
    'Online Archive - info@recuperatudinero.com', i com que l'arxiu surt abans a
    la llista guanyava ell i no es trobava cap calaix. Si queda ambigu, s'avisa."""
    n = (nombre or "").strip().lower()
    noms = [str(getattr(s, "Name", "") or "") for s in stores]
    for prova in (lambda x: x.lower() == n, lambda x: x.lower().startswith(n)):
        tri = [s for s, x in zip(stores, noms) if prova(x)]
        if tri:
            return tri[0]
    conte = [s for s, x in zip(stores, noms) if n in x.lower()]
    if not conte:
        return None
    if len(conte) > 1:
        print(f"    (AVIS: {len(conte)} busties contenen '{nombre}': "
              + ", ".join(str(getattr(c, "Name", "")) for c in conte) + ")")
        print(f"    (s'usa '{getattr(conte[0], 'Name', '')}' — poseu el nom exacte)")
    return conte[0]


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
18. NO PIDAS NUNCA el año de matriculacion, la fecha de compra, la marca, el modelo ni el numero de bastidor: esos datos los tenemos nosotros en la base de datos a partir de la matricula. Si necesitas identificar el vehiculo, pide SOLO la matricula, y solo si el cliente no la ha dado ya.
17. El saludo va SIN NOMBRE: escribe exactamente "Buenos dias," y nunca "Buenos dias Leonor Cruz Gimenez,". No pongas el nombre ni los apellidos del cliente en ninguna parte del correo.
16. Si la categoria es de DESISTIMIENTO o cancelacion: NUNCA termines con "Quedamos a la espera.". El cliente quiere cerrar, no se le espera nada. Cierra con "Saludos cordiales,".
11. NUNCA digas que adjuntas, envias o dejas adjunto un documento, un PDF, una propuesta de honorarios o un formulario. NO PUEDES adjuntar archivos. Si el cliente necesita un documento, di que se lo haremos llegar, nunca que va adjunto.
12. NO PIDAS un dato que el cliente YA HA ESCRITO en su mensaje o en el hilo citado. Antes de pedir la matricula, el DNI, el telefono o el nombre, RELEELO: si esta, usalo. Pedir algo que acaba de dar es el peor error posible.
13. RESPONDE PRIMERO A LO QUE PREGUNTA. Si pregunta por el estado, habla del estado; si pregunta si esta en plazo, contesta eso. Solo despues, y solo si hace falta, pide algo. Una respuesta que ignora la pregunta y pide un dato es una respuesta fallida.
14. Si el cliente plantea una cuestion juridica (prescripcion, plazos legales, articulos de una ley, interpretacion de una norma) NO la resuelvas: confianza "baja" y di que un companyo del equipo se lo confirmara.
15. Si el cliente expresa espera larga o malestar ("hace meses", "ha pasado muchisimo tiempo", "no tengo noticias"), reconocelo en UNA linea antes de nada. No lo ignores.
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


# Categories que parlen de documents/arxius: el client n'envia, en confirma,
# en reclama o te problemes per pujar-los. Ampliable des de l'escenario.json amb
# "categorias_documentos". Es desactiva tot amb "documentos_sin_borrador": "no".
CATEGORIAS_DOCUMENTOS = {"envio_documentacion", "confirmacion_documentacion",
                         "problema_web_subida", "falta_factura_precio"}

# Frases de "ja us he enviat / us adjunto" — captura els mails que PARLEN
# d'arxius encara que no en portin cap.
# Senyal FORT per si sol: "os adjunto", "se adjunta", "adjuntamos"...
RE_DOC_FUERTE = re.compile(r"\b(adjunt[oa]s?|adjuntamos|adjuntando|anexo|anexamos)\b",
                           re.IGNORECASE)
# Senyal DEBIL: nomes compta si hi ha un verb d'enviar I un nom de document.
RE_DOC_VERBO = re.compile(
    r"\b(envi[eéoó]\w*|envia\w+|mand[eéoó]\w*|manda\w+|sub[ií]\w*|"
    r"remit[ií]\w*|pas[eéoó]\w*|aport[eéoó]\w*|entregu[eé]\w*)\b", re.IGNORECASE)
RE_DOC_NOMBRE = re.compile(
    r"\b(document\w+|factura\w*|papel\w+|archiv\w+|fichero\w*|justificant\w*|"
    r"pdf|foto\w*|copia\w*|dni|nie|permiso\s+de\s+circulaci\w+|ficha\s+t[eé]cnica|"
    r"contrato\s+de\s+compra|modelo\s*576|escritura\w*)\b", re.IGNORECASE)


def habla_de_documentos(asunto, cuerpo, solo_fuerte=False):
    """(bool, motiu) — el mail gira al voltant d'arxius encara que no en porti.

    Cal un senyal FORT ("os adjunto") o bé un verb d'enviar JUNT amb un nom de
    document ("ya os envie la factura"). Nomes el verb no compta: "os envio un
    saludo" no parla de documents."""
    t = (asunto or "") + " " + (cuerpo or "")
    m = RE_DOC_FUERTE.search(t)
    if m:
        return True, "menciona: " + m.group(0)[:30]
    if solo_fuerte:
        return False, ""
    v = RE_DOC_VERBO.search(t)
    n = RE_DOC_NOMBRE.search(t)
    if v and n:
        return True, f"menciona: {v.group(0)[:16]} + {n.group(0)[:16]}"
    return False, ""


CARPETA_MULTIPLES = "MULTIPLES"
CALAIX_DIFICIL = "2 DIFICIL"   # on van els casos que demanen una persona   # nom configurable: "carpeta_multiples"

# Bloc intern que es posa DINS de l'esborrany perque el treballador vegi el mail
# previ sense anar-lo a buscar. VA DELIMITAT expressament: aixi es pot esborrar
# d'una passada, i qualsevol enviament automatic futur l'ha de treure SEMPRE
# abans d'enviar (quitar_nota_interna) o el client el rebria.
NOTA_INI = "===== NOTA INTERNA — ESBORRAR ABANS D'ENVIAR ====="
NOTA_FIN = "===== FI DE LA NOTA INTERNA ====="
CATEGORIA_NO_AUTO = "Agente-NoAuto"


def _insertar_arriba_html(html_original, texto_plano):
    """Posa el text nou A DALT d'una resposta HTML d'Outlook, conservant el fil
    citat que Outlook ja ha muntat a sota. Si no es pot, torna None."""
    if not html_original:
        return None
    try:
        import html as _html
        bloque = ("<div style=\"font-family:Calibri,sans-serif;font-size:11pt\">"
                  + _html.escape(texto_plano).replace("\n", "<br>")
                  + "</div><br>")
        m = re.search(r"<body[^>]*>", html_original, re.IGNORECASE)
        if m:
            i = m.end()
            return html_original[:i] + bloque + html_original[i:]
        return bloque + html_original
    except Exception:
        return None


def cita_propia(msg, max_chars=4000):
    """Construeix la citacio del mail original pel nostre compte.

    Outlook no sempre inclou el text original en una resposta: depen de la
    configuracio ("No incluir el mensaje original"), del format del mail i de si
    el cos es va quedar buit. Si falta, la posem nosaltres: el treballador ha de
    poder llegir que deia el client sense sortir de l'esborrany."""
    try:
        de = remitente_smtp(msg) or str(getattr(msg, "SenderName", "") or "")
        cuando = fecha_texto(componentes_fecha(msg))
        asunto = str(getattr(msg, "Subject", "") or "")
        cuerpo = str(getattr(msg, "Body", "") or "")[:max_chars]
    except Exception:
        return ""
    return ("\n\n-----------------------------------------------\n"
            f"De: {de}\nEnviado: {cuando}\nAsunto: {asunto}\n\n{cuerpo}")


def falta_la_cita(texto_respuesta, msg):
    """True si al cos de l'esborrany no hi ha rastre del mail del client."""
    try:
        original = str(getattr(msg, "Body", "") or "").strip()
    except Exception:
        return False
    if len(original) < 25:
        return False          # mail sense cos: no hi ha res a citar
    # Es busca un tros significatiu del text original dins de l'esborrany.
    trozo = " ".join(original.split())[:40]
    return trozo not in " ".join((texto_respuesta or "").split())


def respuesta_vacia(t):
    """True si, despres de netejar, no queda missatge de veritat.

    Si l'unic que deia la resposta era una peticio que hem eliminat (l'any de
    matriculacio, el bastidor...), queda nomes salutacio i comiat. Aixo no es
    pot enviar: val mes que ho miri una persona."""
    if not t:
        return True
    cos = _re.sub(r"^(Buenos d[ií]as|Buenas tardes|Buenas noches|Hola)[,.:]?\s*", "",
                  str(t).strip(), flags=_re.IGNORECASE)
    cos = _re.sub(r"(Quedamos a la espera\.?|Saludos cordiales,?|"
                  r"Quedamos a su disposici[oó]n[^.]*\.?)\s*$", "", cos.strip(),
                  flags=_re.IGNORECASE).strip()
    return len(cos) < 25


def quitar_nota_interna(texto):
    """Treu el bloc intern d'un text d'esborrany. OBLIGATORI abans de qualsevol
    enviament automatic."""
    if not texto or NOTA_INI not in texto:
        return texto
    ini = texto.index(NOTA_INI)
    fin = texto.find(NOTA_FIN)
    if fin == -1:
        return texto[:ini].strip()
    return (texto[:ini] + texto[fin + len(NOTA_FIN):]).strip()


def tiene_nota_interna(texto):
    return bool(texto) and NOTA_INI in texto


def tiene_adjuntos_reales(msg, min_bytes=8000):
    """True si el mail porta arxius de veritat (no firmes ni imatges incrustades)."""
    try:
        for i in range(msg.Attachments.Count):
            a = msg.Attachments.Item(i + 1)
            nom = str(getattr(a, "FileName", "") or "").lower()
            if nom.endswith((".png", ".gif", ".jpg", ".jpeg", ".bmp")) and \
               int(getattr(a, "Size", 0) or 0) < min_bytes:
                continue          # imatge de firma
            return True
    except Exception:
        pass
    return False


def _buscar_o_crear(padre, nombre):
    for i in range(padre.Folders.Count):
        f = padre.Folders.Item(i + 1)
        if f.Name.lower() == nombre.lower():
            return f
    return padre.Folders.Add(nombre)


def mover_a_calaix(msg, carpeta_origen, nombre):
    """Mou un mail a un calaix GERMA del que s'esta processant (1 FACIL ->
    2 DIFICIL, per exemple). Si ja hi es, no fa res. Torna True si s'ha mogut."""
    try:
        if str(getattr(carpeta_origen, "Name", "")).lower() == nombre.lower():
            return False
        padre = carpeta_origen.Parent
        if padre is None:
            return False
        destino = _buscar_o_crear(padre, nombre)
        movido = msg.Move(destino)
        try:
            donde = str(movido.Parent.Name)
            if donde.lower() != nombre.lower():
                print(f"    ATENCIO: ha acabat a '{donde}', no a '{nombre}'.")
        except Exception:
            pass
        return True
    except Exception as e:
        print(f"    (no s'ha pogut moure a {nombre}: {str(e)[:70]})")
        return False


def mover_a_multiples(msg, carpeta_origen, nombre=CARPETA_MULTIPLES):
    """Aparta un germa del pack a la carpeta MULTIPLES.

    N'hi ha UNA de sola per a tota la bustia: es crea al costat dels calaixos
    (germana de '1 FACIL', '2 DIFICIL'...), no dins de cadascun. Aixi el
    treballador te un unic lloc on mirar, vinguin els mails del calaix que
    vinguin. Torna True si s'ha mogut."""
    try:
        destino = None
        try:
            padre = carpeta_origen.Parent      # la carpeta que conte els calaixos
            if padre is not None:
                destino = _buscar_o_crear(padre, nombre)
        except Exception:
            destino = None
        if destino is None:                    # sense pare accessible: dins del calaix
            destino = _buscar_o_crear(carpeta_origen, nombre)
        movido = msg.Move(destino)
        # Verificacio: confirmar ON ha acabat de veritat. Si Outlook el posa en
        # un altre lloc (regles, neteja de converses), ha de quedar escrit.
        try:
            donde = str(movido.Parent.Name)
            if donde.lower() != nombre.lower():
                print(f"    ATENCIO: el mail ha acabat a '{donde}', no a '{nombre}'."
                      " Reviseu regles d'Outlook o la neteja de converses.")
            else:
                print(f"    (verificat: ara es a {donde})")
        except Exception:
            pass
        return True
    except Exception as e:
        print(f"    (no s'ha pogut moure a {nombre}: {str(e)[:70]})")
        return False


def mas_nuevo_en_hilo(msg, propia=""):
    """Torna el mail MES NOU del mateix fil (si n'hi ha un de mes nou que aquest),
    o None.

    NO escaneja la bustia: fa servir la conversa que Outlook ja te muntada per a
    cada mail — la mateixa informacio que alimenta la vista "Mostrar mensajes en
    conversaciones". Es una consulta local i barata, per mail.
    S'ignoren els esborranys i els correus enviats per nosaltres: nomes compten
    els mails que ha escrit el client."""
    try:
        conv = msg.GetConversation()
        if conv is None:
            return None
        tabla = conv.GetTable()
    except Exception:
        return None          # el magatzem no suporta converses: es continua sense
    mia = componentes_fecha(msg)
    if not mia:
        return None
    try:
        mio_id = str(msg.EntryID)
    except Exception:
        mio_id = None
    sess = getattr(msg, "Session", None)
    if sess is None:
        return None
    propia = (propia or "").strip().lower()
    mejor = None
    try:
        while not tabla.EndOfTable:
            fila = tabla.GetNextRow()
            try:
                eid = str(fila("EntryID"))
            except Exception:
                continue
            if mio_id and eid == mio_id:
                continue
            try:
                otro = sess.GetItemFromID(eid)
            except Exception:
                continue
            if getattr(otro, "Class", 0) != 43:
                continue
            try:
                if getattr(otro, "Sent", True) is False:
                    continue          # esborrany
            except Exception:
                pass
            if propia and (remitente_smtp(otro) or "") == propia:
                continue              # resposta nostra, no del client
            f = componentes_fecha(otro)
            if f and f > mia and (mejor is None or f > mejor[0]):
                mejor = (f, otro)
    except Exception:
        return None
    return mejor[1] if mejor else None


def mails_de_carpeta(carpeta, orden="antiguos", limite=50, dias=0):
    return mails_de_carpeta_items(carpeta.Items, orden, limite, dias)


def mails_de_carpeta_items(items, orden="antiguos", limite=50, dias=0):
    """Treu els mails d'una carpeta SENSE carregar-la sencera a memoria.

    list(carpeta.Items) materialitza un objecte COM per cada correu: amb milers,
    l'Outlook (32 bits) es queda sense recursos i cau ("Out of memory or system
    resources"). Aqui es demana a Outlook que ORDENI ell (rapid, ho fa sobre la
    seva taula) i despres es recorre amb GetFirst/GetNext agafant NOMES els que
    calen. La memoria usada depen del limit, no de la mida de la bustia.

    dias > 0 limita a mes a mes la finestra temporal (Restrict), cosa que redueix
    encara mes el treball d'Outlook."""
    desc = orden in ("recientes", "nuevos", "desc")
    if dias and int(dias) > 0:
        try:
            desde = (datetime.now() - timedelta(days=int(dias))).strftime("%m/%d/%Y %I:%M %p")
            items = items.Restrict(f"[ReceivedTime] >= '{desde}'")
        except Exception as e:
            print(f"    (no s'ha pogut limitar per dies: {str(e)[:60]})")
    try:
        items.Sort("[ReceivedTime]", desc)     # ordena Outlook, no Python
    except Exception as e:
        print(f"    (Outlook no ha pogut ordenar: {str(e)[:60]} — ordre natural)")
    out = []
    try:
        m = items.GetFirst()
        while m is not None and len(out) < limite:
            try:
                if getattr(m, "Class", 0) == 43:
                    out.append(m)
            except Exception:
                pass
            m = items.GetNext()
    except Exception as e:
        print(f"    (lectura interrompuda: {str(e)[:70]})")
    print(f"    llegits {len(out)} mails (limit {limite}"
          + (f", ultims {dias} dies" if dias else "") + ")")
    return out


RE_DNI = re.compile(r"\b\d{7,8}\s?[A-HJ-NP-TV-Z]\b", re.IGNORECASE)
RE_TEL = re.compile(r"\b[6789]\d{2}[\s.-]?\d{3}[\s.-]?\d{3}\b")


RE_AVIS_MS = re.compile(
    r"^[\s\[]*No suele recibir correo[^\n]*?"
    r"(?:aka\.ms/LearnAboutSenderIdentification[^\s>\]]*)[\s>\]]*", re.IGNORECASE)


def quitar_aviso_ms(t):
    """Treu l'avis de Microsoft del principi del cos.

    'No suele recibir correo electronico de X. Por que es esto importante
    <https://aka.ms/...>' va ENGANXAT a la pregunta del client i ocupa les
    primeres linies, que son les que mes pesen al classificador."""
    t = str(t or "")
    nou_t = RE_AVIS_MS.sub("", t.lstrip(), count=1).lstrip(" []>\n\t")
    return nou_t if nou_t.strip() else t


def cuerpo_nuevo(cuerpo, minimo=15):
    """Nomes el que ha escrit el client AQUESTA vegada, sense el fil citat.

    Si no es talla, un "muchas gracias" de dues linies es classifica pel
    contingut del correu anterior que queda citat a sota, i acaba al calaix
    equivocat. El que mana es l'ultim missatge."""
    # Fora l'avis de Microsoft: ocupa les primeres linies i despista
    # el classificador, que es el que mes pes te.
    t = quitar_aviso_ms(cuerpo)
    marques = ("\nEl ", "\n-----Original", "\nFrom:", "\nDe:", "\n________",
               "\n> ", "\nEnviado desde", "\nObtener Outlook", "\nSent:", "\nEnviado el:")
    tall = len(t)
    for m in marques:
        i = t.find(m)
        if 0 <= i < tall:
            tall = i
    nou_t = t[:tall].strip()
    return nou_t if len(nou_t) >= minimo else t.strip()


def datos_ya_facilitados(asunto, cuerpo):
    """Dades que el client JA ha escrit al missatge actual.

    Un model petit no segueix de manera fiable una regla que diu "no demanis el
    que ja t'han donat", sobretot si al fil citat nosaltres li vam demanar. Aixi
    que la dada s'extreu aqui, de manera determinista, i se li dona feta amb una
    prohibicio explicita. Es mira NOMES el missatge actual, no el fil citat:
    el que va escriure el client fa mesos no compta com a facilitat ara."""
    txt = (asunto or "") + "\n" + (cuerpo or "")
    # Tallar el fil citat: nomes interessa el que ha escrit ELL aquesta vegada
    for marca in ("\nEl ", "\n-----Original", "\nFrom:", "\nDe:", "\n________"):
        i = txt.find(marca)
        if i > 40:
            txt = txt[:i]
            break
    trobat = []
    m = REGEX_MATRICULA.search(txt.upper())
    if m:
        trobat.append(("matricula", m.group(0).replace(" ", "").upper()))
    m = RE_DNI.search(txt)
    if m:
        trobat.append(("DNI", m.group(0).upper()))
    m = RE_TEL.search(txt)
    if m:
        trobat.append(("telefono", m.group(0)))
    return trobat


CARPETA_ADJUNTOS = os.path.join(AQUI, "adjuntos")


def ruta_adjunto(nombre):
    """Ruta completa d'un document a adjuntar, o None si no hi es."""
    if not nombre:
        return None
    p = os.path.join(CARPETA_ADJUNTOS, str(nombre).strip())
    return p if os.path.isfile(p) else None


def comprobar_adjuntos(nombres):
    """(llista_de_rutes, llista_de_noms_que_FALTEN).

    Si en falta algun, el correu NO s'ha de redactar: hi hauria un text dient
    'le adjuntamos...' sense el document. Val mes que ho miri una persona."""
    rutas, falten = [], []
    for n in nombres or []:
        r = ruta_adjunto(n)
        (rutas if r else falten).append(r or n)
    return rutas, falten


def limpiar_adjuntos_heredados(reply):
    """Treu els adjunts que Outlook arrossega del fil citat.

    Quan es conserva el missatge original a sota, les imatges incrustades
    (image001.png, logos, firmes) viatgen a l'esborrany amb les referencies
    trencades i es veuen com a quadrats blancs. L'esborrany nomes ha de portar
    els documents que hi posem nosaltres."""
    trets = 0
    try:
        for i in range(reply.Attachments.Count, 0, -1):
            try:
                reply.Attachments.Item(i).Delete()
                trets += 1
            except Exception:
                pass
    except Exception:
        pass
    if trets:
        print(f"    (trets {trets} adjunts heretats del fil citat)")
    return trets


def adjuntar(reply, rutas):
    """Enganxa els documents a l'esborrany. Torna quants n'ha pogut posar."""
    posats = 0
    for r in rutas or []:
        try:
            reply.Attachments.Add(r)
            posats += 1
            print(f"    adjuntat: {os.path.basename(r)}")
        except Exception as e:
            print(f"    (no s'ha pogut adjuntar {os.path.basename(r)}: {str(e)[:60]})")
    return posats


RUTA_RESPUESTAS = os.path.join(AQUI, "respuestas.json")
RUTA_AFECTACIONES = os.path.join(AQUI, "afectaciones.csv")
ARBOL = {}
AFECTACIONES = {}


def cargar_arbol():
    """Carrega l'arbre de respostes i la taula d'afectacions."""
    global ARBOL, AFECTACIONES
    try:
        with open(RUTA_RESPUESTAS, encoding="utf-8") as fh:
            ARBOL = json.load(fh)
        print(f"    arbre de respostes: {len(ARBOL.get('por_estado', {}))} estats,"
              f" {len(ARBOL.get('por_intencion', {}))} intencions")
    except Exception as e:
        print(f"    (AVIS: respuestas.json no carregat: {str(e)[:70]})")
    try:
        import csv as _csv
        with open(RUTA_AFECTACIONES, encoding="utf-8") as fh:
            for r in _csv.DictReader(fh):
                AFECTACIONES[r["marca"].strip().upper()] = (r["inicio"], r["fin"])
        print(f"    taula d'afectacions: {len(AFECTACIONES)} marques")
    except Exception as e:
        print(f"    (AVIS: afectaciones.csv no carregat: {str(e)[:70]})")


def esta_afectado(marca, fecha_compra):
    """True/False/None (None = marca desconeguda o data desconeguda)."""
    m = (marca or "").strip().upper()
    if m not in AFECTACIONES or not fecha_compra:
        return None
    ini, fin = AFECTACIONES[m]
    return ini <= str(fecha_compra)[:10] <= fin


def calaix_para_humano(categoria):
    """A quin calaix ha d'anar un cas que no s'automatitza.

    Un desistiment ha d'anar a 3 DESISTIMIENTO, no a 2 DIFICIL: el treballador
    ha de trobar cada cosa al seu lloc."""
    m = ARBOL.get("calaix_humano", {})
    return m.get(categoria or "", m.get("_defecto", CALAIX_DIFICIL))


def codigo_estado(ficha):
    """Codi d'estat (E1..E6) de la fitxa.

    Si el camp 'estado_codigo' no hi es —perque la fitxa ve del triatge i no
    del conector— es dedueix de la frase de l'estat. Aixi l'arbre funciona
    vingui la fitxa d'on vingui."""
    f = ficha or {}
    cod = f.get("estado_codigo")
    if cod:
        return cod
    t = str(f.get("estado") or "").lower()
    if not t:
        return None
    if "no elegible" in t or "no apto" in t:
        return "E6"
    if "informe pericial" in t and ("recibido" in t or "completado" in t or "con informe" in t):
        return "E2"
    if "espera del informe" in t or "sin informe" in t:
        return "E1"
    if "diligencias" in t:
        if "sin el modelo 576" in t or "no solicitado" in t:
            return "E4"
        return "E5" if ("576 solicitado" in t or "con el modelo 576" in t) else "E4"
    if "pendiente de documentaci" in t or "falta" in t:
        return "E3.1"
    if "sin información de estado" in t or "sin informacion de estado" in t:
        return "E8"
    if "revisi" in t or "preparaci" in t or "tramitaci" in t:
        return "E3"
    return None


def arbol_como_referencia(ficha):
    """Tot l'arbre en text, perque el model triï ell mateix.

    Es dona el text de l'ESTAT del client (nomes el que li correspon) i tots
    els textos per tema. Si cap encaixa, el model respon amb logica amb la
    fitxa i l'historial, sense inventar dades."""
    if not ARBOL:
        return ""
    cod = codigo_estado(ficha)
    parts = []
    if cod:
        t = ARBOL.get("por_estado", {}).get(cod, "")
        if t and t not in ("__HUMANO__", "__NO_RESPONDER__"):
            if "{IMPORTE}" in t:
                imp = (ficha or {}).get("importe_perito")
                t = t.replace("{IMPORTE}", imp) if imp else ""
            if "{DOC_FALTA}" in t:
                falten = (ficha or {}).get("documentos_pendientes") or []
                dt = ARBOL.get("doc_textos", {})
                t = (t.replace("{DOC_FALTA}", "\n".join("- " + dt.get(str(d).lower(), d)
                                                        for d in falten)) if falten else "")
            if t:
                parts.append("TEXTO OFICIAL PARA EL ESTADO ACTUAL DE ESTE CLIENTE "
                             f"({cod}). Si el cliente pregunta por el estado de su "
                             "reclamacion, por plazos o por como va su caso, responde "
                             "CON ESTE TEXTO, tal cual:\n" + t)
    otros = []
    NOM = {"I2": "si pregunta cuando cobrara o por que tarda",
           "I3": "si pregunta cuanto dinero recuperara",
           "I4": "si pregunta que documentacion hace falta",
           "I5": "si dice que ya ha enviado documentacion",
           "I6": "si quiere desistir o cancelar",
           "I9": "si pregunta por la comision u honorarios",
           "I11": "si pide que le llamen por telefono",
           "I13": "si pregunta por la prescripcion o si aun esta a tiempo"}
    for k, desc in NOM.items():
        t = ARBOL.get("por_intencion", {}).get(k, "")
        if t and t not in ("__HUMANO__", "__NO_RESPONDER__") and "{DOC_FALTA}" not in t:
            otros.append(f"· {desc}:\n  {t}")
    if otros:
        parts.append("TEXTOS OFICIALES POR TEMA (usa el que encaje, tal cual):\n"
                     + "\n".join(otros))
    parts.append(
        "SI NINGUNO DE LOS TEXTOS ANTERIORES RESPONDE A LO QUE PREGUNTA EL CLIENTE: "
        "no fuerces ninguno. Responde de forma logica y breve a lo que pregunta, "
        "usando los DATOS VERIFICADOS y el historial. No inventes importes, plazos, "
        "fechas ni documentos. Si no tienes el dato, di que un companero se lo "
        "confirmara. No pongas saludo ni despedida: se anaden aparte.")
    return "\n\n".join(parts)


def texto_del_arbol(categoria, ficha):
    """Text OBLIGATORI segons l'arbre: (texto, adjuntos, accion).

    accion: "texto" (hi ha text), "humano" (no automatitzar),
            "no_responder" (marcar llegit i prou), "" (no hi ha regla)."""
    if not ARBOL:
        return "", [], ""
    inten = ARBOL.get("categoria_a_intencion", {}).get(categoria or "")
    if not inten:
        # No hi ha regla per a aquesta categoria. Si sabem l'estat del client,
        # es fa servir el text de l'estat com a base (el model l'adapta a la
        # pregunta). Nomes si no sabem res, ho mira una persona.
        cod = codigo_estado(ficha)
        if cod and ARBOL.get("fallback_estado"):
            base = ARBOL.get("por_estado", {}).get(cod, "")
            if base and base not in ("__HUMANO__", "__NO_RESPONDER__"):
                if "{IMPORTE}" in base:
                    imp = (ficha or {}).get("importe_perito")
                    if not imp:
                        return "", [], "humano"
                    base = base.replace("{IMPORTE}", imp)
                if "{DOC_FALTA}" in base:
                    base = ARBOL.get("por_estado", {}).get("E3", "")
                    if not base:
                        return "", [], "humano"
                # Abans es tornava el text de l'estat com a obligatori. Pero si el
                # client pregunta una altra cosa (apoderaments, certificat digital),
                # aquell text no respon res. Millor donar-li l'arbre sencer i que
                # triï: si cap encaixa, respon amb logica.
                return "", [], ""
        return "", [], ""
    adj = []
    # I1 (estat de la reclamacio) depen de l'ESTAT de l'expedient
    if inten == "I1":
        cod = codigo_estado(ficha) or "E7"
        print(f"    ARBRE: intencio {inten} | estat {cod}"
              + (" | import: " + str((ficha or {}).get("importe_perito"))
                 if cod == "E2" else ""))
        txt = ARBOL.get("por_estado", {}).get(cod, "")
        adj = list(ARBOL.get("adjuntos_estado", {}).get(cod, []))
    else:
        txt = ARBOL.get("por_intencion", {}).get(inten, "")
    if txt == "__HUMANO__":
        print(f"    ARBRE: aquest estat/intencio esta marcat com a HUMA a respuestas.json")
        return "", [], "humano"
    if txt == "__NO_RESPONDER__":
        return "", [], "no_responder"
    if not txt:
        return "", [], ""
    # {DOC_FALTA}: llista de documents pendents segons la fitxa
    if "{DOC_FALTA}" in txt:
        falten = (ficha or {}).get("documentos_pendientes") or []
        dt = ARBOL.get("doc_textos", {})
        if not falten:
            # No sabem quin document falta: en lloc d'enviar-ho a una persona,
            # es dona la resposta general de "expedient en revisio".
            gen = ARBOL.get("por_estado", {}).get("E3", "")
            return (gen, [], "texto") if gen else ("", [], "humano")
        txt = txt.replace("{DOC_FALTA}",
                          "\n".join("- " + dt.get(d.lower(), d) for d in falten))
        # El PDF que toca segons el document que falta. Es mira el mapa GLOBAL,
        # aixi funciona vingui el cas per la intencio o per l'estat: abans nomes
        # s'adjuntava si entrava per I4, i el 576 es mencionava sense adjuntar-lo.
        mapa = dict(ARBOL.get("adjuntos_documento", {}))
        mapa.update(ARBOL.get("adjuntos_intencion", {}).get(inten, {}))
        for d in falten:
            a = mapa.get(str(d).strip().lower())
            if a and a not in adj:
                adj.append(a)
    # {IMPORTE}: si no el tenim, no s'automatitza
    if "{IMPORTE}" in txt:
        imp = (ficha or {}).get("importe_perito")
        if not imp:
            return "", [], "humano"
        txt = txt.replace("{IMPORTE}", imp)
    return txt, adj, "texto"


def destinatario_real(msg, rem, reenviadores, destino_fijo=""):
    """A qui ha d'anar l'esborrany.

    Si el correu ens l'ha REENVIAT un company, msg.Reply() respondria al company,
    no al client. Aqui es posa el client que ja s'ha resolt (extret del cos del
    reenviament). Si no es pot determinar, es deixa el que hi posi Outlook i
    s'avisa: val mes que ho miri una persona que enviar-ho a qui no toca."""
    if destino_fijo:
        return destino_fijo, "destinatari fix de l'escenario"
    r = (rem or "").strip().lower()
    if "@" not in r or r == "desconocido":
        return "", "no s'ha pogut determinar el client"
    if es_reenviador(r, reenviadores):
        return "", "el client resolt es una adreca interna"
    try:
        actual = str(getattr(msg, "SenderEmailAddress", "") or "").lower()
    except Exception:
        actual = ""
    if es_reenviador(remitente_smtp(msg) or actual, reenviadores):
        return r, "REENVIAMENT: s'adreca al client, no al company"
    return r, ""


def remitente_smtp(msg):
    """Adreca SMTP REAL del remitent, o None.

    Els mails interns d'Exchange tenen un SenderEmailAddress del tipus
    /o=ExchangeLabs/ou=.../cn=... que NO es un email. Si no es resol, el mail
    no sembla un reenviament, no s'extreu el client del cos, i l'agrupacio no
    pot funcionar. Es prova: usuari d'Exchange -> PR_SENDER_SMTP_ADDRESS ->
    PR_SMTP_ADDRESS -> el camp directe."""
    try:
        if str(getattr(msg, "SenderEmailType", "") or "") == "EX":
            try:
                ex = msg.Sender.GetExchangeUser()
                if ex and ex.PrimarySmtpAddress and "@" in str(ex.PrimarySmtpAddress):
                    return str(ex.PrimarySmtpAddress).lower()
            except Exception:
                pass
    except Exception:
        pass
    for prop in ("http://schemas.microsoft.com/mapi/proptag/0x5D01001F",   # PR_SENDER_SMTP_ADDRESS
                 "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"):  # PR_SMTP_ADDRESS
        try:
            v = str(msg.PropertyAccessor.GetProperty(prop) or "").strip().lower()
            if "@" in v:
                return v
        except Exception:
            pass
    try:
        v = str(getattr(msg, "SenderEmailAddress", "") or "").strip().lower()
        if "@" in v:
            return v
    except Exception:
        pass
    return None


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


def formatear_respuesta(texto, categoria=""):
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
    # DESISTIMENT: mai "Quedamos a la espera.". El client vol tancar; quedar a
    # l'espera sona a que se li demana alguna cosa mes o a que no s'accepta.
    es_desist = "desist" in str(categoria).lower() or "cancelacion" in str(categoria).lower()
    if es_desist:
        t = _re.sub(r"\s*Quedamos a la espera\.?\s*$", "", t, flags=_re.IGNORECASE)
        t = _re.sub(r"\n\s*Quedamos a la espera\.?\s*\n", "\n", t, flags=_re.IGNORECASE)
    # MAI els dos tancaments alhora: si hi ha "Saludos cordiales" al final,
    # qualsevol "Quedamos a la espera" anterior sobra (i a l'inreves).
    if _re.search(r"saludos cordiales\s*[.,]?\s*$", t, _re.IGNORECASE):
        t = _re.sub(r"\n[ \t]*Quedamos a la espera\.?[ \t]*(?=\n)", "", t, flags=_re.IGNORECASE)
        t = _re.sub(r"\n[ \t]*Quedamos a su disposici[oó]n[^\n]*(?=\n)", "", t, flags=_re.IGNORECASE)
    # frases buides de farciment que l'equip no vol
    t = _re.sub(r"[^.\n]*\bagradecemos su paciencia\b[^.\n]*\.\s*", "", t, flags=_re.IGNORECASE)
    t = _re.sub(r"\n{3,}", "\n\n", t)
    # tancament garantit: peticio -> "Quedamos a la espera." | informatiu -> "Saludos cordiales,"
    if not _re.search(r"(quedamos a la espera|saludos cordiales)\s*[.,]?\s*$", t, _re.IGNORECASE):
        pide = _re.search(r"(env[ií]e|env[ií]enos|nos mande|m[aá]ndenos|facil[ií]t|adjunte|reenv[ií]e|"
                          r"necesitamos que|puede indicarnos|puede facilitarnos|\?)", t, _re.IGNORECASE)
        if es_desist:
            pide = None                      # en desistiment, sempre comiat net
        t = t.rstrip() + "\n\n" + ("Quedamos a la espera." if pide else "Saludos cordiales,")
    # Fora les peticions de dades que ja tenim a la BBDD (any de matriculacio,
    # marca, model, bastidor). Es feina nostra, no del client.
    t = _re.sub(r"[^.\n]*\b(a[nñ]o de matriculaci[oó]n|fecha de matriculaci[oó]n|"
                r"a[nñ]o del veh[ií]culo|n[uú]mero de bastidor|"
                r"marca y modelo del veh[ií]culo)\b[^.\n]*\.\s*", "", t, flags=_re.IGNORECASE)
    t = _re.sub(r"\n{3,}", "\n\n", t)
    # SALUTACIO SENSE NOM: "Buenos dias Leonor Cruz Gimenez," -> "Buenos dias,".
    # El nom sencer amb cognoms queda fred i impersonal; el saludo net es mes
    # natural i no pot equivocar-se de persona.
    # "Estimado/Apreciado/Querido" necessiten un substantiu al darrere: sense nom
    # queden coixos ("Estimado,"), aixi que es converteixen en "Buenos dias,".
    t = _re.sub(r"^(Estimad[oa]s?|Apreciad[oa]s?|Querid[oa]s?)[ \t]+[^\n,:]{2,60}?[,:]",
                "Buenos días,", t, count=1, flags=_re.IGNORECASE)
    t = _re.sub(r"^(Estimad[oa]s?|Apreciad[oa]s?|Querid[oa]s?)[ \t]*[,:.]",
                "Buenos días,", t, count=1, flags=_re.IGNORECASE)
    t = _re.sub(r"^(Buenos d[ií]as|Buenas tardes|Buenas noches|Hola)[ \t]+[^\n,:.]{2,60}?([,:.])",
                r"\1\2", t, count=1, flags=_re.IGNORECASE)
    # restes del model: una linia solta amb "cordial," o "atentamente" enmig
    t = _re.sub(r"\n[ \t]*(cordial(mente)?|atentamente|un cordial saludo)[,.]?[ \t]*(?=\n)",
                "", t, flags=_re.IGNORECASE)
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


def _componentes_de(v):
    """Treu (any, mes, dia, hora, minut, segon) d'un valor de data d'Outlook."""
    if v is None:
        return None
    try:                      # cami normal: l'objecte porta els components
        anyo, mes, dia = int(v.year), int(v.month), int(v.day)
        if 1990 <= anyo <= 2100 and 1 <= mes <= 12 and 1 <= dia <= 31:
            return (anyo, mes, dia, int(getattr(v, "hour", 0) or 0),
                    int(getattr(v, "minute", 0) or 0), int(getattr(v, "second", 0) or 0))
    except Exception:
        pass
    t = str(v)
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?", t)
    if m:                     # ISO: 2026-09-15 10:04:00
        return tuple(int(g or 0) for g in m.groups())
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})\D{0,3}(\d{1,2})?:?(\d{2})?:?(\d{2})?", t)
    if m:                     # local: 15/09/2026 10:04
        dia, mes, anyo = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if mes > 12:          # venia en format mm/dd
            dia, mes = mes, dia
        if 1 <= mes <= 12 and 1 <= dia <= 31:
            return (anyo, mes, dia, int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0))
    return None


def componentes_fecha(msg):
    """Data REAL del mail com a tupla (any, mes, dia, hora, minut, segon).

    Comparar aquesta tupla compara PRIMER l'any, despres el mes, despres el dia,
    despres l'hora... que es l'unic ordre correcte. Mai s'ordena per text: en
    format local '31/08/2026' surt DESPRES de '01/09/2026' alfabeticament, i per
    aixo el 'mes nou' acabava sent el mes vell.
    Torna None si cap font dona una data creible (llavors NO s'agrupa)."""
    for origen in ("ReceivedTime", "SentOn", "CreationTime"):
        t = _componentes_de(getattr(msg, origen, None))
        if t:
            return t
    try:
        return _componentes_de(msg.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x0E060040"))
    except Exception:
        return None


def fecha_texto(t):
    if not t:
        return "sense data"
    return "%04d-%02d-%02d %02d:%02d:%02d" % t


def procesar_carpeta(nombre_buzon, nombre_carpeta, esc, con, bbdd, dry, max_mails,
                     solo_clasificar=False, ruta_csv=""):
    global PLANTILLAS
    import win32com.client
    ia = esc["ia"]
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = escoger_buzon(stores, nombre_buzon)
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
    # Adreca del propi buzon: serveix per no confondre les NOSTRES respostes
    # dins d'un fil amb mails nous del client.
    buzon_propio = ""
    try:
        for cuenta_i in range(ns.Accounts.Count):
            cta = ns.Accounts.Item(cuenta_i + 1)
            if nombre_buzon.lower() in str(cta.DisplayName or "").lower() or \
               nombre_buzon.lower() in str(cta.SmtpAddress or "").lower():
                buzon_propio = str(cta.SmtpAddress or "").lower()
                break
    except Exception:
        pass
    if buzon_propio:
        print(f"    (adreca del buzon: {buzon_propio})")
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
    # Lectura MANDROSA: Outlook ordena i nomes s'agafen els que calen. Amb
    # milers de correus, carregar-los tots tomba l'Outlook per falta de memoria.
    orden = str(esc.get("orden_tanda", "antiguos")).lower()
    dias = int(esc.get("dias_max", 0) or 0)
    # Es demana un marge sobre el topall: molts mails es descarten pel registre
    # o per la marca, i sense marge la tanda es quedaria curta.
    margen = max(int(max_mails) * 3, int(max_mails) + 40)
    lista_items = mails_de_carpeta_items(items, orden, margen, dias)
    print(f"    ordre de la tanda: es comenca pels mails mes "
          f"{'NOUS' if orden in ('recientes','nuevos','desc') else 'VELLS'}")

    # AGRUPACIO PER CLIENT: si un client te diversos mails pendents, es respon
    # NOMES el mes recent (amb els anteriors com a context); la resta es marca.
    reenv_pre = esc.get("reenviadores", [])
    def _clave_cliente(m):
        """Qui es el client d'aquest mail (mateixa resolucio que el bucle
        principal: si no, l'agrupacio i el redactor parlarien de clients
        diferents — que es exactament el que passava)."""
        try:
            snd = remitente_smtp(m)
            if not snd:
                return None
            if es_reenviador(snd, reenv_pre):
                return extraer_cliente_de_reenvio(
                    str(getattr(m, "Body", "") or "")[:4000], internos=reenv_pre)
            return snd
        except Exception:
            return None

    def _clave_conversacion(m):
        """El fil d'Outlook. Serveix de segona xarxa: si el client no s'identifica
        igual als dos mails (reenviaments, adreces diferents), el fil els uneix."""
        try:
            v = getattr(m, "ConversationID", None)
            if v:
                return "conv:" + str(v)
        except Exception:
            pass
        try:
            return None   # NOMES el fil real. L'assumpte NO serveix de pont:
            # dos clients amb l'assumpte "Consulta" no son la mateixa persona.
        except Exception:
            return None

    # Agrupacio per DUES claus alhora (client i fil): si dos mails comparteixen
    # qualsevol de les dues, van al mateix grup. Union-find simple.
    padre = {}
    def _arrel(x):
        while padre.get(x, x) != x:
            padre[x] = padre.get(padre[x], padre[x])
            x = padre[x]
        return x
    def _unir(a, b):
        ra, rb = _arrel(a), _arrel(b)
        if ra != rb:
            padre[rb] = ra

    candidatos = []
    for m in lista_items:
        if te_marca_agente(m):
            continue
        claves = [k for k in (_clave_cliente(m), _clave_conversacion(m)) if k]
        if not claves:
            continue
        for k in claves:
            padre.setdefault(k, k)
        for k in claves[1:]:
            _unir(claves[0], k)
        candidatos.append((m, claves))

    grupos_brutos = {}
    for m, claves in candidatos:
        grupos_brutos.setdefault(_arrel(claves[0]), []).append(m)

    # GUARDA ANTI-BARREJA: un pack no pot contenir dos clients identificats
    # diferents. Si el fil ha unit persones distintes (reenviaments encadenats,
    # fils compartits), es desfa la unio i es torna a agrupar nomes per client.
    # Val mes perdre una agrupacio que deixar un client sense resposta.
    grupos = {}
    for raiz, ms in grupos_brutos.items():
        clientes = {c for c in (_clave_cliente(m) for m in ms) if c}
        if len(clientes) <= 1:
            grupos[raiz] = ms
            continue
        print(f"    (AVIS: el fil unia {len(clientes)} clients diferents"
              f" — es desfa i s'agrupa nomes per client)")
        for m in ms:
            c = _clave_cliente(m)
            if c:
                grupos.setdefault("cli:" + c, []).append(m)
            # sense client identificat dins d'un fil barrejat: va sol
    agrupados_omitir = {}   # mail_id -> mail_id del guanyador
    contexto_grupo = {}     # mail_id guanyador -> [(fecha, asunto, extracte), ...]
    for k, ms in grupos.items():
        if len(ms) < 2:
            continue
        fechados = [(componentes_fecha(m), m) for m in ms]
        if any(f is None for f, _ in fechados):
            print(f"    (AVIS: no puc datar tots els mails de {k} — NO s'agrupen;"
                  " cada mail rep el seu esborrany)")
            continue
        # Ordre REAL: compara any -> mes -> dia -> hora -> minut -> segon
        fechados.sort(key=lambda p: p[0])
        ganador, hermanos = fechados[-1][1], [m for _, m in fechados[:-1]]
        gid = id_estable(ganador)
        contexto_grupo[gid] = [(fecha_texto(f),
                                str(getattr(h, "Subject", "") or "")[:60],
                                str(getattr(h, "Body", "") or "")[:300].replace("\r\n", " ").replace("\n", " "))
                               for f, h in fechados[:-1]]
        for h in hermanos:
            agrupados_omitir[id_estable(h)] = gid
        print(f"    PACK {k}: {len(ms)} mails (ordenats per any/mes/dia/hora)")
        for f, m in fechados[:-1]:
            print(f"       previ     {fecha_texto(f)}  {str(getattr(m, 'Subject', '') or '')[:38]}")
        print(f"       RESPON -> {fecha_texto(fechados[-1][0])}"
              f"  {str(getattr(ganador, 'Subject', '') or '')[:38]}   <= EL MES NOU")
    if agrupados_omitir:
        print(f"    (agrupacio: {len(contexto_grupo)} clients amb multiples mails;"
              f" {len(agrupados_omitir)} mails es contesten dins del mes recent)")

    n = creados = 0
    omesos_registre = 0
    ya_con_borrador = set()   # clients/fils que JA tenen esborrany en aquesta passada
    # Si l'equip vol veure el comptador de pendents, els mails tractats NO es
    # marquen com a llegits (l'etiqueta 'Agente' hi va igualment).
    MARCAR_LEIDO = str(esc.get("marcar_leido", "si")).lower() not in ("no", "false", "off", "0")
    filas_csv = []            # una fila per mail, per al CSV de sortida
    motivos_sin = {}          # per que NO s'ha redactat cada mail

    def _apunta(motivo):
        motivos_sin[motivo] = motivos_sin.get(motivo, 0) + 1
    for msg in lista_items:
        try:
            if msg.Class != 43:
                continue
            mail_id = id_estable(msg)
            ya_registrado = bool(con.execute(
                "SELECT 1 FROM borradores WHERE mail_id=?", (mail_id,)).fetchone())
            if ya_registrado and mail_id not in agrupados_omitir:
                omesos_registre += 1
                continue
            # Un germa d'un pack NO se salta pel registre: encara que ja hi
            # constes, s'ha d'apartar a MULTIPLES. Si no, el mail es queda per
            # sempre al calaix i la carpeta no es crea mai.
            n += 1
            if n > max_mails:
                break
            if mail_id in agrupados_omitir:
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                _apunta("agrupat: es contesta al mail mes recent")
                print("    AGRUPAT: es contesta dins del mail mes recent del mateix client")
                movido_mult = False
                if tiene_adjuntos_reales(msg):
                    print("    PORTA ARXIUS: NO s'aparta ni es marca —"
                          " ha de quedar visible per a una persona")
                    print("    RESULTADO -> ESCALADO (adjunto en mail agrupado)")
                    con.commit()
                    continue
                if ya_registrado:
                    print("    (ja constava al registre: igualment s'aparta)")
                if not dry:
                    marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                    # El germa no es queda al calaix: s'aparta a MULTIPLES, aixi
                    # el calaix nomes conte el mail que s'ha contestat.
                    movido_mult = mover_a_multiples(msg, carpeta)
                else:
                    print(f"    (DRY: no es mou res; en real aniria a {CARPETA_MULTIPLES})")
                    if movido_mult:
                        print(f"    -> apartat a la carpeta {CARPETA_MULTIPLES}")
                con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (datetime.now(timezone.utc).isoformat(), mail_id, "", str(msg.Subject or ""),
                             "", "", "", "", "",
                             "AGRUPADO -> " + CARPETA_MULTIPLES if movido_mult
                             else "AGRUPADO (contestado en el mail mas reciente del cliente)", ""))
                con.commit()
                continue
            # LECTURA DEL FIL: si aquest mail ja te una continuacio mes nova del
            # client, no es contesta — s'aparta i la resposta anira al mes recent.
            # Aixo es el que fa viable anar dels mails vells cap als nous sense
            # escanejar la bustia sencera.
            mirar = str(esc.get("mirar_hilo", "si")).lower() not in ("no", "false", "off", "0")
            if mirar and not buzon_propio:
                # Sense saber la nostra adreca no podem distingir una resposta
                # NOSTRA dins del fil d'un mail nou del client: tot semblaria
                # "ja contestat" i no es redactaria res. Millor no mirar el fil.
                if not globals().get("_aviso_hilo_dado"):
                    print("    (lectura del fil DESACTIVADA: no s'ha pogut determinar"
                          " l'adreca del buzon)")
                    globals()["_aviso_hilo_dado"] = True
                mirar = False
            if mirar:
                nuevo = mas_nuevo_en_hilo(msg, buzon_propio)
                if nuevo is not None:
                    print(f"\n=== {str(msg.Subject or '')[:50]}")
                    _apunta("fil: hi ha un mail mes nou del client")
                    print("    FIL: el client ja ha escrit mes tard en aquesta conversa"
                          f" ({fecha_texto(componentes_fecha(nuevo))})"
                          " — es contestara al mail mes recent")
                    movido_h = False
                    if tiene_adjuntos_reales(msg):
                        print("    PORTA ARXIUS: NO s'aparta ni es marca —"
                              " ha de quedar visible per a una persona")
                    elif not dry:
                        marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                        movido_h = mover_a_multiples(msg, carpeta)
                        if movido_h:
                            print(f"    -> apartat a la carpeta {CARPETA_MULTIPLES}")
                    con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                (datetime.now(timezone.utc).isoformat(), mail_id, "", str(msg.Subject or ""),
                                 "", "", "", "", "",
                                 "HILO (existe un mail mas reciente del cliente)", ""))
                    con.commit()
                    continue
            claves_msg = [k for k in (_clave_cliente(msg), _clave_conversacion(msg)) if k]
            if any(k in ya_con_borrador for k in claves_msg):
                # XARXA FINAL: aquest client ja te esborrany en aquesta passada.
                # Com que es treballa del mes nou al mes vell, el que ja te
                # esborrany es SEMPRE el mes recent: aquest es un germa antic.
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                _apunta("agrupat: el client ja te esborrany")
                print("    AGRUPAT (xarxa): el client ja te esborrany en aquesta passada"
                      " — es contesta al mail mes recent")
                if not dry:
                    marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                            " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (datetime.now(timezone.utc).isoformat(), mail_id, "", str(msg.Subject or ""),
                             "", "", "", "", "", "AGRUPADO (un solo borrador por cliente y pasada)", ""))
                con.commit()
                continue
            if te_marca_agente(msg):
                print(f"\n=== {str(msg.Subject or '')[:50]}")
                _apunta("ja portava la marca Agente")
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
            rem = remitente_smtp(msg) or "desconocido"
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
            if mail_id in contexto_grupo:
                avisos += f" [AGRUPAT: respon tambe {len(contexto_grupo[mail_id])} mails previs del client]"

            # ESTRATEGIA: BORRADOR PER A TOT, excepte adjunts, sistema i DESCARTES.
            dominios = DOMINIOS_DESCARTES_BASE + [d.strip().lower().lstrip("@")
                                                  for d in esc.get("dominios_descartes", [])]
            if es_dominio_descartes(rem, dominios):
                destino, motivo = "DESCARTE", "dominio_no_cliente"
            mover_a_dificil = False
            sin_importe = False
            docs_off = str(esc.get("documentos_sin_borrador", "no")).lower() not in ("no", "false", "off", "0")
            sin_borrador = (destino in ("SISTEMA", "DESCARTE")) or (motivo == "adjunto_real")
            # Client NO identificat: no te sentit redactar. L'unic que es pot dir
            # es una frase generica demanant la matricula, que sovint el client
            # JA ha escrit. Millor que ho vegi una persona.
            if str(esc.get("borrador_sin_ficha", "no")).lower() in ("no", "false", "off", "0"):
                if destino == "HUMANO" and motivo in ("sin_ficha_bbdd", "remitente_sin_historial"):
                    sin_borrador = True
                    mover_a_dificil = True
                    print(f"    SENSE ESBORRANY: client no identificat ({motivo})"
                          " — ho mira una persona")
            motivo_docs = ""
            if docs_off and not sin_borrador:
                # Porta 1: parla d'arxius encara que no en porti cap
                # Nomes senyal fort: el combo verb+nom bloquejava preguntes
                # d'estat normals ("ya os envie la documentacion, como va?").
                habla, det = habla_de_documentos(asunto, cuerpo, solo_fuerte=True)
                if habla:
                    sin_borrador, motivo_docs = True, det
            # SEGURETAT: si el client menciona una matricula i la fitxa trobada
            # es d'UNA ALTRA, no es respon. Podria ser una confusio de dades a la
            # BBDD o dues persones relacionades; en tot cas, dir-li l'estat d'un
            # expedient que no es el seu seria greu.
            try:
                _m = REGEX_MATRICULA.search((asunto + " " + cuerpo_nuevo(cuerpo)).upper())
                _mf = str((ficha or {}).get("matricula") or "").replace(" ", "").upper()
                if _m and _mf:
                    _mm = _m.group(0).replace(" ", "").upper()
                    if _mm != _mf:
                        print(f"    ATENCIO: el client parla de {_mm} pero la fitxa"
                              f" trobada es de {_mf} — ho mira una persona")
                        sin_borrador = True
                        mover_a_dificil = True
            except Exception:
                pass
            # Si la fitxa diu que l'informe pericial esta fet pero NO tenim l'import,
            # la resposta que toca (plantilla E2) no es pot omplir. Abans d'inventar
            # una xifra o deixar un buit, ho mira una persona.
            try:
                _f = ficha or {}
                if _f and _f.get("estado", "").startswith("informe pericial completado") \
                   and not _f.get("importe_perito"):
                    sin_borrador = True
                    mover_a_dificil = True
                    sin_importe = True
                    print("    SENSE ESBORRANY: informe pericial fet pero SENSE import a la"
                          " BBDD (total_perito buit) — ho mira una persona")
            except Exception:
                pass
            movido_a_otro = False
            if mover_a_dificil and not dry and \
               str(esc.get("sin_ficha_a_dificil", "si")).lower() not in ("no", "false", "off", "0"):
                if mover_a_calaix(msg, carpeta, CALAIX_DIFICIL):
                    movido_a_otro = True
                    print(f"    -> mogut a {CALAIX_DIFICIL} (SENSE marcar:"
                          " es podra tractar alla)")
            if movido_a_otro:
                # Ja s'ha mogut a un altre calaix: NO es marca ni s'apunta al
                # registre, perque alla el sistema l'ha de poder tractar.
                print("    (no es marca ni s'apunta: es tractara al calaix nou)")
                continue
            if sin_borrador and motivo_docs:
                _apunta("parla d'arxius")
                resultado = f"DOCUMENTOS ({motivo_docs}) — SENSE esborrany (ho mira una persona)"
                print(f"    SENSE ESBORRANY: el mail parla d'arxius ({motivo_docs})")
            elif sin_borrador:
                _apunta(f"triaje: {motivo}")
                resultado = destino + f" ({motivo}) — SENSE esborrany (per disseny)"
            else:
                # Historial real del buzon amb aquest client (abans de classificar: el fa servir)
                hilo_txt = buscar_historial(inbox, enviados, rem) or "(sin historial en el buzon)"
                if mail_id in contexto_grupo:
                    lineas_g = "\n".join(f"[{f}] ({a}): {c}" for f, a, c in contexto_grupo[mail_id])
                    hilo_txt += ("\nOTROS MAILS RECIENTES DEL MISMO CLIENTE, AUN SIN RESPONDER "
                                 "(esta respuesta debe atenderlos TODOS en un solo correo):\n" + lineas_g)
                    print(f"    agrupats: aquest esborrany respon tambe {len(contexto_grupo[mail_id])} mails previs del client")
                if "(sin historial" not in hilo_txt:
                    print(f"    historial del buzon: {hilo_txt.count(chr(10)) + 1} missatges previs trobats")
                # Classificar sempre (amb o sense fitxa): la categoria tria la plantilla
                # El cos SENCER hi es (fa falta per al context), pero separat:
                # el que decideix la categoria es el que ha escrit ARA el client,
                # no el correu nostre que queda citat a sota.
                cuerpo_ult = cuerpo_nuevo(cuerpo)
                resto = str(cuerpo or "")[len(cuerpo_ult):].strip()
                bloque = "LO QUE ACABA DE ESCRIBIR EL CLIENTE (esto es lo que hay que responder):\n" + cuerpo_ult
                if resto:
                    bloque += ("\n\nMENSAJES ANTERIORES DEL HILO:\n" + resto[:2500] +
                               "\n\nCLASIFICA por lo que acaba de escribir el cliente. PERO si su"
                               " mensaje se apoya en algo anterior ('como le dije', 'en el mensaje"
                               " anterior', 'segun le comente', 'ya le pregunte'), esa peticion"
                               " SIGUE PENDIENTE y forma parte de lo que hay que responder:"
                               " clasifica por ella.")
                    print(f"    (pesa l'ultim missatge: {len(cuerpo_ult)} car. nous"
                          f" + {len(resto)} de context)")
                c = llamar(ia, PROMPT_CLASIFICADOR, f"HISTORIAL PREVIO con este cliente:\n{hilo_txt}\n\nMENSAJE ACTUAL:\nAsunto: {asunto}\n{bloque}", rapido=True)
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
                if solo_clasificar:
                    filas_csv.append({
                        "remitente": rem, "asunto": asunto,
                        "pregunta": cuerpo_nuevo(cuerpo)[:900].replace("\n", " "),
                        "categoria": categoria, "flags": flags,
                        "estado": (ficha or {}).get("estado", ""),
                        "estado_codigo": codigo_estado(ficha) or "",
                        "matricula": (ficha or {}).get("matricula", ""),
                        "importe": (ficha or {}).get("importe_perito", ""),
                        "decision": "(solo clasificar)", "respuesta": ""})
                    print(f"    RESULTADO -> nomes classificat: {categoria}")
                    con.commit()
                    continue
                cats_doc = CATEGORIAS_DOCUMENTOS | {c_.strip().lower() for c_ in
                                                    esc.get("categorias_documentos", [])}
                if False and docs_off and categoria in cats_doc:
                    # Porta 2: el classificador diu que el tema SON els documents.
                    resultado = f"DOCUMENTOS (categoria {categoria}) — SENSE esborrany (ho mira una persona)"
                    _apunta(f"categoria de documents ({categoria})")
                    print(f"    SENSE ESBORRANY: categoria de documents ({categoria})")
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
                plantillas_on = str(esc.get("plantillas", "no")).lower() not in ("no", "off", "false", "0")
                pl = PLANTILLAS.get(categoria, {}) if plantillas_on else {}
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
                  txt_arbol, adj_arbol, accion_arbol = texto_del_arbol(categoria, ficha)
                  # El text literal NOMES si la pregunta es clarament d'estat.
                  # Si el client pregunta una altra cosa, forçar-li el text de
                  # l'estat no respon res: millor donar-li l'arbre al model.
                  if accion_arbol == "texto" and \
                     str(esc.get("literal_solo_estado", "si")).lower() not in ("no","false","off","0") \
                     and categoria not in ("estado_reclamacion", "coste_comision", "cortesia_breve"):
                      print(f"    ARBRE: '{categoria}' no es una pregunta d'estat —"
                            " el model decideix amb l'arbre davant")
                      txt_arbol, accion_arbol = "", ""
                  if accion_arbol == "" and txt_arbol == "":
                      ref = arbol_como_referencia(ficha)
                      if ref:
                          txt_arbol, accion_arbol = ref, "referencia"
                          print("    ARBRE: cap text exacte — es dona l'arbre sencer"
                                " al model perque triï")
                  if accion_arbol == "" and str(esc.get("solo_arbol", "no")).lower() \
                          not in ("no", "false", "off", "0"):
                      # No hi ha regla a l'arbre per a aquest cas. Abans que el model
                      # s'inventi una resposta, ho mira una persona.
                      print(f"    ARBRE: cap regla per a '{categoria}' — sense esborrany")
                      accion_arbol = "sin_regla"
                  if accion_arbol == "sin_regla":
                      # No hi ha regla: NO es redacta res, pero el correu es queda
                      # AL SEU CALAIX. Moure'l a DIFICIL nomes perque encara no
                      # tenim la regla escrita desordena la safata sense motiu.
                      resultado = f"ARBOL -> sin regla para '{categoria}' (sin borrador)"
                      print(f"    RESULTADO -> {resultado}")
                      if not dry:
                          marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                      con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                  " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                  (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto,
                                   categoria, flags, "", "", "", resultado, ""))
                      con.commit(); continue
                  if accion_arbol == "humano":
                      print("    ARBRE: aquest cas no s'automatitza — ho mira una persona")
                      resultado = "ARBOL -> HUMANO (sin borrador)"
                      print(f"    RESULTADO -> {resultado}")
                      _cal = calaix_para_humano(categoria)
                      _mogut = False
                      if not dry:
                          _mogut = mover_a_calaix(msg, carpeta, _cal)
                          if _mogut:
                              print(f"    -> mogut a {_cal} (SENSE marcar: es podra"
                                    " tractar alla)")
                          else:
                              marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                      # Si s'ha mogut, NO es marca ni s'apunta al registre: si no,
                      # al nou calaix el sistema el saltaria per "ja tractat".
                      if not _mogut:
                          con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                      " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                      (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto,
                                       categoria, flags, "", "", "", resultado, ""))
                      con.commit(); continue
                  if accion_arbol == "no_responder":
                      print("    ARBRE: no s'ha de respondre — es marca com a llegit")
                      resultado = "ARBOL -> NO RESPONDER (marcado leido)"
                      print(f"    RESULTADO -> {resultado}")
                      if not dry:
                          marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                      con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                  " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                  (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto,
                                   categoria, flags, "", "", "", resultado, ""))
                      con.commit(); continue
                  # Els documents promesos han d'existir de veritat
                  rutas_adj, falten_adj = comprobar_adjuntos(adj_arbol)
                  if falten_adj:
                      print(f"    ARBRE: falten documents al servidor ({', '.join(falten_adj)})"
                            " — ho mira una persona")
                      resultado = "ARBOL -> falta adjunto en el servidor"
                      print(f"    RESULTADO -> {resultado}")
                      if not dry:
                          marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                          mover_a_calaix(msg, carpeta, CALAIX_DIFICIL)
                      con.commit(); continue
                  arbol_txt = ""
                  if txt_arbol:
                      print(f"    ARBRE: text fixat ({len(txt_arbol)} car."
                            + (f", {len(rutas_adj)} adjunt(s)" if rutas_adj else "") + ")")
                      if accion_arbol == "referencia":
                          arbol_txt = ("\n=== TEXTOS OFICIALES DE LA EMPRESA ===\n"
                                       + txt_arbol + "\n=== FIN ===\n")
                      else:
                          arbol_txt = ("\n=== TEXTO OBLIGATORIO DE LA RESPUESTA ===\n" + txt_arbol +
                                   "\n=== FIN DEL TEXTO OBLIGATORIO ===\n"
                                   "Este texto es la respuesta que la empresa ha decidido para este caso. "
                                   "USALO COMO CONTENIDO: puedes adaptar el tono y enlazarlo con naturalidad, "
                                   "pero NO cambies los datos, NO quites informacion y NO anadas nada que no "
                                   "este aqui. No pongas saludo ni despedida: se anaden aparte.\n")
                  if txt_arbol and accion_arbol == "texto" and \
                          str(esc.get("texto_literal", "si")).lower() \
                          not in ("no", "false", "off", "0"):
                      # COPY-PASTE: la situacio es clara i l'equip ja ha decidit el text.
                      # No es crida el redactor: s'envia EXACTAMENT el que diu l'arbre.
                      # Aixi no hi ha marge per inventar res, i a mes no gasta model.
                      respuesta = formatear_respuesta(txt_arbol, categoria)
                      confianza, doc = "alta", "NINGUNO"
                      print("    ARBRE: text LITERAL (no passa pel redactor)")
                      print("    ---- RESPUESTA " + "-" * 38)
                      for lin in respuesta.split("\n"):
                          print(f"    | {lin}")
                      print("    " + "-" * 53)
                      resultado = ("BORRADOR (dry: no creado) [ARBOL literal]" if dry
                                   else "BORRADOR CREADO [ARBOL literal]")
                      if not dry:
                          claves_final = set(claves_msg)
                          _rn = (rem or "").strip().lower()
                          if "@" in _rn:
                              claves_final.add("cli:" + _rn)
                          if claves_final & ya_con_borrador:
                              print("    AGRUPAT: el client ja te esborrany en aquesta passada")
                              continue
                          ya_con_borrador |= claves_final
                          reply = msg.Reply()
                          limpiar_adjuntos_heredados(reply)
                          _dest, _why = destinatario_real(
                              msg, rem, esc.get("reenviadores", []),
                              str(esc.get("borradores_para", "") or "").strip())
                          if _dest:
                              reply.To = _dest
                              if _why:
                                  print(f"    destinatari: {_dest} ({_why})")
                          elif _why:
                              print(f"    ATENCIO destinatari: {_why} — REVISAR abans d'enviar")
                          reply.Body = respuesta
                          if rutas_adj:
                              adjuntar(reply, rutas_adj)
                          reply.Save()
                          marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                          creados += 1
                          print(f"    esborrany desat (per a: {_dest or rem})")
                      con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                  " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                  (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto,
                                   categoria, flags, confianza, doc, "ARBOL", resultado, respuesta))
                      con.commit()
                      continue
                  ya_dados = datos_ya_facilitados(asunto, cuerpo)
                  ya_dados_txt = ""
                  if ya_dados:
                      detall = "; ".join(f"{k}: {v}" for k, v in ya_dados)
                      ya_dados_txt = ("\nDATOS QUE EL CLIENTE YA HA FACILITADO EN ESTE MISMO MENSAJE: "
                                      + detall +
                                      ". PROHIBIDO pedirle estos datos: ya los tienes. Si el hilo"
                                      " citado se los pedia, ES ANTIGUO y ya estan respondidos.\n")
                      print(f"    el client ja ha donat: {detall} (no se li demanara)")
                  red = llamar(ia, PROMPT_REDACTOR,
                               f"DATOS VERIFICADOS: {datos}\nHISTORIAL PREVIO con este cliente:\n{hilo_txt}\nCATEGORIA: {categoria}\n{guia}"
                               + arbol_txt + ya_dados_txt +
                               f"MENSAJE del cliente:\nAsunto: {asunto}\n"
                               f"LO QUE ACABA DE ESCRIBIR (responde A ESTO):\n{cuerpo_nuevo(cuerpo)}\n"
                               f"HILO COMPLETO:\n{str(cuerpo or '')[:3000]}\n"
                               "Responde a lo ultimo que ha escrito. Si se refiere a algo que pidio"
                               " antes y sigue sin respuesta ('como le dije', 'en el mensaje"
                               " anterior', 'ya le pregunte'), responde TAMBIEN a aquella peticion.")
                  confianza, doc, respuesta = red.get("confianza", ""), red.get("documento_salida", ""), red.get("respuesta", "")
                  respuesta = formatear_respuesta(respuesta, categoria)
                  if respuesta_vacia(respuesta):
                      print("    RESPOSTA BUIDA despres de netejar — ho mira una persona")
                      resultado = "SIN CONTENIDO UTIL (sin borrador)"
                      print(f"    RESULTADO -> {resultado}")
                      if not dry:
                          marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                      con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                                  " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                                  (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto,
                                   categoria, flags, confianza, doc, "", resultado, ""))
                      con.commit(); continue
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

                # ULTIMA COMPROVACIO, amb la IDENTITAT DEFINITIVA del client.
                # Fins aqui s'han fet servir claus provisionals (remitent, fil).
                # 'rem' es qui el pipeline ha resolt de veritat — el mateix nom
                # que surt a "BORRADOR CREADO (per a: ...)". Si aquest client ja
                # te esborrany en aquesta passada, no se'n crea un segon.
                claves_final = set(claves_msg)
                _rem_norm = (rem or "").strip().lower()
                if "@" in _rem_norm:
                    claves_final.add("cli:" + _rem_norm)
                print("    identitat: " + " | ".join(sorted(claves_final))[:110])
                if claves_final & ya_con_borrador:
                    print("    AGRUPAT (identitat final): aquest client ja te esborrany"
                          " en aquesta passada — es contesta al mail mes recent")
                    resultado = "AGRUPADO (un solo borrador por cliente y pasada)"
                    if not dry:
                        marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                elif mail_id in agrupados_omitir:
                    print("    BLOQUEJAT: no es el mail mes nou del client —"
                          " l'esborrany va al mes recent")
                    resultado = "AGRUPADO (no es el mail mas reciente)"
                    if not dry:
                        marcar_agente(msg, marcar_leido=MARCAR_LEIDO)
                elif dry:
                    ya_con_borrador |= claves_final
                    resultado = "BORRADOR (dry: no creado)" + avisos
                else:
                    ya_con_borrador |= claves_final
                    reply = msg.Reply()
                    limpiar_adjuntos_heredados(reply)
                    _dest, _why = destinatario_real(
                        msg, rem, esc.get("reenviadores", []),
                        str(esc.get("borradores_para", "") or "").strip())
                    if _dest:
                        reply.To = _dest
                        if _why:
                            print(f"    destinatari: {_dest} ({_why})")
                    elif _why:
                        print(f"    ATENCIO destinatari: {_why} — REVISAR abans d'enviar")
                    cuerpo_final = respuesta   # nomes el missatge generat, sense fil citat
                    previos = contexto_grupo.get(mail_id) or []
                    if previos:
                        bloque = [NOTA_INI,
                                  f"Aquest esborrany respon tambe {len(previos)} mail(s) previ(s)"
                                  " del mateix client. Text original, per no haver d'anar a buscar-lo:"]
                        for f_, a_, c_ in previos:
                            bloque += ["", f"[{f_}] Assumpte: {a_}", c_]
                        bloque += [NOTA_FIN, ""]
                        cuerpo_final = "\n".join(bloque) + "\n" + respuesta
                    citar = str(esc.get("citar_original", "si")).lower() not in ("no", "false", "off", "0")
                    escrito = False
                    if citar:
                        # Outlook ja ha muntat la resposta amb el mail del client
                        # citat a sota: nomes cal posar el text nou a dalt.
                        try:
                            nuevo_html = _insertar_arriba_html(reply.HTMLBody, cuerpo_final)
                        except Exception:
                            nuevo_html = None
                        if nuevo_html:
                            try:
                                reply.HTMLBody = nuevo_html
                                escrito = True
                            except Exception:
                                escrito = False
                        if not escrito:
                            try:
                                reply.Body = cuerpo_final + "\n\n" + str(reply.Body or "")
                                escrito = True
                            except Exception:
                                escrito = False
                    if not escrito:
                        reply.Body = cuerpo_final   # nomes el missatge generat
                    if citar:
                        # Comprovacio: Outlook ha inclos de veritat el mail del
                        # client? Si no, el posem nosaltres.
                        try:
                            actual = str(reply.Body or "")
                        except Exception:
                            actual = ""
                        if falta_la_cita(actual, msg):
                            try:
                                reply.Body = actual + cita_propia(msg)
                                print("    (citacio afegida pel nostre compte:"
                                      " Outlook no havia inclos el mail original)")
                            except Exception as e:
                                print(f"    (no s'ha pogut citar l'original: {str(e)[:50]})")
                    destino_seguro = esc.get("borradores_para", "").strip()
                    reply.To = destino_seguro if destino_seguro else rem
                    if previos:
                        # Segell perque cap enviament automatic futur l'agafi tal qual.
                        try:
                            cats = [c_ for c_ in str(reply.Categories or "").split(",") if c_.strip()]
                            if CATEGORIA_NO_AUTO not in cats:
                                cats.append(CATEGORIA_NO_AUTO)
                            reply.Categories = ", ".join(cats)
                        except Exception:
                            pass
                    if rutas_adj:
                        adjuntar(reply, rutas_adj)
                    reply.Save()  # <- ESBORRANY. Mai .Send()
                    marcar_agente(msg, marcar_leido=str(esc.get("marcar_leido", "si")).lower() in ("si", "sí", "true", "on", "1"))
                    creados += 1
                    resultado = ("BORRADOR CREADO (per a: "
                                 + (destino_seguro if destino_seguro else rem) + ")" + avisos)
                print("    ---- RESPUESTA " + "-" * 38)
                for lin in respuesta.split("\n"):
                    print(f"    | {lin}")
                print("    " + "-" * 53)
            filas_csv.append({
                "remitente": rem, "asunto": asunto,
                "pregunta": cuerpo_nuevo(cuerpo)[:900].replace("\n", " "),
                "categoria": categoria, "flags": flags,
                "estado": (ficha or {}).get("estado", ""),
                "estado_codigo": codigo_estado(ficha) or "",
                "matricula": (ficha or {}).get("matricula", ""),
                "importe": (ficha or {}).get("importe_perito", ""),
                "decision": resultado, "respuesta": (respuesta or "").replace("\n", " ")[:1500]})
            print(f"    RESULTADO -> {resultado}")
            con.execute("INSERT OR IGNORE INTO borradores(ts,mail_id,remitente,asunto,categoria,flags,confianza,doc,veredicto_ia,resultado,respuesta)"
                        " VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                        (datetime.now(timezone.utc).isoformat(), mail_id, rem, asunto, categoria,
                         flags, confianza, doc, ver_txt, resultado, respuesta))
            con.commit()
        except Exception as e:
            print(f"    ERROR con este mail (se continua): {str(e)[:120]}")
    extra = f" ({omesos_registre} omesos en silenci: ja al registre)" if omesos_registre else ""
    if motivos_sin:
        print("\n    PER QUE NO S'HA REDACTAT (recompte):")
        for mot, cnt in sorted(motivos_sin.items(), key=lambda p: -p[1]):
            print(f"       {cnt:4}  {mot}")
    if ruta_csv and filas_csv:
        import csv as _csv
        cols = ["remitente", "asunto", "pregunta", "categoria", "flags", "estado",
                "estado_codigo", "matricula", "importe", "decision", "respuesta"]
        try:
            # utf-8-sig perque Excel obri be els accents
            with open(ruta_csv, "w", newline="", encoding="utf-8-sig") as fh:
                w = _csv.DictWriter(fh, fieldnames=cols, delimiter=";")
                w.writeheader()
                for f_ in filas_csv:
                    w.writerow({c: f_.get(c, "") for c in cols})
            print(f"\n    CSV desat: {ruta_csv}  ({len(filas_csv)} files)")
        except Exception as e:
            print(f"\n    (no s'ha pogut desar el CSV: {str(e)[:70]})")
    if filas_csv:
        from collections import Counter
        print("\n    RECOMPTE PER CATEGORIA")
        for k, v in Counter(f_["categoria"] or "(cap)" for f_ in filas_csv).most_common():
            print(f"       {v:5}  {k}")
        print("    RECOMPTE PER ESTAT")
        for k, v in Counter(f_["estado_codigo"] or "(sense fitxa)" for f_ in filas_csv).most_common():
            print(f"       {v:5}  {k}")
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
    ap.add_argument("--solo-clasificar", dest="solo_clasificar", action="store_true",
                    help="nomes classifica (sense redactar): scan barat de tota la bustia")
    ap.add_argument("--csv", metavar="FITXER", default="",
                    help="desa un CSV amb pregunta, resposta i decisio de cada mail")
    args = ap.parse_args()
    esc = cargar_escenario()
    PLANTILLAS.update(cargar_plantillas())
    cargar_arbol()
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
    plant_on = str(esc.get("plantillas", "no")).lower() not in ("no", "off", "false", "0")
    print(" Plantilles al redactor:", "SI" if plant_on else "NO (pur agent: nomes dades + regles)")
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
        store = escoger_buzon(stores, args.outlook)
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
            marcar_agente(msg, marcar_leido=MARCAR_LEIDO)  # marca indeleble + llegit: veterans segellats
        con.commit()
        print(f"Vacunats {n} mails de '{args.carpeta}': marca 'Agente' + llegits + registre. Mai mes es tocaran.")
        con.close(); os._exit(0)
    elif args.outlook:
        escalfar(esc["ia"])
        procesar_carpeta(args.outlook, args.carpeta, esc, con, bbdd, args.dry, args.max,
                     args.solo_clasificar, args.csv)
        print(); informe(con)
        import gc; gc.collect()
        con.close()
        os._exit(0)
    else:
        print("Cal --outlook \"NOM\" (i opcionalment --carpeta, --dry, --max).")