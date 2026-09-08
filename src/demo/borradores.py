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
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)
from reglas import triaje

RUTA_ESC = os.path.join(AQUI, "escenario.json")

MODO_CATEGORIA = {
    "estado_reclamacion": "AUTO", "confirmacion_documentacion": "AUTO",
    "problema_web_subida": "AUTO", "informacion_general": "AUTO",
    "coste_comision": "AUTO", "cortesia_breve": "AUTO",
    "elegibilidad_vehiculo": "BORRADOR", "falta_factura_precio": "BORRADOR",
    "poderes_pleitos": "BORRADOR", "envio_documentacion": "BORRADOR",
}

PROMPT_CLASIFICADOR = """Eres el clasificador del buzon de atencion de una empresa de reclamaciones de vehiculos. Devuelve EXCLUSIVAMENTE un JSON valido:
{"categoria": "<una>", "sospecha_sucesion": true|false, "cancelacion": true|false, "complejo": true|false, "repregunta_insatisfecha": true|false, "motivo": "<una frase>"}
Categorias: estado_reclamacion, falta_factura_precio, envio_documentacion, confirmacion_documentacion, problema_web_subida, elegibilidad_vehiculo, informacion_general, coste_comision, poderes_pleitos, titularidad_caso_especial, cancelacion_desistimiento, cortesia_breve, contacto_llamada, ambiguo
Reglas: sospecha_sucesion=true ante CUALQUIER mencion a fallecimiento/herencia/viudedad/"era cliente" (ante la duda, true). cancelacion=true si expresa voluntad de desistir. complejo=true si varias peticiones, enojo, excepciones o dudas. cortesia_breve = agradecimientos/acuses SIN peticion nueva. Nada fuera del JSON."""

PROMPT_REDACTOR = """Eres el redactor de respuestas del buzon de atencion de una empresa de reclamaciones de vehiculos. Tono cercano y claro, frases cortas, cero jerga juridica, en castellano. Tratamiento SIEMPRE de usted (nunca tutees). Firma SIEMPRE exactamente asi, en dos lineas finales: "Un saludo," y "El equipo de atencion".

Reglas INQUEBRANTABLES:
1. Solo afirmas datos que esten en DATOS VERIFICADOS o en el HILO. Si el dato necesario para responder NO esta, no lo inventes: confianza "baja".
2. Sin promesas de plazos ni resultados que no esten en DATOS VERIFICADOS.
3. Documento de salida: SOLO si el cliente PIDE explicitamente ese contenido en su mensaje. Si el cliente solo pregunta por el estado, agradece o hace una consulta puntual: "NINGUNO". Adjuntar un documento no pedido es un ERROR.
   Catalogo: D1 = Instructivo general del proceso (solo si pregunta como funciona el proceso). D2 = Guia de documentacion alternativa a la factura (solo si dice que no tiene o no encuentra la factura).
4. 50-130 palabras. Saludo con el nombre si consta, respuesta directa, siguiente paso si lo hay, despedida. Un solo tema.
5. Si el cliente dice que adjunta algo pero no consta: pide que lo reenvie, no confirmes recepciones.

Devuelve EXCLUSIVAMENTE: {"respuesta":"<texto>","documento_salida":"D1|D2|NINGUNO","confianza":"alta|media|baja","motivo_confianza":"<una frase>"}"""

PROMPT_VERIFICADOR = """Eres el verificador final de respuestas automaticas de un buzon de atencion. Comprueba en orden:
1. Todo dato concreto de la RESPUESTA (fases, fechas, importes, referencias) aparece en DATOS VERIFICADOS o en el MENSAJE del cliente.
2. No menciona a terceras personas ni otros expedientes.
3. No promete plazos ni resultados no verificados.
4. Documento adjunto: SOLO es valido si el cliente lo PIDE explicitamente en su mensaje (D1 solo si pregunta como funciona el proceso; D2 solo si dice no tener factura). Si hay documento adjunto sin peticion explicita del cliente: RECHAZADO.
5. Tratamiento de usted en todo el texto (si tutea: RECHAZADO) y firma como "El equipo de atencion".
6. Tono adecuado y responde a lo que se pregunta.
Ante cualquier duda razonable: RECHAZADO (un rechazo solo cuesta revision humana; una aprobacion erronea llega a un cliente).
Devuelve EXCLUSIVAMENTE: {"veredicto":"APROBADO|RECHAZADO","problemas":["..."],"riesgo":"bajo|medio|alto"}"""


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


def extraer_cliente_de_reenvio(cuerpo):
    """Busca al cos citat la linia 'De:/From:' del mail original i en treu l'email.
    Nomes es crida quan el remitent real es un reenviador de confianca."""
    for m in RE_DE.finditer(cuerpo or ""):
        etiqueta = m.group(0).strip().lower()
        if etiqueta.startswith(("para", "to", "a:")):
            continue
        return m.group(2).lower()
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
        if os.path.exists(os.path.join(AQUI, "credenciales.json")):
            try:
                from conector_bbdd import ficha_por_email
                self._real = ficha_por_email
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


def llamar(ia, system, user):
    payload = json.dumps({"model": ia["modelo"], "temperature": 0,
        "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}).encode("utf-8")
    req = urllib.request.Request(ia["base_url"].rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {ia['api_key']}"})
    with urllib.request.urlopen(req, timeout=300) as r:
        data = json.loads(r.read().decode("utf-8"))
    t = data["choices"][0]["message"]["content"].replace("```json", "").replace("```", "").strip()
    return json.loads(t[t.find("{"):t.rfind("}") + 1], strict=False)


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
            if msg.SenderEmailType == "EX":
                ex = msg.Sender.GetExchangeUser()
                rem = (ex.PrimarySmtpAddress if ex else "desconocido@exchange").lower()
            else:
                rem = (msg.SenderEmailAddress or "desconocido").lower()
            asunto = str(msg.Subject or "")
            cuerpo = str(msg.Body or "")[:5000]
            adjuntos = [{"nombre": str(a.FileName),
                         "is_inline": str(a.FileName).lower().startswith("image00"),
                         "bytes": getattr(a, "Size", 0)} for a in msg.Attachments]
            reenviadores = esc.get("reenviadores", [])
            if es_reenviador(rem, reenviadores):
                extraido = extraer_cliente_de_reenvio(cuerpo)
                if extraido:
                    print(f"\n=== reenviament de {rem}")
                    print(f"    client extret del cos: {extraido}")
                    rem = extraido
            mail = {"remitente": rem, "asunto": asunto, "cuerpo": cuerpo, "adjuntos": adjuntos}

            print(f"=== {rem} | {asunto[:50]}")
            historial = esc["historial"] if isinstance(esc["historial"], HistorialAbierto) else set(esc["historial"])
            destino, motivo, ficha = triaje(mail, historial, bbdd)
            if destino == "CIRCUITO" and ficha and ficha.get("sospecha_sucesion_bbdd"):
                destino, motivo = "HUMANO_SUCESION", "herencia_marcada_en_bbdd"
            categoria = flags = confianza = doc = ver_txt = respuesta = ""
            resultado = destino + " (" + motivo + ")"

            if destino == "CIRCUITO":
                c = llamar(ia, PROMPT_CLASIFICADOR, f"HILO: (no disponible)\n\nMENSAJE:\nAsunto: {asunto}\nCuerpo: {cuerpo}")
                categoria = c.get("categoria", "ambiguo")
                flags = ", ".join(k for k in ("sospecha_sucesion", "cancelacion", "complejo",
                                              "repregunta_insatisfecha") if c.get(k))
                print(f"    categoria: {categoria} | flags: {flags or 'ninguno'}")
                if flags:
                    resultado = "HUMANO (flags)"
                elif MODO_CATEGORIA.get(categoria) is None:
                    resultado = f"HUMANO (categoria {categoria})"
                else:
                    datos = ficha_a_datos(ficha)
                    print(f"    ficha: {datos[:100]}...")
                    print("    REDACTOR -> escribiendo...")
                    red = llamar(ia, PROMPT_REDACTOR,
                                 f"DATOS VERIFICADOS: {datos}\nHILO: (no disponible)\nCATEGORIA: {categoria}\n"
                                 f"MENSAJE del cliente:\nAsunto: {asunto}\nCuerpo: {cuerpo}")
                    confianza, doc, respuesta = red.get("confianza", ""), red.get("documento_salida", ""), red.get("respuesta", "")
                    print(f"    confianza: {confianza} | doc: {doc}")
                    if confianza == "baja":
                        resultado = "HUMANO (confianza baja)"
                    else:
                        print("    VERIFICADOR -> revisando...")
                        ver = llamar(ia, PROMPT_VERIFICADOR,
                                     f"DATOS VERIFICADOS: {datos}\nMENSAJE del cliente: {cuerpo}\n"
                                     f"RESPUESTA PROPUESTA (doc: {doc}): {respuesta}")
                        ver_txt = ver.get("veredicto", "")
                        if ver_txt != "APROBADO":
                            resultado = "HUMANO (verificador RECHAZO: " + "; ".join(ver.get("problemas", []))[:120] + ")"
                        else:
                            if dry:
                                resultado = "BORRADOR (dry: no creado)"
                            else:
                                reply = msg.Reply()
                                reply.Body = respuesta + "\n\n" + reply.Body
                                destino_seguro = esc.get("borradores_para", "").strip()
                                if destino_seguro:
                                    reply.To = destino_seguro  # cinturo del pilot: mai a clients reals
                                reply.Save()  # <- ESBORRANY. Mai .Send()
                                creados += 1
                                resultado = ("BORRADOR CREADO en Outlook"
                                             + (f" (destinatari forcat: {destino_seguro})" if destino_seguro else ""))
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
        procesar_carpeta(args.outlook, args.carpeta, esc, con, bbdd, args.dry, args.max)
        print(); informe(con)
        import gc; gc.collect()
        con.close()
        os._exit(0)
    else:
        print("Cal --outlook \"NOM\" (i opcionalment --carpeta, --dry, --max).")