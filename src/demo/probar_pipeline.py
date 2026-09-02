# -*- coding: utf-8 -*-
"""Pipeline complet en proves: REDACTOR + VERIFICADOR sobre mails ja classificats.

Dos modes:
  python src\\demo\\probar_pipeline.py --log     -> processa els CIRCUITO del registre
                                                  de probar_mail.py (log_pruebas.sqlite)
  python src\\demo\\probar_pipeline.py           -> mode interactiu: mail nou, tot el
                                                  recorregut (triatge->classif->redactor->verif)

Usa el mateix escenario.json (fitxes BBDD, model, URL). Tot local, cost zero.
"""
import json, os, sys, sqlite3, urllib.request, argparse
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
from reglas import triaje

RUTA_ESC = os.path.join(AQUI, "escenario.json")

MODO_CATEGORIA = {  # mode de la categoria segons el doc 03
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
        print("No existe escenario.json — ejecuta antes probar_mail.py una vez, o crealo."); sys.exit(1)
    esc = json.load(open(RUTA_ESC, encoding="utf-8-sig"))
    esc["historial"] = [h.strip().lower() for h in esc["historial"]]
    esc["bbdd"] = {k.strip().lower(): v for k, v in esc["bbdd"].items()}
    return esc


def llamar(ia, system, user):
    payload = json.dumps({"model": ia["modelo"], "temperature": 0,
        "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}).encode("utf-8")
    req = urllib.request.Request(ia["base_url"].rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {ia['api_key']}"})
    with urllib.request.urlopen(req, timeout=300) as r:
        data = json.loads(r.read().decode("utf-8"))
    t = data["choices"][0]["message"]["content"].replace("```json", "").replace("```", "").strip()
    return json.loads(t[t.find("{"):t.rfind("}") + 1])


def log_init():
    con = sqlite3.connect(os.path.join(AQUI, "log_pipeline.sqlite"))
    con.execute("""CREATE TABLE IF NOT EXISTS pipeline(
        id INTEGER PRIMARY KEY, ts TEXT, remitente TEXT, asunto TEXT, categoria TEXT,
        confianza TEXT, doc_salida TEXT, veredicto_ia TEXT, resultado TEXT,
        respuesta TEXT, veredicto_humano TEXT)""")
    return con


def procesar(ia, esc, con, remitente, asunto, cuerpo, categoria, flags, hilo=""):
    """Redactor + verificador per a un mail ja classificat com a CIRCUITO."""
    print(f"\n=== {remitente} | {asunto}")
    print(f"    categoria: {categoria} | flags: {flags or 'ninguno'}")
    if flags:
        print("    RESULTADO -> HUMANO (flags encendidos: el circuito no lo toca)")
        return
    modo = MODO_CATEGORIA.get(categoria)
    if modo is None:
        print(f"    RESULTADO -> HUMANO (categoria '{categoria}' es de gestion humana)")
        return
    ficha = esc["bbdd"].get(remitente.lower())
    if not ficha:
        print("    RESULTADO -> HUMANO (sin ficha en BBDD)")
        return
    datos = f"Titular: {ficha['titular']} · Matricula: {ficha['matricula']} · Estado del expediente: {ficha['estado']}"
    print(f"    ficha: {datos}")
    print("    REDACTOR -> escribiendo...")
    red = llamar(ia, PROMPT_REDACTOR,
                 f"DATOS VERIFICADOS: {datos}\nHILO: {hilo or '(vacio)'}\nCATEGORIA: {categoria}\n"
                 f"MENSAJE del cliente:\nAsunto: {asunto}\nCuerpo: {cuerpo}")
    print(f"    confianza: {red.get('confianza')} | doc: {red.get('documento_salida')}")
    print("    ---- RESPUESTA PROPUESTA " + "-" * 30)
    for linea in red.get("respuesta", "").split("\n"):
        print(f"    | {linea}")
    print("    " + "-" * 55)
    resultado, ver = "", {}
    if red.get("confianza") == "baja":
        resultado = "HUMANO (confianza baja del redactor)"
    else:
        print("    VERIFICADOR -> revisando...")
        ver = llamar(ia, PROMPT_VERIFICADOR,
                     f"DATOS VERIFICADOS: {datos}\nMENSAJE del cliente: {cuerpo}\n"
                     f"RESPUESTA PROPUESTA (doc: {red.get('documento_salida')}): {red.get('respuesta')}")
        if ver.get("veredicto") != "APROBADO":
            resultado = f"HUMANO (verificador RECHAZO: {'; '.join(ver.get('problemas', [])) or 'sin detalle'})"
        else:
            resultado = "ENVIO AUTOMATICO (via cola diferida)" if modo == "AUTO" else "BORRADOR para aprobar por una persona"
            print(f"    VERIFICADOR -> APROBADO (riesgo {ver.get('riesgo')})")
    print(f"    RESULTADO FINAL -> {resultado}")
    ok = input("    ¿Te parece correcto el conjunto? (s/n/enter): ").strip()
    con.execute("INSERT INTO pipeline(ts,remitente,asunto,categoria,confianza,doc_salida,veredicto_ia,resultado,respuesta,veredicto_humano)"
                " VALUES(?,?,?,?,?,?,?,?,?,?)",
                (datetime.now(timezone.utc).isoformat(), remitente, asunto, categoria,
                 red.get("confianza", ""), red.get("documento_salida", ""),
                 ver.get("veredicto", ""), resultado, red.get("respuesta", ""), ok))
    con.commit()


def modo_log(ia, esc, con):
    ruta = os.path.join(AQUI, "log_pruebas.sqlite")
    if not os.path.exists(ruta):
        print("No hay log_pruebas.sqlite — haz antes pruebas con probar_mail.py"); return
    filas = sqlite3.connect(ruta).execute(
        "SELECT remitente,asunto,cuerpo,categoria,flags FROM pruebas WHERE destino='CIRCUITO' AND categoria<>''").fetchall()
    print(f"Encontrados {len(filas)} mails de CIRCUITO ya clasificados en el registro.\n(Los de HUMANO no se procesan: ya son cosa del equipo.)")
    for rem, asu, cue, cat, flags in filas:
        procesar(ia, esc, con, rem, asu, cue, cat, "" if flags in ("", "ninguno") else flags)


def modo_interactivo(ia, esc, con):
    print("Mail nuevo (recorrido completo). Enter en remitente = cliente.uno@mail.com; 'salir' para acabar.")
    while True:
        rem = (input("\nRemitente [cliente.uno@mail.com]: ").strip() or "cliente.uno@mail.com").lower()
        if rem == "salir":
            break
        asu = input("Asunto [Consulta]: ").strip() or "Consulta"
        print("Cuerpo (linea vacia para terminar):")
        lineas = []
        while True:
            l = input()
            if not l.strip(): break
            lineas.append(l)
        cue = "\n".join(lineas)
        adj = input("¿Adjunto real? (s/n) [n]: ").strip().lower().startswith("s")
        hilo = input("Hilo previo (enter si no hay): ").strip()
        mail = {"remitente": rem, "asunto": asu, "cuerpo": cue,
                "adjuntos": [{"nombre": "doc.pdf", "is_inline": False, "bytes": 200000}] if adj else []}
        destino, motivo, _ = triaje(mail, set(esc["historial"]), esc["bbdd"])
        print(f"TRIAJE -> {destino} ({motivo})")
        if destino != "CIRCUITO":
            print("(va al equipo con resumen; el pipeline no lo toca)"); continue
        print("CLASIFICADOR -> pensando...")
        c = llamar(ia, PROMPT_CLASIFICADOR, f"HILO: {hilo or '(vacio)'}\n\nMENSAJE:\nAsunto: {asu}\nCuerpo: {cue}")
        flags = ", ".join(k for k in ("sospecha_sucesion", "cancelacion", "complejo", "repregunta_insatisfecha") if c.get(k))
        procesar(ia, esc, con, rem, asu, cue, c.get("categoria", "ambiguo"), flags, hilo)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--log", action="store_true", help="procesar los CIRCUITO del registro existente")
    args = ap.parse_args()
    esc = cargar_escenario()
    con = log_init()
    print("=" * 64)
    print(" PIPELINE COMPLET: redactor + verificador | model:", esc["ia"]["modelo"])
    print("=" * 64)
    (modo_log if args.log else modo_interactivo)(esc["ia"], esc, con)
    n = con.execute("SELECT COUNT(*) FROM pipeline").fetchone()[0]
    print(f"\nRegistro: {n} respuestas acumuladas en src\\demo\\log_pipeline.sqlite")