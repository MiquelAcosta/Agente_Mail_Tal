# -*- coding: utf-8 -*-
"""Banc de proves interactiu de l'agent.
Escriu qualsevol mail, canvia les variables (adjunt, historial, fitxa BBDD)
i mira que decideixen el triatge i la IA. Tot en local, cost zero.

Executar des de l'arrel del repo:   python src\\demo\\probar_mail.py

Les variables de l'escenari (historial, fitxes de la BBDD, model) es poden
editar al fitxer src\\demo\\escenario.json sense tocar codi.
"""
import json, os, sys, sqlite3, urllib.request
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
from reglas import triaje

RUTA_ESC = os.path.join(AQUI, "escenario.json")

ESC_DEFECTO = {
    "ia": {"base_url": "http://localhost:11434/v1", "api_key": "ollama", "modelo": "qwen2.5:7b"},
    "historial": ["cliente.uno@mail.com", "cliente.dos@mail.com"],
    "bbdd": {
        "cliente.uno@mail.com": {"matricula": "2717GJD", "titular": "Cliente Uno",
                                  "estado": "En fase de demanda presentada"},
        "cliente.dos@mail.com": {"matricula": "4501KLM", "titular": "Cliente Dos",
                                  "estado": "Pendiente de documentación: falta factura de compra"},
    },
}

PROMPT_CLASIFICADOR = """Eres el clasificador del buzon de atencion de una empresa de reclamaciones de vehiculos. Tu unica tarea es clasificar; no respondes al cliente.

Recibiras MENSAJE (lo que escribe el cliente) e HILO (resumen de la conversacion previa, puede estar vacio). Clasifica el MENSAJE usando el HILO como contexto.

Devuelve EXCLUSIVAMENTE un JSON valido, sin texto antes ni despues:
{"categoria": "<una de la lista>", "sospecha_sucesion": true|false, "cancelacion": true|false, "complejo": true|false, "repregunta_insatisfecha": true|false, "motivo": "<una frase>"}

Categorias permitidas:
estado_reclamacion, falta_factura_precio, envio_documentacion, confirmacion_documentacion, problema_web_subida, elegibilidad_vehiculo, informacion_general, coste_comision, poderes_pleitos, titularidad_caso_especial, cancelacion_desistimiento, cortesia_breve, contacto_llamada, ambiguo

Reglas:
1. sospecha_sucesion=true ante CUALQUIER mencion a fallecimiento, defuncion, herencia, herederos, testamento, viudedad, "era cliente". Ante la duda: true.
2. cancelacion=true si expresa voluntad de desistir, aunque sea dudosa.
3. complejo=true si hay varias peticiones, enojo o amenaza, peticion de excepcion, o no estas seguro.
4. repregunta_insatisfecha=true si responde con queja o insistencia a una respuesta previa del buzon sobre lo mismo.
5. cortesia_breve = "Gracias", "Recibido", "Ok", "Muchas gracias por la informacion": agradecimientos y acuses SIN ninguna peticion nueva.
6. Nada de texto fuera del JSON."""


def cargar_escenario():
    if not os.path.exists(RUTA_ESC):
        with open(RUTA_ESC, "w", encoding="utf-8") as f:
            json.dump(ESC_DEFECTO, f, ensure_ascii=False, indent=2)
        print(f"(Creado {RUTA_ESC} con el escenario por defecto — editalo para cambiar variables)\n")
    esc = json.load(open(RUTA_ESC, encoding="utf-8-sig"))
    # Normalizar: mayusculas y espacios en el escenario ya no importan
    esc["historial"] = [h.strip().lower() for h in esc["historial"]]
    esc["bbdd"] = {k.strip().lower(): v for k, v in esc["bbdd"].items()}
    return esc


def clasificar(ia, asunto, cuerpo, hilo=""):
    payload = json.dumps({
        "model": ia["modelo"], "temperature": 0,
        "messages": [
            {"role": "system", "content": PROMPT_CLASIFICADOR},
            {"role": "user", "content": f"HILO: {hilo or '(vacio)'}\n\nMENSAJE:\nAsunto: {asunto}\nCuerpo: {cuerpo}"},
        ],
    }).encode("utf-8")
    req = urllib.request.Request(
        ia["base_url"].rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {ia['api_key']}"})
    with urllib.request.urlopen(req, timeout=300) as r:
        data = json.loads(r.read().decode("utf-8"))
    texto = data["choices"][0]["message"]["content"]
    limpio = texto.replace("```json", "").replace("```", "").strip()
    return json.loads(limpio[limpio.find("{"):limpio.rfind("}") + 1])


def preguntar(texto, defecto=""):
    v = input(f"{texto}{' [' + defecto + ']' if defecto else ''}: ").strip()
    return v or defecto


def leer_cuerpo():
    print("Cuerpo del mail (varias lineas; linea vacia para terminar):")
    lineas = []
    while True:
        l = input()
        if not l.strip():
            break
        lineas.append(l)
    return "\n".join(lineas)


def log_init():
    con = sqlite3.connect(os.path.join(AQUI, "log_pruebas.sqlite"))
    con.execute("""CREATE TABLE IF NOT EXISTS pruebas(
        id INTEGER PRIMARY KEY, ts TEXT, remitente TEXT, asunto TEXT, cuerpo TEXT,
        destino TEXT, motivo TEXT, categoria TEXT, flags TEXT, veredicto TEXT)""")
    return con


def main():
    esc = cargar_escenario()
    ia = esc["ia"]
    con = log_init()
    print("=" * 64)
    print(" BANC DE PROVES DE L'AGENT — escriu mails, mira que decideix")
    print(f" Model: {ia['modelo']} | Escenari editable a: escenario.json")
    print(" (escriu 'salir' com a remitent per acabar, 'esc' per veure escenari)")
    print("=" * 64)
    while True:
        print()
        rem = preguntar("Remitente", "cliente.uno@mail.com").lower()
        if rem == "salir":
            break
        if rem == "esc":
            print(json.dumps({"historial": esc["historial"],
                              "bbdd": list(esc["bbdd"].keys())}, ensure_ascii=False, indent=2))
            continue
        asunto = preguntar("Asunto", "Consulta")
        cuerpo = leer_cuerpo()
        adj = preguntar("¿Trae adjunto real? (s/n)", "n").lower().startswith("s")
        hilo = preguntar("Resumen del hilo previo (enter si no hay)", "")

        mail = {"remitente": rem, "asunto": asunto, "cuerpo": cuerpo,
                "adjuntos": [{"nombre": "doc.pdf", "is_inline": False, "bytes": 200000}] if adj else []}
        destino, motivo, ficha = triaje(mail, set(esc["historial"]), esc["bbdd"])

        print("-" * 64)
        print(f" TRIAJE  -> {destino}  (motivo: {motivo})")
        categoria, flags_txt = "", ""
        if ficha:
            print(f" FICHA   -> {ficha['titular']} · {ficha['matricula']} · {ficha['estado']}")
        if destino == "CIRCUITO":
            print(" IA      -> clasificando (la 1a vez tarda: carga del modelo)...")
            try:
                res = clasificar(ia, asunto, cuerpo, hilo)
                categoria = res.get("categoria", "?")
                flags = [k for k in ("sospecha_sucesion", "cancelacion", "complejo",
                                     "repregunta_insatisfecha") if res.get(k)]
                flags_txt = ", ".join(flags) if flags else "ninguno"
                print(f" IA      -> categoria: {categoria} | flags: {flags_txt}")
                print(f"            motivo: {res.get('motivo','')}")
                if flags:
                    print("            (con estos flags, el sistema real lo derivaria a humano)")
            except Exception as e:
                print(f" ERROR llamando a la IA: {e}")
                print(" ¿Esta Ollama corriendo? (icono en la bandeja o 'ollama serve')")
        else:
            print(" (no se clasifica: en el sistema real iria a la carpeta del equipo,")
            print("  con el resumen del caso preparado)")
        print("-" * 64)
        ok = preguntar("¿La decision te parece CORRECTA? (s/n/enter para saltar)", "")
        con.execute("INSERT INTO pruebas(ts,remitente,asunto,cuerpo,destino,motivo,categoria,flags,veredicto)"
                    " VALUES(?,?,?,?,?,?,?,?,?)",
                    (datetime.now(timezone.utc).isoformat(), rem, asunto, cuerpo,
                     destino, motivo, categoria, flags_txt, ok))
        con.commit()
    n = con.execute("SELECT COUNT(*) FROM pruebas").fetchone()[0]
    malas = con.execute("SELECT COUNT(*) FROM pruebas WHERE veredicto='n'").fetchone()[0]
    print(f"\nSesion guardada: {n} pruebas acumuladas ({malas} marcadas como incorrectas)")
    print("Detalle en src\\demo\\log_pruebas.sqlite — las incorrectas son la lista de mejoras del prompt.")


if __name__ == "__main__":
    main()