# -*- coding: utf-8 -*-
"""Primera clasificacion con IA: pasa los mails de la demo por el triaje y,
los que llegan a CIRCUITO, los clasifica con el prompt v1.0 contra Ollama
(endpoint compatible OpenAI). Solo libreria estandar: no necesita pip.

Ejecutar desde la raiz del repo:   python src\\demo\\clasificar_demo.py

El dia que llegue la API real de OpenAI, basta cambiar la configuracion:
  BASE_URL = "https://api.openai.com/v1"   API_KEY = "sk-..."   MODEL = "gpt-5-mini"
"""
import json, os, sys, urllib.request

# ----------------- Configuracion -----------------
BASE_URL = os.environ.get("IA_BASE_URL", "http://localhost:11434/v1")
API_KEY  = os.environ.get("IA_API_KEY", "ollama")          # Ollama ignora la clave
MODEL    = os.environ.get("IA_MODELO", "qwen2.5:7b")

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
from reglas import triaje

HISTORIAL_MOCK = {"cliente.uno@mail.com", "cliente.dos@mail.com", "cliente.tres@mail.com"}
BBDD_MOCK = {
    "cliente.uno@mail.com": {"matricula": "2717GJD", "titular": "Cliente Uno",
                             "estado": "En fase de demanda presentada"},
    "cliente.dos@mail.com": {"matricula": "4501KLM", "titular": "Cliente Dos",
                             "estado": "Pendiente de documentacion"},
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
5. Nada de texto fuera del JSON."""


def clasificar(mensaje, hilo=""):
    payload = json.dumps({
        "model": MODEL,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": PROMPT_CLASIFICADOR},
            {"role": "user", "content": f"HILO: {hilo or '(vacio)'}\n\nMENSAJE:\nAsunto: {mensaje['asunto']}\nCuerpo: {mensaje['cuerpo']}"},
        ],
    }).encode("utf-8")
    req = urllib.request.Request(
        BASE_URL.rstrip("/") + "/chat/completions", data=payload,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {API_KEY}"})
    with urllib.request.urlopen(req, timeout=180) as r:
        data = json.loads(r.read().decode("utf-8"))
    texto = data["choices"][0]["message"]["content"]
    limpio = texto.replace("```json", "").replace("```", "").strip()
    ini, fin = limpio.find("{"), limpio.rfind("}")
    return json.loads(limpio[ini:fin + 1])


def main():
    mails = json.load(open(os.path.join(AQUI, "mails_demo.json"), encoding="utf-8"))
    print(f"Modelo: {MODEL} en {BASE_URL}\n")
    algun_circuito = False
    for m in mails:
        destino, motivo, ficha = triaje(m, HISTORIAL_MOCK, BBDD_MOCK)
        if destino != "CIRCUITO":
            print(f"[triaje] {m['remitente']:28} -> {destino} ({motivo})  [no se clasifica]")
            continue
        algun_circuito = True
        print(f"[triaje] {m['remitente']:28} -> CIRCUITO. Clasificando con IA...")
        try:
            res = clasificar(m)
            flags = [k for k in ("sospecha_sucesion", "cancelacion", "complejo",
                                 "repregunta_insatisfecha") if res.get(k)]
            print(f"         Categoria: {res.get('categoria','?'):26} "
                  f"flags: {', '.join(flags) if flags else 'ninguno'}")
            print(f"         Motivo: {res.get('motivo','')}\n")
        except Exception as e:
            print(f"         ERROR llamando a la IA: {e}")
            print("         Comprueba que Ollama esta corriendo (icono en la bandeja,"
                  " o 'ollama serve') y que el modelo esta descargado (ollama list).\n")
            return
    if algun_circuito:
        print("Hito conseguido: el agente ha CLASIFICADO con IA por primera vez.")


if __name__ == "__main__":
    main()
