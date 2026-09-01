# -*- coding: utf-8 -*-
"""Demo del triaje v2: filtros 0-4 con buzon, historial y BBDD simulados.
Ejecutar desde la raiz del repo:  python src\\demo\\ejecutar_demo.py
"""
import json, sqlite3, sys, os
from datetime import datetime, timezone

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
from reglas import triaje

# --- Mocks: cuando lleguen credenciales reales, esto se sustituye ---
HISTORIAL_MOCK = {"cliente.uno@mail.com", "cliente.dos@mail.com", "cliente.tres@mail.com"}
BBDD_MOCK = {
    "cliente.uno@mail.com":  {"matricula": "2717GJD", "titular": "Cliente Uno",
                              "estado": "En fase de demanda presentada"},
    "cliente.dos@mail.com":  {"matricula": "4501KLM", "titular": "Cliente Dos",
                              "estado": "Pendiente de documentación"},
    # cliente.tres tiene historial de correo pero NO ficha en BBDD (caso filtro 3)
}


def log_init(path):
    con = sqlite3.connect(path)
    con.execute("""CREATE TABLE IF NOT EXISTS decisiones(
        id INTEGER PRIMARY KEY, ts TEXT, remitente TEXT, asunto TEXT,
        destino TEXT, motivo TEXT, matricula TEXT)""")
    return con


def main():
    mails = json.load(open(os.path.join(AQUI, "mails_demo.json"), encoding="utf-8"))
    con = log_init(os.path.join(AQUI, "log_decisiones.sqlite"))
    print(f"{'REMITENTE':30} {'DESTINO':17} MOTIVO")
    print("-" * 75)
    for m in mails:
        destino, motivo, ficha = triaje(m, HISTORIAL_MOCK, BBDD_MOCK)
        matricula = ficha["matricula"] if ficha else None
        # Fase 1 real: aqui se moveria el mail de carpeta via Graph.
        # Fase 2: si destino == CIRCUITO -> clasificador -> redactor -> verificador.
        con.execute("INSERT INTO decisiones(ts,remitente,asunto,destino,motivo,matricula)"
                    " VALUES(?,?,?,?,?,?)",
                    (datetime.now(timezone.utc).isoformat(), m["remitente"],
                     m["asunto"], destino, motivo, matricula))
        print(f"{m['remitente']:30} {destino:17} {motivo}")
    con.commit()
    n = con.execute("SELECT COUNT(*) FROM decisiones").fetchone()[0]
    print(f"\nLog: {n} decisiones acumuladas en src\\demo\\log_decisiones.sqlite")
    print("Todo correcto: el triaje v2 (filtros 0-4) funciona en esta maquina.")


if __name__ == "__main__":
    main()
