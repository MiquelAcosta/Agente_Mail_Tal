# -*- coding: utf-8 -*-
"""Estadistica dels dos registres de proves (classificador i pipeline).
Executar des de l'arrel del repo:   python src\\demo\\estadistica.py
"""
import sqlite3, os

AQUI = os.path.dirname(os.path.abspath(__file__))

CONFIG = [
    ("Classificador (probar_mail)", "log_pruebas.sqlite", "pruebas", "veredicto", "categoria"),
    ("Pipeline (redactor+verificador)", "log_pipeline.sqlite", "pipeline", "veredicto_humano", "resultado"),
]

for nombre, fichero, tabla, campo, extra in CONFIG:
    ruta = os.path.join(AQUI, fichero)
    print("=" * 60)
    if not os.path.exists(ruta):
        print(f"{nombre}: (encara no hi ha registre)")
        continue
    con = sqlite3.connect(ruta)
    t = con.execute(f"SELECT COUNT(*), SUM({campo}='s'), SUM({campo}='n') FROM {tabla}").fetchone()
    total, ok, mal = t[0], t[1] or 0, t[2] or 0
    sense = total - ok - mal
    print(f"{nombre}")
    linea = f"  Proves: {total} | correctes: {ok} | errades: {mal}"
    if sense:
        linea += f" | sense veredicte: {sense}"
    if ok + mal:
        linea += f" | ENCERT: {ok / (ok + mal) * 100:.0f}%"
    print(linea)
    errades = con.execute(
        f"SELECT ts, asunto, {extra} FROM {tabla} WHERE {campo}='n' ORDER BY id").fetchall()
    if errades:
        print("  Errades marcades (la llista de deures):")
        for ts, asunto, det in errades:
            print(f"   - [{ts[:16]}] {asunto}  ->  {det}")
print("=" * 60)
