# -*- coding: utf-8 -*-
"""REVISIO DIARIA — pregunta del client + resposta que va donar l'agent.

Creua el registre (que te la resposta) amb la bustia d'Outlook (que te el cos
del correu). Nomes LLEGEIX: no crea, no mou, no marca. Cost zero.

Us:
  python src\\demo\\revision_diaria.py                 (avui)
  python src\\demo\\revision_diaria.py --dias 3
  python src\\demo\\revision_diaria.py --fecha 2026-09-28
"""
import os, sys, csv, sqlite3, argparse
from datetime import date, timedelta

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)

from borradores import (cargar_escenario, escoger_buzon, id_estable,
                        componentes_fecha, fecha_texto, cuerpo_nuevo)

DB = os.path.join(AQUI, "log_borradores.sqlite")
CALAIXOS = ["1 FACIL", "2 DIFICIL", "3 DESISTIMIENTO", "4 ESCALADOS",
            "0 DESCARTES", "MULTIPLES", "Bandeja de entrada", "Inbox"]


def recorrer(carpeta, cossos, vistos):
    """Apunta el cos de cada correu indexat pel seu id estable."""
    try:
        items = carpeta.Items
        m = items.GetFirst()
        while m is not None:
            try:
                if getattr(m, "Class", 0) == 43:
                    cossos[id_estable(m)] = {
                        "asunto": str(getattr(m, "Subject", "") or ""),
                        "cuerpo": cuerpo_nuevo(str(getattr(m, "Body", "") or "")),
                        "fecha": fecha_texto(componentes_fecha(m)),
                        "carpeta": str(carpeta.Name),
                    }
                    vistos[0] += 1
            except Exception:
                pass
            m = items.GetNext()
    except Exception:
        pass
    try:
        for i in range(carpeta.Folders.Count):
            recorrer(carpeta.Folders.Item(i + 1), cossos, vistos)
    except Exception:
        pass


def main(buzon, dias, fecha_fija, ruta_csv):
    import win32com.client
    if fecha_fija:
        desde = fecha_fija
    else:
        desde = (date.today() - timedelta(days=max(int(dias) - 1, 0))).isoformat()
    print(f"Registre des de: {desde}")

    con = sqlite3.connect(DB)
    filas = con.execute(
        "SELECT ts, mail_id, remitente, asunto, categoria, flags, confianza,"
        " resultado, respuesta FROM borradores WHERE ts >= ? ORDER BY ts",
        (desde,)).fetchall()
    con.close()
    if not filas:
        print("Cap registre en aquest periode.")
        return
    print(f"Registres trobats: {len(filas)}")

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = escoger_buzon(stores, buzon)
    cossos, vistos = {}, [0]
    if store is not None:
        print("Llegint la bustia (pot trigar)...")
        recorrer(store, cossos, vistos)
        print(f"Correus indexats: {vistos[0]}")
    else:
        print(f"(AVIS: no trobo el buzon '{buzon}'; no hi haura el cos dels mails)")

    cols = ["fecha", "carpeta", "cliente", "asunto", "pregunta",
            "categoria", "flags", "confianza", "resultado", "respuesta_agente"]
    n_amb = 0
    with open(ruta_csv, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, delimiter=";")
        w.writeheader()
        for ts, mid, rem, asu, cat, fl, conf, res, resp in filas:
            c = cossos.get(mid, {})
            if c:
                n_amb += 1
            w.writerow({
                "fecha": c.get("fecha") or ts[:16].replace("T", " "),
                "carpeta": c.get("carpeta", ""),
                "cliente": rem or "",
                "asunto": asu or c.get("asunto", ""),
                "pregunta": " ".join((c.get("cuerpo") or "").split())[:1200],
                "categoria": cat or "", "flags": fl or "", "confianza": conf or "",
                "resultado": res or "",
                "respuesta_agente": " ".join((resp or "").split())[:1800],
            })
    print("\n" + "=" * 56)
    print(f"  Files al CSV .............. {len(filas)}")
    print(f"  Amb el cos del mail ....... {n_amb}")
    print(f"  CSV ....................... {ruta_csv}")
    print("=" * 56)
    from collections import Counter
    print("\n  RESULTATS DEL DIA")
    for k, v in Counter((f[7] or "")[:44] for f in filas).most_common(12):
        print(f"     {v:4}  {k}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="BUSTIA", default=None)
    ap.add_argument("--dias", type=int, default=1, help="1 = nomes avui")
    ap.add_argument("--fecha", metavar="AAAA-MM-DD", default="")
    ap.add_argument("--csv", metavar="FITXER", default="")
    a = ap.parse_args()
    destino = a.outlook or cargar_escenario().get("agente", {}).get("buzon", "")
    if not destino:
        print('Cal --outlook "part-del-nom-de-la-bustia".'); sys.exit(1)
    ruta = a.csv or f"revisio_{a.fecha or date.today().isoformat()}.csv"
    main(destino, a.dias, a.fecha, ruta)
    import gc; gc.collect()
    os._exit(0)
