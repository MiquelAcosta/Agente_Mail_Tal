# -*- coding: utf-8 -*-
"""RECOMPTE DE VOLUM — quants correus entren cada dia.

Nomes compta dates. No crida cap model, no toca res: cost zero.
Serveix per estimar el cost d'operacio amb dades reals.

Us:
  python src\\demo\\contar_volumen.py --dias 7
  python src\\demo\\contar_volumen.py --dias 30 --csv volum.csv
"""
import os, sys, csv, argparse, collections
from datetime import datetime, timedelta, date

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)

from borradores import cargar_escenario, escoger_buzon, componentes_fecha, remitente_smtp

DIES = ["dilluns", "dimarts", "dimecres", "dijous", "divendres", "dissabte", "diumenge"]


def recorrer(carpeta, limit, dades, nivell=0):
    """Compta els correus de la carpeta i les seves subcarpetes."""
    try:
        nom = str(carpeta.Name)
        items = carpeta.Items
        m = items.GetFirst()
        while m is not None:
            try:
                if getattr(m, "Class", 0) == 43:
                    f = componentes_fecha(m)
                    if f:
                        d = date(f[0], f[1], f[2])
                        if d >= limit:
                            dades.append((d, f[3], nom, remitente_smtp(m) or ""))
            except Exception:
                pass
            m = items.GetNext()
    except Exception:
        pass
    if nivell < 3:
        try:
            for i in range(carpeta.Folders.Count):
                recorrer(carpeta.Folders.Item(i + 1), limit, dades, nivell + 1)
        except Exception:
            pass


def main(buzon, dias, ruta_csv):
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = escoger_buzon(stores, buzon)
    if store is None:
        print(f"Cap buzon conte '{buzon}'.")
        return
    limit = date.today() - timedelta(days=int(dias))
    print(f"Bustia: {store.Name} | des de {limit.isoformat()}")
    print("Comptant (pot trigar)...")
    dades = []
    recorrer(store, limit, dades)
    if not dades:
        print("Cap correu en aquest periode.")
        return

    per_dia = collections.Counter(d for d, _, _, _ in dades)
    laborables = {d: n for d, n in per_dia.items() if d.weekday() < 5}
    print("\n" + "=" * 54)
    print("  CORREUS PER DIA")
    print("=" * 54)
    for d in sorted(per_dia):
        marca = "" if d.weekday() < 5 else "  (cap de setmana)"
        print(f"  {d.isoformat()}  {DIES[d.weekday()]:10} {per_dia[d]:5}{marca}")

    tot = sum(per_dia.values())
    mitj_lab = sum(laborables.values()) / max(len(laborables), 1)
    print("\n" + "=" * 54)
    print(f"  Total del periode .............. {tot}")
    print(f"  Dies laborables ................ {len(laborables)}")
    print(f"  MITJANA per dia laborable ...... {mitj_lab:.0f}")
    print(f"  Projeccio mensual (22 dies) .... {mitj_lab*22:.0f}")
    print(f"  Projeccio ANUAL (250 dies) ..... {mitj_lab*250:.0f}")
    print("=" * 54)

    # Cost amb la projeccio anual
    N = mitj_lab * 250
    con_b = N * 0.62
    USD_EUR = 0.92
    cost = (((N - N*0.18) * 700 / 1e6) * 0.40 + ((N - N*0.18) * 120 / 1e6) * 1.60
            + (con_b * 1800 / 1e6) * 2.00 + (con_b * 420 / 1e6) * 8.00
            + (con_b * 900 / 1e6) * 0.40 + (con_b * 90 / 1e6) * 1.60) * USD_EUR
    print(f"\n  COST ESTIMAT (gpt-4.1 + gpt-4.1-mini)")
    print(f"    Anual ....... {cost:8.2f} €")
    print(f"    Mensual ..... {cost/12:8.2f} €")
    print(f"    Per correu .. {cost/max(N,1)*100:8.3f} centims")

    print("\n  PER FRANJA HORARIA (dies laborables)")
    hores = collections.Counter(h for d, h, _, _ in dades if d.weekday() < 5)
    for h in sorted(hores):
        barra = "#" * min(int(hores[h] / max(max(hores.values()), 1) * 40), 40)
        print(f"    {h:02d}h {hores[h]:5}  {barra}")

    if ruta_csv:
        with open(ruta_csv, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh, delimiter=";")
            w.writerow(["data", "dia", "correus", "laborable"])
            for d in sorted(per_dia):
                w.writerow([d.isoformat(), DIES[d.weekday()], per_dia[d],
                            "si" if d.weekday() < 5 else "no"])
        print(f"\n  CSV: {ruta_csv}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="BUSTIA", default=None)
    ap.add_argument("--dias", type=int, default=7)
    ap.add_argument("--csv", metavar="FITXER", default="")
    a = ap.parse_args()
    destino = a.outlook or cargar_escenario().get("agente", {}).get("buzon", "")
    if not destino:
        print('Cal --outlook "part-del-nom".'); sys.exit(1)
    main(destino, a.dias, a.csv)
    import gc; gc.collect()
    os._exit(0)
