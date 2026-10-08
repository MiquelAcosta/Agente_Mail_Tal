"""diag_autoria.py — PER QUE SURT EL QUE SURT

Dues preguntes:
  1) Quines paraules en MAJUSCULES disparen el senyal de "persona"?
     Si surten sempre les mateixes, es la signatura i el senyal no val.
  2) Per que nomes 19 dels 1713 esborranys registrats casen amb un enviat?

Nomes llegeix. No mou, no escriu, no envia.

Us:
  python diag_autoria.py --outlook "info@recuperatudinero.com" --dias 5
"""

import argparse
import os
import re
import sqlite3
import sys
from collections import Counter
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from informe_autoria import (carpeta_enviats, llegir_enviats, cos_net,
                             paraules, similitud, RE_CRIDAT, RUTA_REGISTRE)


def main(buzon, desde):
    carpeta = carpeta_enviats(buzon)
    enviats = llegir_enviats(carpeta, desde, 2000)
    print(f"\n  {len(enviats)} enviats al periode.\n")
    if not enviats:
        return

    # ---------------------------------------------------- 1. les majuscules
    print("  " + "=" * 58)
    print("  QUINES MAJUSCULES DISPAREN EL SENYAL")
    print("  " + "=" * 58)

    compta = Counter()
    mails_amb = 0
    for e in enviats:
        trobades = set(RE_CRIDAT.findall(cos_net(e["cos"])))
        if len(RE_CRIDAT.findall(cos_net(e["cos"]))) >= 3:
            mails_amb += 1
        compta.update(trobades)

    print(f"    {mails_amb} de {len(enviats)} mails tenen 3+ paraules en majuscules.\n")
    print("    Les 20 mes repetides (i en quants mails surten):\n")
    for paraula, n in compta.most_common(20):
        pct = 100 * n / len(enviats)
        marca = "  <-- SIGNATURA?" if pct > 70 else ""
        print(f"      {n:5}  ({pct:5.1f} %)  {paraula}{marca}")

    universals = [p for p, n in compta.items() if n / len(enviats) > 0.70]
    print()
    if universals:
        print(f"    VEREDICTE: {len(universals)} paraules surten a mes del 70 % dels")
        print("    mails. Son signatura o peu legal, NO escriptura humana.")
        print("    El senyal de majuscules NO serveix tal com esta.")
        print("\n    Recompte si s'ignoren aquestes paraules:")
        ignora = set(universals)
        nous = 0
        for e in enviats:
            resta = [w for w in RE_CRIDAT.findall(cos_net(e["cos"]))
                     if w not in ignora]
            if len(resta) >= 3:
                nous += 1
        print(f"      {nous} de {len(enviats)} mails ({100*nous/len(enviats):.1f} %)"
              " mantenen el senyal.")
    else:
        print("    VEREDICTE: cap paraula es universal. El senyal sembla bo.")

    # ------------------------------------------- 2. els esborranys que falten
    print("\n  " + "=" * 58)
    print("  PER QUE NO CASEN ELS ESBORRANYS")
    print("  " + "=" * 58)

    if not os.path.exists(RUTA_REGISTRE):
        print("    No hi ha log_borradores.sqlite en aquesta carpeta.")
        return

    con = sqlite3.connect(RUTA_REGISTRE)
    print("\n    Que va passar amb cada correu processat (camp 'resultado'):")
    for res, n in con.execute(
            "SELECT COALESCE(resultado,'(buit)'), COUNT(*) FROM borradores"
            " GROUP BY 1 ORDER BY 2 DESC").fetchall():
        print(f"      {n:6}  {res}")

    amb_text = con.execute(
        "SELECT COUNT(*) FROM borradores"
        " WHERE respuesta IS NOT NULL AND respuesta <> ''").fetchone()[0]
    print(f"\n    Amb text d'esborrany desat: {amb_text}")

    # Quants d'aquests esborranys son del periode que estem mirant
    files = con.execute(
        "SELECT ts, remitente, asunto, respuesta FROM borradores"
        " WHERE respuesta IS NOT NULL AND respuesta <> ''").fetchall()
    con.close()

    recents = []
    for ts, rem, asu, txt in files:
        try:
            d = datetime.fromisoformat(str(ts)[:19])
        except Exception:
            d = None
        if d is None or d >= desde:
            recents.append({"rem": (rem or "").lower(), "asunto": asu or "",
                            "txt": txt or ""})
    print(f"    D'aquests, del periode mirat: {len(recents)}")

    # Quants destinataris del registre apareixen als enviats
    destins = " ".join((e["para"] or "").lower() for e in enviats)
    per_remitent = sum(1 for r in recents if r["rem"] and r["rem"] in destins)
    print(f"\n    Esborranys el destinatari dels quals SI apareix als enviats:"
          f" {per_remitent}")
    print(f"    (si aquest numero es alt pero les coincidencies son poques,")
    print(f"     el problema es l'assumpte; si es baix, es que no s'han enviat)")

    # Ara afluixant el criteri: nomes text, sense mirar remitent ni assumpte
    bosses_env = [paraules(cos_net(e["cos"])) for e in enviats]
    casen_text = 0
    for r in recents:
        br = paraules(r["txt"])
        if any(similitud(br, be) >= 0.60 for be in bosses_env):
            casen_text += 1
    print(f"\n    Esborranys el TEXT dels quals apareix a algun enviat"
          f" (sense mirar remitent ni assumpte): {casen_text}")
    print(f"    Amb el criteri estricte de l'informe n'hi havia 19.")
    if casen_text > 40:
        print("\n    => L'encreuament es massa estricte. Cal afluixar-lo.")
    else:
        print("\n    => L'equip no esta enviant els esborranys. El numero es real.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", default=None)
    ap.add_argument("--dias", type=int, default=5)
    a = ap.parse_args()
    main(a.outlook, datetime.now() - timedelta(days=a.dias))
