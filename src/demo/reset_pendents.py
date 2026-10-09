"""reset_pendents.py — ALLIBERA DEL REGISTRE ELS CORREUS NO RESPOSTOS

Treu del log_borradores.sqlite les files de correus ANTERIORS (o iguals) a una
data, EXCLOENT els que ja estan contestats. Aixi l'agent els torna a processar
i refa els esborranys.

Que NO es resetea (per disseny):
  OMITIDO%  -> ja respost
  HILO%     -> el client ha tornat a escriure despres; aquell correu ja no es
               el viu del fil

Per defecte NOMES COMPTA. Cal --ejecutar per esborrar de debo, i sempre fa
copia de seguretat abans.

AIXO NO ESBORRA ELS ESBORRANYS DE L'OUTLOOK. Nomes allibera el registre.
Els esborranys dolents s'han de treure a ma de la carpeta Esborranys.

Us:
  python reset_pendents.py --fins 2026-09-30
  python reset_pendents.py --fins 2026-09-30 --ejecutar
"""

import argparse
import os
import shutil
import sqlite3
import sys
from collections import Counter
from datetime import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(AQUI, "log_borradores.sqlite")

# Motius que NO s'alliberen: el correu ja esta resolt.
JA_RESOLTS = ("OMITIDO%", "HILO%")


def condicio():
    """Clausula WHERE comuna: anteriors a la data i no resolts."""
    nots = " AND ".join(f"resultado NOT LIKE '{p}'" for p in JA_RESOLTS)
    return f"date(ts) <= ? AND {nots}"


def main(fins, executar):
    if not os.path.exists(RUTA):
        print(f"\n  No trobo el registre a:\n    {RUTA}")
        print("  Posa aquest script a la MATEIXA carpeta que log_borradores.sqlite")
        print("  (normalment src\\demo\\) i torna-ho a provar.\n")
        sys.exit(1)

    try:
        datetime.strptime(fins, "%Y-%m-%d")
    except ValueError:
        sys.exit("La data ha d'anar en format AAAA-MM-DD, p. ex. 2026-09-30")

    con = sqlite3.connect(RUTA)
    total = con.execute("SELECT COUNT(*) FROM borradores").fetchone()[0]
    print(f"\n  Registre: {total} files en total.")
    print(f"  Data de tall: fins al {fins} inclos.\n")

    # --- que hi ha en aquest periode, per motiu
    print("  TOT el que hi ha en aquest periode, per motiu:\n")
    files = con.execute(
        "SELECT resultado, COUNT(*) FROM borradores WHERE date(ts) <= ?"
        " GROUP BY 1 ORDER BY 2 DESC", (fins,)).fetchall()

    # Agrupa pel prefix abans del paréntesi, que si no surten centenars de linies.
    grups = Counter()
    for res, n in files:
        clau = (res or "(buit)").split(" (")[0].split(" [")[0]
        grups[clau] += n
    for clau, n in grups.most_common(12):
        resolt = any(clau.startswith(p.rstrip("%")) for p in JA_RESOLTS)
        marca = "  <- NO es tocara (ja resolt)" if resolt else ""
        print(f"    {n:6}  {clau}{marca}")

    # --- quants s'alliberarien
    quants = con.execute(
        f"SELECT COUNT(*) FROM borradores WHERE {condicio()}", (fins,)).fetchone()[0]
    print(f"\n  {'=' * 52}")
    print(f"  A RESETEJAR: {quants} correus")
    print(f"  {'=' * 52}")

    if quants == 0:
        print("\n  Res a fer.\n")
        con.close()
        return

    # --- mostra'n uns quants per comprovar
    print("\n  Mostra dels 8 mes recents que s'alliberarien:\n")
    for ts, rem, asu in con.execute(
            f"SELECT ts, remitente, asunto FROM borradores WHERE {condicio()}"
            " ORDER BY ts DESC LIMIT 8", (fins,)):
        print(f"    {str(ts)[:16]}  {str(rem)[:34]:34} {str(asu)[:40]}")

    if not executar:
        print("\n  MODE RECOMPTE: no s'ha tocat res.")
        print("  Per fer-ho de debo, repeteix la comanda amb  --ejecutar\n")
        con.close()
        return

    # --- copia de seguretat SEMPRE
    segell = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, f"log_borradores_BACKUP_{segell}.sqlite")
    con.close()
    shutil.copy(RUTA, copia)
    print(f"\n  Copia de seguretat: {os.path.basename(copia)}")

    con = sqlite3.connect(RUTA)
    n = con.execute(f"DELETE FROM borradores WHERE {condicio()}", (fins,)).rowcount
    con.commit()
    queden = con.execute("SELECT COUNT(*) FROM borradores").fetchone()[0]
    con.close()

    print(f"  Alliberats: {n} correus.   Queden al registre: {queden}\n")
    print("  SEGUENTS PASSOS (en aquest ordre):")
    print("   1. Treu de la carpeta Esborranys els esborranys dolents d'aquests")
    print("      correus. Aixo NO ho fa l'script: ordena per data i esborra'ls.")
    print("   2. Comprova que la finestra de sincronitzacio de l'Outlook arriba")
    print("      a aquesta data, o l'agent no podra tornar a llegir els correus.")
    print("   3. Prova amb pocs abans de deixar-ho corrent:")
    print("      python borradores.py --outlook \"info@recuperatudinero.com\"")
    print("         --carpeta \"1 FACIL\" --real --max 5 --dry --csv prova.csv\n")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--fins", metavar="AAAA-MM-DD", required=True,
                    help="data de tall, inclosa")
    ap.add_argument("--ejecutar", action="store_true",
                    help="esborra de debo (sense aixo, nomes compta)")
    a = ap.parse_args()
    main(a.fins, a.ejecutar)
