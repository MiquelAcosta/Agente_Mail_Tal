"""reset_pendents.py — ALLIBERA ELS CORREUS NO RESPOSTOS DELS CALAIXOS

Recorre NOMES les carpetes de l'automatitzacio (no la safata d'entrada), agafa
els correus REBUTS fins a una data i que ENCARA NO TENEN RESPOSTA, i treu la
seva fila del log_borradores.sqlite perque l'agent els torni a processar.

Com sap si un correu esta respost:
  Llegeix la propietat interna d'Outlook 'last verb' (102 = respost,
  103 = respost a tots). Aixo es l'estat REAL d'ara mateix, aixi que detecta
  tambe les respostes que l'equip ha enviat a ma. El 104 (reenviat) no compta.

Com casa un correu amb el registre:
  Pel Message-ID d'internet, que es el que guarda el registre i sobreviu als
  moviments de carpeta (l'EntryID no).

Per defecte NOMES COMPTA. Cal --ejecutar per esborrar, i sempre fa copia abans.

AIXO NO ESBORRA ELS ESBORRANYS DE L'OUTLOOK: nomes allibera el registre.

Us:
  python reset_pendents.py --outlook "info@recuperatudinero.com" --fins 2026-09-30
  python reset_pendents.py --outlook "info@recuperatudinero.com" --fins 2026-09-30 --ejecutar
  python reset_pendents.py --outlook "..." --fins 2026-09-30 --carpetas "1 FACIL"
"""

import argparse
import os
import shutil
import sqlite3
import sys
from collections import Counter
from datetime import datetime

try:
    import win32com.client
except ImportError:
    sys.exit("Falta pywin32:  pip install pywin32")

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(AQUI, "log_borradores.sqlite")

# Nomes els calaixos de l'automatitzacio. La safata d'entrada NO hi es.
CALAIXOS = ["1 FACIL", "2 DIFICIL", "3 DESISTIMIENTO"]

# Les mateixes propietats que fa servir borradores.py
PROP_MSG_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PROP_LAST_VERB = "http://schemas.microsoft.com/mapi/proptag/0x10810003"


def escoger_buzon(stores, nombre):
    """Igual que a borradores.py: nom exacte -> comenca per -> conte.
    Sense aixo, 'Online Archive - info@...' guanya i no es troba cap calaix."""
    n = (nombre or "").strip().lower()
    noms = [str(getattr(s, "Name", "") or "") for s in stores]
    for prova in (lambda x: x.lower() == n, lambda x: x.lower().startswith(n)):
        tri = [s for s, x in zip(stores, noms) if prova(x)]
        if tri:
            return tri[0]
    conte = [s for s, x in zip(stores, noms) if n in x.lower()]
    return conte[0] if conte else None


def id_estable(msg):
    """El Message-ID d'internet: sobreviu als moviments de carpeta."""
    try:
        mid = msg.PropertyAccessor.GetProperty(PROP_MSG_ID)
        if mid:
            return "msgid:" + str(mid).strip()
    except Exception:
        pass
    return "outlook:" + str(msg.EntryID)


def ya_respondido(msg):
    """True si el mail ja te resposta enviada. 104 (reenviat) NO compta."""
    try:
        return msg.PropertyAccessor.GetProperty(PROP_LAST_VERB) in (102, 103)
    except Exception:
        return False


def subcarpeta(arrel, nom, nivell=0):
    """Busca la carpeta a QUALSEVOL nivell: els calaixos solen penjar de la
    safata d'entrada, no de l'arrel de la bustia."""
    if nivell > 4:
        return None
    try:
        for i in range(arrel.Folders.Count):
            f = arrel.Folders.Item(i + 1)
            if str(f.Name).strip().lower() == nom.strip().lower():
                return f
            trobat = subcarpeta(f, nom, nivell + 1)
            if trobat is not None:
                return trobat
    except Exception:
        pass
    return None


def arbre(arrel, nivell=0, maxim=3):
    """Imprimeix l'arbre de carpetes, per saber on son els calaixos."""
    if nivell > maxim:
        return
    try:
        for i in range(arrel.Folders.Count):
            f = arrel.Folders.Item(i + 1)
            try:
                n = f.Items.Count
            except Exception:
                n = "?"
            print(f"    {'  ' * nivell}- {f.Name}  ({n})")
            arbre(f, nivell + 1, maxim)
    except Exception:
        pass


def recorrer(carpeta, fins):
    """Torna (pendents, respostos, posteriors, errors) per a una carpeta."""
    pendents, respostos, posteriors, errors = [], 0, 0, 0
    try:
        items = carpeta.Items
    except Exception:
        return pendents, respostos, posteriors, 1
    for it in items:
        try:
            rebut = getattr(it, "ReceivedTime", None)
            if rebut is not None and rebut.replace(tzinfo=None).date() > fins:
                posteriors += 1
                continue
            if ya_respondido(it):
                respostos += 1
                continue
            pendents.append({
                "id": id_estable(it),
                "data": rebut.replace(tzinfo=None) if rebut else None,
                "de": str(getattr(it, "SenderEmailAddress", "") or "")[:40],
                "asunto": str(getattr(it, "Subject", "") or "")[:46],
            })
        except Exception:
            errors += 1
    return pendents, respostos, posteriors, errors


def main(buzon, fins_txt, carpetes, executar):
    if not os.path.exists(RUTA):
        sys.exit(f"\n  No trobo el registre a:\n    {RUTA}\n"
                 "  Posa aquest script al costat del log_borradores.sqlite.\n")
    try:
        fins = datetime.strptime(fins_txt, "%Y-%m-%d").date()
    except ValueError:
        sys.exit("La data ha d'anar en format AAAA-MM-DD, p. ex. 2026-09-30")

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = escoger_buzon(stores, buzon)
    if store is None:
        print("  Magatzems disponibles:")
        for s in stores:
            print(f"     - {s.Name}")
        sys.exit(f"\n  No trobo cap bustia que contingui '{buzon}'")

    print(f"\n  Bustia: {store.Name}")
    print(f"  Rebuts fins al {fins} inclos. NOMES correus sense resposta.\n")

    tots, resum = [], []
    for nom in carpetes:
        c = subcarpeta(store, nom)
        if c is None:
            print(f"    {nom:18}  NO EXISTEIX en aquesta bustia")
            continue
        pend, resp, post, err = recorrer(c, fins)
        tots.extend(pend)
        resum.append((nom, len(pend), resp, post, err))
        extra = f"   ({err} no llegits)" if err else ""
        print(f"    {nom:18}  {len(pend):5} pendents   "
              f"{resp:5} ja respostos   {post:5} posteriors a la data{extra}")

    if not resum:
        print("\n  No s'ha trobat cap dels calaixos demanats."
              "\n  Aquestes son les carpetes de la bustia:\n")
        arbre(store)
        print("\n  Passa el nom exacte amb --carpetas.\n")
        return

    if not tots:
        print("\n  Cap correu compleix el filtre. Res a fer.\n")
        return

    con = sqlite3.connect(RUTA)
    dins = [p for p in tots
            if con.execute("SELECT 1 FROM borradores WHERE mail_id=?",
                           (p["id"],)).fetchone()]
    print(f"\n  {'=' * 56}")
    print(f"  Pendents trobats als calaixos: {len(tots)}")
    print(f"  D'aquests, amb fila al registre: {len(dins)}  <- els que s'alliberen")
    print(f"  {'=' * 56}")

    if len(tots) - len(dins):
        print(f"\n  ({len(tots) - len(dins)} no tenen fila al registre:"
              " l'agent encara no els havia vist. Ja es processaran sols.)")

    if dins:
        print("\n  Per motiu registrat:\n")
        motius = Counter()
        for p in dins:
            r = con.execute("SELECT resultado FROM borradores WHERE mail_id=?",
                            (p["id"],)).fetchone()[0] or "(buit)"
            motius[r.split(" (")[0].split(" [")[0]] += 1
        for m, n in motius.most_common(10):
            print(f"    {n:5}  {m}")

        print("\n  Mostra dels 8 mes recents:\n")
        for p in sorted(dins, key=lambda x: x["data"] or datetime.min,
                        reverse=True)[:8]:
            d = p["data"].strftime("%Y-%m-%d %H:%M") if p["data"] else "?"
            print(f"    {d}  {p['de']:40} {p['asunto']}")

    if not executar:
        print("\n  MODE RECOMPTE: no s'ha tocat res.")
        print("  Per fer-ho de debo, repeteix amb  --ejecutar\n")
        con.close()
        return
    if not dins:
        print("\n  Res per alliberar al registre.\n")
        con.close()
        return

    segell = datetime.now().strftime("%Y%m%d_%H%M%S")
    copia = os.path.join(AQUI, f"log_borradores_BACKUP_{segell}.sqlite")
    con.close()
    shutil.copy(RUTA, copia)
    print(f"\n  Copia de seguretat: {os.path.basename(copia)}")

    con = sqlite3.connect(RUTA)
    n = 0
    for p in dins:
        n += con.execute("DELETE FROM borradores WHERE mail_id=?",
                         (p["id"],)).rowcount
    con.commit()
    queden = con.execute("SELECT COUNT(*) FROM borradores").fetchone()[0]
    con.close()

    print(f"  Alliberats: {n}.   Queden al registre: {queden}\n")
    print("  SEGUENTS PASSOS:")
    print("   1. Treu de la carpeta Esborranys els esborranys dolents d'aquests")
    print("      correus. Aixo NO ho fa l'script.")
    print("   2. Comprova que la finestra de sincronitzacio arriba a aquesta data.")
    print("   3. Prova amb pocs abans de deixar-ho corrent (--max 5 --dry).\n")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="BUSTIA", required=True)
    ap.add_argument("--fins", metavar="AAAA-MM-DD", required=True)
    ap.add_argument("--carpetas", metavar="LLISTA", default=",".join(CALAIXOS),
                    help="calaixos separats per comes")
    ap.add_argument("--ejecutar", action="store_true",
                    help="esborra de debo (sense aixo, nomes compta)")
    a = ap.parse_args()
    main(a.outlook, a.fins, [c.strip() for c in a.carpetas.split(",") if c.strip()],
         a.ejecutar)