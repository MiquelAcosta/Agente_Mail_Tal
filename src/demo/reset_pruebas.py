# -*- coding: utf-8 -*-
"""REINICIAR PROVES — deixa uns mails com si l'agent no els hagues vist mai.

Un mail no es reprocessa per DUES raons independents, i cal desfer les dues:
  1. Esta apuntat al registre log_borradores.sqlite  -> s'esborra la fila
  2. Porta l'etiqueta 'Agente' al propi mail         -> es treu l'etiqueta

Opcionalment torna els mails de la subcarpeta MULTIPLES al calaix del qual van
sortir, per poder repetir la prova d'agrupacio amb els mateixos mails.

NO esborra cap mail. Nomes treu etiquetes, files de registre i, si es demana,
retorna els mails de MULTIPLES a la carpeta pare.

Us:
  python src\\demo\\reset_pruebas.py --carpeta "1 FACIL" --dry
  python src\\demo\\reset_pruebas.py --carpeta "1 FACIL"
  python src\\demo\\reset_pruebas.py --carpeta "1 FACIL" --devolver-multiples
"""
import os, sys, argparse, sqlite3

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)

from borradores import (cargar_escenario, id_estable, CATEGORIA_AGENTE,
                        CARPETA_MULTIPLES)

DB = os.path.join(AQUI, "log_borradores.sqlite")


def escoger_buzon(stores, nombre):
    """Tria la bustia. PRIORITAT: nom exacte -> comenca per -> conte.

    Sense aixo, 'info@recuperatudinero.com' tambe coincideix amb
    'Online Archive - info@recuperatudinero.com', i com que l'arxiu surt abans a
    la llista guanyava ell i no es trobava cap calaix. Si queda ambigu, s'avisa."""
    n = (nombre or "").strip().lower()
    noms = [str(getattr(s, "Name", "") or "") for s in stores]
    for prova in (lambda x: x.lower() == n, lambda x: x.lower().startswith(n)):
        tri = [s for s, x in zip(stores, noms) if prova(x)]
        if tri:
            return tri[0]
    conte = [s for s, x in zip(stores, noms) if n in x.lower()]
    if not conte:
        return None
    if len(conte) > 1:
        print(f"    (AVIS: {len(conte)} busties contenen '{nombre}': "
              + ", ".join(str(getattr(c, "Name", "")) for c in conte) + ")")
        print(f"    (s'usa '{getattr(conte[0], 'Name', '')}' — poseu el nom exacte)")
    return conte[0]


def buscar(raiz, nombre):
    for i in range(raiz.Folders.Count):
        f = raiz.Folders.Item(i + 1)
        if f.Name.lower() == nombre.lower():
            return f
        sub = buscar(f, nombre)
        if sub:
            return sub
    return None


def limpiar(msgs, con, dry, etiqueta):
    quitadas = borradas = 0
    for m in msgs:
        asunto = str(getattr(m, "Subject", "") or "")[:44]
        mid = id_estable(m)
        cats = str(getattr(m, "Categories", "") or "")
        tiene = etiqueta.lower() in cats.lower()
        fila = con.execute("SELECT COUNT(*) FROM borradores WHERE mail_id=?",
                           (mid,)).fetchone()[0]
        if not tiene and not fila:
            continue
        print(f"  {asunto:46} etiqueta={'SI' if tiene else 'no'} registre={fila}")
        if dry:
            continue
        if tiene:
            nuevas = [c for c in cats.split(";") if c.strip().lower() != etiqueta.lower()]
            try:
                m.Categories = "; ".join(c.strip() for c in nuevas if c.strip())
                m.UnRead = True
                m.Save()
                quitadas += 1
            except Exception as e:
                print(f"      (no s'ha pogut treure l'etiqueta: {str(e)[:60]})")
        if fila:
            con.execute("DELETE FROM borradores WHERE mail_id=?", (mid,))
            borradas += 1
    if not dry:
        con.commit()
    return quitadas, borradas


def main(buzon, nombre_carpeta, dry, devolver):
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = escoger_buzon(stores, buzon)
    if store is None:
        print(f"Cap buzon conte '{buzon}'. Visibles:")
        for s_ in stores:
            print("   -", s_.Name)
        return
    carpeta = buscar(store, nombre_carpeta)
    if carpeta is None:
        print(f"No trobo '{nombre_carpeta}' dins de '{store.Name}'.")
        return

    con = sqlite3.connect(DB)
    print("=" * 66)
    print(f" REINICIAR PROVES: {store.Name} > {carpeta.Name}")
    print(f" {'MODE DRY: no es canvia res' if dry else 'CANVIS REALS'}")
    print("=" * 66)

    # Tornar els mails apartats a MULTIPLES
    devueltos = 0
    mult = None
    for i in range(carpeta.Folders.Count):
        f = carpeta.Folders.Item(i + 1)
        if f.Name.lower() == CARPETA_MULTIPLES.lower():
            mult = f
            break
    if mult is not None:
        en_mult = [m for m in list(mult.Items) if getattr(m, "Class", 0) == 43]
        print(f"\n{CARPETA_MULTIPLES}: {len(en_mult)} mails")
        if devolver and not dry:
            for m in list(en_mult):
                try:
                    m.Move(carpeta)
                    devueltos += 1
                except Exception as e:
                    print(f"   (no s'ha pogut tornar: {str(e)[:50]})")
            print(f"   tornats al calaix: {devueltos}")
        elif devolver:
            print("   (amb --devolver-multiples es tornarien al calaix)")
        else:
            limpiar(en_mult, con, dry, CATEGORIA_AGENTE)

    mails = [m for m in list(carpeta.Items) if getattr(m, "Class", 0) == 43]
    print(f"\n{carpeta.Name}: {len(mails)} mails")
    q, b = limpiar(mails, con, dry, CATEGORIA_AGENTE)

    print("\n" + "-" * 66)
    if dry:
        print(" DRY: no s'ha canviat res. Treu --dry per aplicar-ho.")
    else:
        print(f" Etiquetes tretes: {q}   Files de registre esborrades: {b}"
              f"   Tornats de {CARPETA_MULTIPLES}: {devueltos}")
        print(" Aquests mails ja es poden tornar a processar.")
    con.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="BUSTIA", default=None)
    ap.add_argument("--carpeta", metavar="CARPETA", default="1 FACIL")
    ap.add_argument("--dry", action="store_true", help="ensenya que faria, sense tocar res")
    ap.add_argument("--devolver-multiples", dest="devolver", action="store_true",
                    help="torna els mails de MULTIPLES al calaix pare")
    a = ap.parse_args()
    destino = a.outlook or cargar_escenario().get("agente", {}).get("buzon", "")
    if not destino:
        print('Cal --outlook "part-del-nom-de-la-bustia".')
        sys.exit(1)
    main(destino, a.carpeta, a.dry, a.devolver)
    import gc
    gc.collect()
    os._exit(0)