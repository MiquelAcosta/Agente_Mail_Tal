# -*- coding: utf-8 -*-
"""DIAGNOSTIC DEL PACK — nomes MIRA. No mou, no marca, no crea esborranys.

Ensenya, mail per mail de la carpeta: quin client li assigna, quina data en
treu (i de quina propietat), i quin mail sortiria triat com "el mes nou" de
cada client. Si el borradores respon el mail equivocat, aixo diu exactament on
es trenca.

Us (Outlook obert):
  python src\\demo\\diagnostico_pack.py --outlook "NOM" --carpeta "1 FACIL"
"""
import os, sys, argparse

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)

from borradores import (cargar_escenario, componentes_fecha, _componentes_de,
                        fecha_texto, id_estable, es_reenviador,
                        extraer_cliente_de_reenvio, te_marca_agente)


def clave_cliente(m, reenv):
    try:
        snd = str(getattr(m, "SenderEmailAddress", "") or "").lower()
        if "@" not in snd or snd.startswith("/o="):
            try:
                snd = str(m.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001F")).lower()
            except Exception:
                pass
        if es_reenviador(snd, reenv):
            return extraer_cliente_de_reenvio(str(getattr(m, "Body", "") or "")[:4000],
                                              internos=reenv), "reenviament"
        return (snd if "@" in snd else None), "remitent directe"
    except Exception as e:
        return None, f"error: {str(e)[:40]}"


def main(buzon, nombre_carpeta):
    import win32com.client
    esc = cargar_escenario()
    reenv = esc.get("reenviadores", [])
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = next((s for s in stores if buzon.lower() in s.Name.lower()), None)
    if store is None:
        print(f"Cap buzon conte '{buzon}'."); return

    def buscar(raiz, nombre):
        for i in range(raiz.Folders.Count):
            f = raiz.Folders.Item(i + 1)
            if f.Name.lower() == nombre.lower():
                return f
            sub = buscar(f, nombre)
            if sub:
                return sub
        return None

    carpeta = buscar(store, nombre_carpeta)
    if carpeta is None:
        print(f"No trobo '{nombre_carpeta}' dins de '{store.Name}'."); return

    mails = [m for m in list(carpeta.Items) if getattr(m, "Class", 0) == 43]
    print("=" * 74)
    print(f" DIAGNOSTIC: {store.Name} > {carpeta.Name}  ({len(mails)} mails)")
    print(" NOMES LECTURA: no es mou ni es crea res.")
    print("=" * 74)

    grupos = {}
    for i, m in enumerate(mails, 1):
        asunto = str(getattr(m, "Subject", "") or "")[:44]
        cli, via = clave_cliente(m, reenv)
        marca = te_marca_agente(m)
        print(f"\n[{i}] {asunto}")
        print(f"    client    : {cli or '(NO IDENTIFICAT -> no s agrupa)'}   [{via}]")
        for origen in ("ReceivedTime", "SentOn", "CreationTime"):
            v = getattr(m, origen, None)
            comp = _componentes_de(v)
            print(f"    {origen:13}: valor={str(v)[:30]:30} tipus={type(v).__name__:12}"
                  f" -> {fecha_texto(comp)}")
        fecha = componentes_fecha(m)
        print(f"    DATA USADA: {fecha_texto(fecha)}")
        print(f"    marca Agente: {'SI (quedaria fora del pack)' if marca else 'no'}")
        print(f"    id_estable: {id_estable(m)[:60]}")
        if cli and fecha and not marca:
            grupos.setdefault(cli, []).append((fecha, asunto))

    print("\n" + "=" * 74)
    print(" PACKS QUE ES FORMARIEN")
    print("=" * 74)
    hi_ha = False
    for cli, items in grupos.items():
        if len(items) < 2:
            continue
        hi_ha = True
        items.sort(key=lambda p: p[0])
        print(f"\n  {cli}: {len(items)} mails")
        for f, a in items[:-1]:
            print(f"     previ     {fecha_texto(f)}  {a}")
        print(f"     RESPON -> {fecha_texto(items[-1][0])}  {items[-1][1]}   <= EL MES NOU")
    if not hi_ha:
        print("\n  Cap pack. Cap client amb 2+ mails datats i sense marca en aquesta carpeta.")
        print("  Causes tipiques: els mails son a carpetes diferents, algun ja porta la")
        print("  marca 'Agente', o el client no s'identifica igual als dos mails.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="NOM", required=True)
    ap.add_argument("--carpeta", metavar="CARPETA", default="1 FACIL")
    a = ap.parse_args()
    main(a.outlook, a.carpeta)
    import gc; gc.collect()
    os._exit(0)
