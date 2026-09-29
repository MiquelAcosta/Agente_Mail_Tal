# -*- coding: utf-8 -*-
"""EXTREURE PREGUNTES I RESPOSTES REALS — nomes LLEGEIX, no toca res.

Recorre la safata d'entrada, i per a cada correu de client busca a ELEMENTS
ENVIATS la resposta que li va donar l'equip (mateixa conversa). Escriu un CSV
amb la parella pregunta/resposta.

NO crida cap model: cost zero.
NO crea esborranys, no mou correus, no marca res.

Us:
  python src\\demo\\extraer_qa.py --dias 30 --max 300 --csv qa.csv
  python src\\demo\\extraer_qa.py --carpeta "Bandeja de entrada" --dias 60
"""
import os, sys, csv, argparse, re
from datetime import datetime, timedelta

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AQUI, "..", "triaje"))
sys.path.insert(0, AQUI)

from borradores import (cargar_escenario, escoger_buzon, remitente_smtp,
                        componentes_fecha, fecha_texto, cuerpo_nuevo,
                        es_reenviador, extraer_cliente_de_reenvio)

NOMS_ENVIATS = {"elementos enviados", "sent items", "elements enviats",
                "correo enviado", "enviados"}


def buscar_carpeta(raiz, nombres):
    """Busca una carpeta pel nom (accepta diverses variants)."""
    try:
        for i in range(raiz.Folders.Count):
            f = raiz.Folders.Item(i + 1)
            if str(f.Name).strip().lower() in nombres:
                return f
            sub = buscar_carpeta(f, nombres)
            if sub:
                return sub
    except Exception:
        pass
    return None


def net(t, limit=1200):
    """Text d'una linia, sense el fil citat."""
    return " ".join(str(cuerpo_nuevo(t or "")).split())[:limit]


def main(buzon, nombre_carpeta, dias, maxim, ruta_csv):
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = escoger_buzon(stores, buzon)
    if store is None:
        print(f"Cap buzon conte '{buzon}'. Visibles:")
        for s in stores:
            print("   -", s.Name)
        return

    entrada = buscar_carpeta(store, {nombre_carpeta.strip().lower()})
    if entrada is None:
        print(f"No trobo '{nombre_carpeta}' dins de '{store.Name}'.")
        return
    enviats = buscar_carpeta(store, NOMS_ENVIATS)
    print(f"Safata:  {store.Name} > {entrada.Name}")
    print(f"Enviats: {enviats.Name if enviats else '(NO TROBADA — no hi haura respostes)'}")

    # Index de respostes per conversa
    per_conversa = {}
    if enviats is not None:
        items = enviats.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        n = 0
        m = items.GetFirst()
        while m is not None and n < 5000:
            n += 1
            try:
                if getattr(m, "Class", 0) == 43:
                    cid = str(getattr(m, "ConversationID", "") or "")
                    if cid:
                        per_conversa.setdefault(cid, []).append(m)
            except Exception:
                pass
            m = items.GetNext()
        print(f"Enviats indexats: {n} correus en {len(per_conversa)} converses")

    limit = datetime.now() - timedelta(days=int(dias)) if dias else None
    reenv = cargar_escenario().get("reenviadores", [])
    items = entrada.Items
    try:
        items.Sort("[ReceivedTime]", False)     # dels mes VELLS als mes nous
    except Exception:
        pass

    files, vistos, amb_resposta = [], 0, 0
    m = items.GetFirst()
    while m is not None and len(files) < maxim:
        try:
            if getattr(m, "Class", 0) != 43:
                m = items.GetNext(); continue
            vistos += 1
            f = componentes_fecha(m)
            if limit and f:
                d = datetime(*f)
                if d < limit:
                    m = items.GetNext(); continue
            rem = remitente_smtp(m) or ""
            cos = str(getattr(m, "Body", "") or "")
            if es_reenviador(rem, reenv):
                ext = extraer_cliente_de_reenvio(cos[:4000], internos=reenv)
                if ext:
                    rem = ext
            if not rem or "@" not in rem:
                m = items.GetNext(); continue

            resposta, data_resp = "", ""
            cid = str(getattr(m, "ConversationID", "") or "")
            for r in per_conversa.get(cid, []):
                fr = componentes_fecha(r)
                if f and fr and fr > f:
                    resposta = net(str(getattr(r, "Body", "") or ""), 2000)
                    data_resp = fecha_texto(fr)
                    break
            if resposta:
                amb_resposta += 1

            files.append({
                "fecha": fecha_texto(f),
                "cliente": rem,
                "asunto": str(getattr(m, "Subject", "") or "")[:150],
                "pregunta": net(cos),
                "fecha_respuesta": data_resp,
                "respuesta_equipo": resposta,
            })
            if len(files) % 25 == 0:
                print(f"   ... {len(files)} parelles ({amb_resposta} amb resposta)")
        except Exception:
            pass
        m = items.GetNext()

    if not files:
        print("\nCap correu. Proveu amb mes --dias o una altra carpeta.")
        return
    cols = ["fecha", "cliente", "asunto", "pregunta", "fecha_respuesta", "respuesta_equipo"]
    with open(ruta_csv, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, delimiter=";")
        w.writeheader()
        for r in files:
            w.writerow(r)
    print(f"\n{'='*58}")
    print(f"  Correus mirats ............ {vistos}")
    print(f"  Parelles desades .......... {len(files)}")
    print(f"  Amb resposta de l'equip ... {amb_resposta}"
          f"  ({amb_resposta*100//max(len(files),1)}%)")
    print(f"  CSV ....................... {ruta_csv}")
    print("="*58)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="BUSTIA", default=None)
    ap.add_argument("--carpeta", metavar="CARPETA", default="Bandeja de entrada")
    ap.add_argument("--dias", type=int, default=45,
                    help="nomes correus dels ultims N dies (0 = tots)")
    ap.add_argument("--max", type=int, default=300)
    ap.add_argument("--csv", metavar="FITXER", default="qa.csv")
    a = ap.parse_args()
    destino = a.outlook or cargar_escenario().get("agente", {}).get("buzon", "")
    if not destino:
        print('Cal --outlook "part-del-nom-de-la-bustia".'); sys.exit(1)
    main(destino, a.carpeta, a.dias, a.max, a.csv)
    import gc; gc.collect()
    os._exit(0)
