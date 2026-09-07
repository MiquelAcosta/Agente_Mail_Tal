# -*- coding: utf-8 -*-
"""Exporta els mails d'una carpeta de l'Outlook classic a fitxers .txt
en format bandeja_prueba (remitent / assumpte / ADJUNTO|NADA / cos).
NOMES LLEGEIX: no mou, no esborra, no envia res.

Preparacio: crea a l'Outlook una carpeta (p. ex. "PruebasAgente") i
arrossega-hi els mails que vulguis exportar.

Us:
  python src\\demo\\exportar_outlook.py "NOM_BUZON" "PruebasAgente"

Els .txt es guarden a src\\demo\\bandeja_prueba\\ amb noms outlook_001.txt, etc.
Requisits: Outlook classic obert en aquesta maquina + pip install pywin32
"""
import os, re, sys

try:
    import win32com.client
except ImportError:
    print("Falta pywin32. Instal·la'l amb:  pip install pywin32")
    sys.exit(1)

AQUI = os.path.dirname(os.path.abspath(__file__))
DESTI = os.path.join(AQUI, "bandeja_prueba")


def smtp_del_remitent(msg):
    try:
        if msg.SenderEmailType == "EX":
            ex = msg.Sender.GetExchangeUser()
            if ex:
                return ex.PrimarySmtpAddress
        return msg.SenderEmailAddress or "desconocido@exchange"
    except Exception:
        return "desconocido@exchange"


def buscar_carpeta(raiz, nombre):
    """Busca una carpeta pel nom (recursiu, per si esta dins d'una altra)."""
    for i in range(raiz.Folders.Count):
        f = raiz.Folders.Item(i + 1)
        if f.Name.lower() == nombre.lower():
            return f
        sub = buscar_carpeta(f, nombre)
        if sub:
            return sub
    return None


def main():
    if len(sys.argv) < 3:
        print('Us:  python src\\demo\\exportar_outlook.py "NOM_BUZON" "NOM_CARPETA"')
        sys.exit(1)
    nom_buzon, nom_carpeta = sys.argv[1], sys.argv[2]

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    store = next((s for s in stores if nom_buzon.lower() in s.Name.lower()), None)
    if store is None:
        print(f"Cap buzon conte '{nom_buzon}'. Visibles: " + ", ".join(s.Name for s in stores))
        sys.exit(1)

    carpeta = buscar_carpeta(store, nom_carpeta)
    if carpeta is None:
        print(f"No trobo la carpeta '{nom_carpeta}' dins de '{store.Name}'.")
        sys.exit(1)

    os.makedirs(DESTI, exist_ok=True)
    existentes = len([f for f in os.listdir(DESTI) if f.startswith("outlook_")])
    n = 0
    for msg in carpeta.Items:
        try:
            if msg.Class != 43:  # nomes mails
                continue
            adj_reales = [a for a in msg.Attachments
                          if not str(a.FileName).lower().startswith("image00")]
            cuerpo = str(msg.Body or "").strip()
            cuerpo = re.sub(r"\r\n", "\n", cuerpo)
            n += 1
            ruta = os.path.join(DESTI, f"outlook_{existentes + n:03d}.txt")
            with open(ruta, "w", encoding="utf-8") as f:
                f.write(smtp_del_remitent(msg).lower() + "\n")
                f.write(str(msg.Subject or "(sin asunto)").strip() + "\n")
                f.write(("ADJUNTO" if adj_reales else "NADA") + "\n")
                f.write(cuerpo + "\n")
            print(f"  exportado: {os.path.basename(ruta)}  <- {msg.Subject}")
        except Exception as e:
            print(f"  (mail saltado, no legible: {e})")
    print(f"\n{n} mails exportados a {DESTI}")
    print("Seguent pas:  python src\\demo\\sombra.py   (els processara com a nous)")


if __name__ == "__main__":
    main()
