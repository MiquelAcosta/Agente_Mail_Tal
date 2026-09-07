# -*- coding: utf-8 -*-
"""Prova de contacte amb l'Outlook d'escriptori (via COM). NOMES LLEGEIX.

Requisits: Outlook classic instal·lat i configurat en aquesta maquina,
i 'pip install pywin32'.

Us:
  python src\\demo\\listar_outlook.py            -> llista els buzons/comptes visibles
  python src\\demo\\listar_outlook.py "NOM"      -> mostra els ultims 5 mails de la
                                                   safata d'entrada del buzon que
                                                   contingui NOM (part del nom val)
"""
import sys

try:
    import win32com.client
except ImportError:
    print("Falta pywin32. Instal·la'l amb:  pip install pywin32")
    sys.exit(1)


def smtp_del_remitent(msg):
    """Treu l'adreça SMTP real (Exchange amaga l'adreça darrere d'un format intern)."""
    try:
        if msg.SenderEmailType == "EX":
            ex = msg.Sender.GetExchangeUser()
            if ex:
                return ex.PrimarySmtpAddress
        return msg.SenderEmailAddress or "(desconegut)"
    except Exception:
        return "(no llegible)"


def main():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
    print("Buzons/comptes visibles en aquest Outlook:")
    for s in stores:
        print(f"  - {s.Name}")

    if len(sys.argv) < 2:
        print('\nAra executa:  python src\\demo\\listar_outlook.py "part del nom del buzon"')
        return

    objectiu = sys.argv[1].lower()
    store = next((s for s in stores if objectiu in s.Name.lower()), None)
    if store is None:
        print(f"\nCap buzon conte '{sys.argv[1]}'. Copia el nom exacte de la llista de dalt.")
        return

    # Buscar la safata d'entrada dins del buzon triat
    inbox = None
    for i in range(store.Folders.Count):
        f = store.Folders.Item(i + 1)
        if f.Name.lower() in ("bandeja de entrada", "inbox", "safata d'entrada"):
            inbox = f
            break
    if inbox is None:
        print(f"\nNo trobo la safata d'entrada dins de '{store.Name}'. Carpetes disponibles:")
        for i in range(store.Folders.Count):
            print(f"  - {store.Folders.Item(i + 1).Name}")
        return

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # mes recents primer
    print(f"\nUltims 5 mails de '{store.Name}' > '{inbox.Name}'  ({inbox.Items.Count} en total):")
    print("-" * 70)
    n = 0
    for msg in items:
        try:
            if msg.Class != 43:  # nomes MailItem (descarta convocatories, informes...)
                continue
            adj_reals = sum(1 for a in msg.Attachments
                            if not str(a.FileName).lower().startswith("image00"))
            print(f"[{msg.ReceivedTime}] {smtp_del_remitent(msg)}")
            print(f"    Assumpte: {msg.Subject}")
            print(f"    Adjunts (aprox., sense logos): {adj_reals}")
            print("-" * 70)
            n += 1
            if n >= 5:
                break
        except Exception as e:
            print(f"    (mail no llegible: {e})")
    if n == 0:
        print("(cap mail llegible a la safata)")
    print("\nContacte OK: l'agent pot llegir aquest buzon. Seguent pas: la fase 1 en ombra.")


if __name__ == "__main__":
    main()
