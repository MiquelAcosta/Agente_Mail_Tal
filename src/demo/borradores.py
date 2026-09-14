# -*- coding: utf-8 -*-
"""Diagnostic: llista cada mail d'una carpeta amb el seu estat REAL de llegit/no-llegit
i les seves marques. La veritat interna d'Outlook, sense percepcions.
Us:  python src\\demo\\comprobar_leidos.py "integracion" "2 DIFICIL"
"""
import sys
import win32com.client

nom_buzon = sys.argv[1] if len(sys.argv) > 1 else "integracion"
nom_carpeta = sys.argv[2] if len(sys.argv) > 2 else "Tests Auto Cartel"

ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
stores = [ns.Folders.Item(i + 1) for i in range(ns.Folders.Count)]
store = next((s for s in stores if nom_buzon.lower() in s.Name.lower()), None)
if store is None:
    print("Buzon no trobat."); sys.exit(1)

def buscar(raiz, nombre):
    for i in range(raiz.Folders.Count):
        f = raiz.Folders.Item(i + 1)
        if f.Name.lower() == nombre.lower():
            return f
        sub = buscar(f, nombre)
        if sub:
            return sub
    return None

carpeta = buscar(store, nom_carpeta)
if carpeta is None:
    print(f"Carpeta '{nom_carpeta}' no trobada."); sys.exit(1)

print(f"\n{store.Name} > {carpeta.Name} — estat REAL de cada mail:\n")
print(f"{'ESTAT':10} {'MARCA':10} {'ASSUMPTE':50}")
print("-" * 75)
for msg in list(carpeta.Items):
    if getattr(msg, "Class", 0) != 43:
        continue
    estat = "NO LLEGIT" if msg.UnRead else "llegit"
    marca = "Agente" if "agente" in str(msg.Categories or "").lower() else "-"
    print(f"{estat:10} {marca:10} {str(msg.Subject or '')[:50]}")
print("\nRegla: l'agent NOMES processa els 'NO LLEGIT' sense marca 'Agente'.")
print("(Negreta a l'Outlook = NO LLEGIT. Mirar un mail al panell NO sempre el marca llegit!)")