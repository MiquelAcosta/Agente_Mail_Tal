# -*- coding: utf-8 -*-
"""L'AGENT EN MARXA — el director d'orquestra del pilot.

Executa el cicle complet en bucle, cada X minuts, indefinidament:
  1. organizar.py  -> reparteix el correu nou de la porta d'entrada (FACIL/DIFICIL/DESCARTES)
  2. borradores.py -> crea esborranys per als mails nous de 1 FACIL i 2 DIFICIL

Els registres SQLite garanteixen que res es processa dues vegades: cada cicle
nomes toca el correu NOU. Si una passada falla, s'apunta i el cicle seguent
ho reintenta. Aturar: Ctrl+C (acaba el pas en curs i surt net).

REQUISITS perque giri sol:
  - La sessio de Windows oberta (pots desconnectar el RDP, pero no tancar sessio)
  - L'Outlook classic obert en aquesta sessio
  - Ollama corrent (ollama serve) amb OLLAMA_KEEP_ALIVE llarg

Us:
  python src\\demo\\agente.py                 (bucle infinit)
  python src\\demo\\agente.py --una-passada   (un sol cicle i para: per provar)

Configuracio a escenario.json (opcional, aquests son els valors per defecte):
  "agente": {"buzon": "integracion", "carpeta_entrada": "Tests Auto Cartel",
              "intervalo_minutos": 5}
"""
import json, os, subprocess, sys, time, argparse
from datetime import datetime

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_ESC = os.path.join(AQUI, "escenario.json")


def config():
    ag = {}
    try:
        ag = json.load(open(RUTA_ESC, encoding="utf-8-sig")).get("agente", {})
    except Exception:
        pass
    return {
        "buzon": ag.get("buzon", "integracion"),
        "entrada": ag.get("carpeta_entrada", "Tests Auto Cartel"),
        "intervalo": max(1, int(ag.get("intervalo_minutos", 5))),
        "hora_inicio": int(ag.get("hora_inicio", 8)),
        "hora_fin": int(ag.get("hora_fin", 19)),
        "fines_de_semana": bool(ag.get("fines_de_semana", True)),
    }


def en_horario(cfg):
    ara = datetime.now()
    if not cfg["fines_de_semana"] and ara.weekday() >= 5:
        return False
    return cfg["hora_inicio"] <= ara.hour < cfg["hora_fin"]


def dormir_fins_horari(cfg):
    """Dorm en trams d'1 minut fins que torni l'horari (Ctrl+C segueix funcionant)."""
    msg = ("[" + datetime.now().strftime("%H:%M") + "] Fora d'horari ("
           + f"{cfg['hora_inicio']:02d}:00-{cfg['hora_fin']:02d}:00"
           + ("" if cfg["fines_de_semana"] else ", laborables")
           + "). L'agent descansa fins l'hora d'inici...")
    print(msg, flush=True)
    while not en_horario(cfg):
        time.sleep(60)
    print(f"[{datetime.now().strftime('%H:%M')}] Bon dia! L'agent torna a la feina.")


def paso(nombre, argumentos):
    """Executa una peca (organizar/borradores) i no deixa que res tombi el bucle."""
    hora = datetime.now().strftime("%H:%M:%S")
    print(f"\n[{hora}] --- {nombre} ---", flush=True)
    try:
        r = subprocess.run([sys.executable, os.path.join(AQUI, argumentos[0])] + argumentos[1:],
                           cwd=os.path.dirname(os.path.dirname(AQUI)), timeout=3600)
        if r.returncode != 0:
            print(f"[{hora}] (avis: {nombre} ha acabat amb codi {r.returncode}; el cicle continua)")
    except subprocess.TimeoutExpired:
        print(f"[{hora}] (avis: {nombre} ha superat 1h i s'ha tallat; el cicle continua)")
    except Exception as e:
        print(f"[{hora}] (avis: {nombre} ha fallat: {str(e)[:80]}; el cicle continua)")


def ciclo(cfg):
    paso("ORGANITZAR la porta d'entrada",
         ["organizar.py", "--outlook", cfg["buzon"], "--carpeta", cfg["entrada"]])
    paso("ESBORRANYS del calaix 1 FACIL",
         ["borradores.py", "--outlook", cfg["buzon"], "--carpeta", "1 FACIL", "--real"])
    paso("ESBORRANYS del calaix 2 DIFICIL",
         ["borradores.py", "--outlook", cfg["buzon"], "--carpeta", "2 DIFICIL", "--real"])
    paso("ESBORRANYS del calaix 3 DESISTIMIENTO",
         ["borradores.py", "--outlook", cfg["buzon"], "--carpeta", "3 DESISTIMIENTO", "--real"])


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--una-passada", action="store_true", help="un sol cicle i sortir")
    args = ap.parse_args()
    cfg = config()
    print("=" * 64)
    print(" AGENT EN MARXA — cicle organitzar + esborranys")
    print(f" Buzon: {cfg['buzon']} | Entrada: {cfg['entrada']} | Cada {cfg['intervalo']} min")
    print(f" Horari: {cfg['hora_inicio']:02d}:00-{cfg['hora_fin']:02d}:00"
          + ("" if cfg["fines_de_semana"] else " (nomes laborables)"))
    print(" Aturar: Ctrl+C. Cap peca d'aquest sistema pot enviar correus.")
    print("=" * 64)
    n = 0
    try:
        while True:
            if not en_horario(cfg) and not args.una_passada:
                dormir_fins_horari(cfg)
            n += 1
            print(f"\n{'#' * 64}\n# CICLE {n} — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n{'#' * 64}")
            ciclo(cfg)
            if args.una_passada:
                print("\nUna passada feta. Sortint."); break
            print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Cicle {n} acabat."
                  f" Dormint {cfg['intervalo']} min (Ctrl+C per aturar)...")
            time.sleep(cfg["intervalo"] * 60)
    except KeyboardInterrupt:
        print("\nAturat per l'usuari. Fins a la propera passada.")
