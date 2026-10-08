"""informe_autoria.py — QUI HA ESCRIT CADA RESPOSTA ENVIADA

Llegeix la carpeta d'Elements enviats i classifica cada resposta com a
PERSONA o IA a partir del text del missatge. NOMES LLEGEIX: no mou, no
escriu, no envia, no toca cap mail.

Heuristiques (senyals de PERSONA):
  - pics de llista (•, -, 1., etc.)
  - imatges o adjunts de debo (les de la signatura no compten)
  - missatge molt curt (excepte si sembla desistiment)
  - majuscules cridades, exclamacions multiples, emoticones
  - signatura escrita a ma (nom de persona al tancament)

Senyal d'IA: text llarg, sense pics, i que s'assembla molt a altres
respostes del mateix periode (la IA repeteix l'arbre de respostes).

IMPORTANT: l'eina contrasta la seva propia heuristica amb log_borradores.sqlite
allà on hi ha registre. Aixo dona un % d'encert REAL de les heuristiques.
Si l'encert es baix, el percentatge global no s'ha de donar per bo.

Us:
  python informe_autoria.py --outlook "info" --desde 2026-10-01 --csv autoria.csv
  python informe_autoria.py --outlook "info" --desde 2026-10-01 --dias 30
"""

import argparse
import csv
import os
import re
import sqlite3
import sys
from collections import Counter
from datetime import datetime, timedelta

try:
    import win32com.client
except ImportError:
    print("Falta pywin32. Instal·la'l amb: pip install pywin32")
    sys.exit(1)

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_REGISTRE = os.path.join(AQUI, "log_borradores.sqlite")

# ---------------------------------------------------------------- constants

# Llindar per sota del qual un missatge es considera "molt curt".
CURT_CARACTERS = 180

# Similitud a partir de la qual dos missatges es consideren "el mateix text".
LLINDAR_BESSONS = 0.72

# Talls tipics del fil citat d'Outlook (castella, catala, angles).
TALLS_CITA = [
    r"\n\s*-{2,}\s*Mensaje original\s*-{2,}",
    r"\n\s*_{5,}",
    r"\n\s*De:\s",
    r"\n\s*From:\s",
    r"\n\s*Enviado el:\s",
    r"\n\s*El .{0,40}escribi[oó]:",
    r"\n\s*On .{0,40}wrote:",
]
RE_TALL = re.compile("|".join(TALLS_CITA), re.IGNORECASE)

# Pics de llista: rodons, quadrats, guions i numeracions a principi de linia.
RE_PIC = re.compile(r"^\s*(?:[•●▪◦‣·○\-\*–—]|\d{1,2}[\.\)])\s+", re.MULTILINE)

# Senyals d'escriptura humana informal.
RE_EXCLAMA = re.compile(r"[!¡]{2,}|\?{2,}")
RE_EMOJI = re.compile(r"[\U0001F300-\U0001FAFF\u2600-\u27BF]")
RE_CRIDAT = re.compile(r"\b[A-ZÁÉÍÓÚÑ]{4,}\b")

# Desistiment: els textos de cancel·lacio son curts per disseny.
RE_DESISTIMENT = re.compile(
    r"\b(desist|cancel|baja|anular|rescind|no deseo continuar|renunci)", re.IGNORECASE)

# Adjunts que NO compten com a imatge de debo (venen de la signatura o del fil).
RE_ADJ_IGNORA = re.compile(r"(^~WRD|^image\d{3}|logo|firma|signature|banner)", re.IGNORECASE)

# Paraules buides que no aporten res a la comparacio de textos.
BUIDES = {"de", "la", "el", "en", "y", "a", "los", "las", "un", "una", "que",
          "del", "al", "se", "su", "sus", "por", "con", "para", "es", "lo",
          "le", "les", "no", "ha", "han", "como", "mas", "o", "si", "ya"}


# ------------------------------------------------------------------ utils

def cos_net(text):
    """Treu el fil citat i la signatura: deixa nomes el missatge nou."""
    if not text:
        return ""
    tall = RE_TALL.search(text)
    if tall:
        text = text[:tall.start()]
    # Treu l'avis de Microsoft si hi es.
    text = re.sub(r"No suele recibir correo.*?$", "", text,
                  flags=re.IGNORECASE | re.MULTILINE)
    return text.strip()


def paraules(text):
    """Bossa de paraules normalitzada per comparar textos."""
    t = re.sub(r"[^\wàáèéíòóúïüçñ\s]", " ", (text or "").lower())
    return [p for p in t.split() if len(p) > 2 and p not in BUIDES]


def similitud(a, b):
    """Jaccard sobre bosses de paraules. 0 = res a veure, 1 = identics."""
    sa, sb = set(a), set(b)
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def adjunts_reals(item):
    """Compta adjunts que no siguin logos/firmes/imatges heretades del fil."""
    n = 0
    try:
        for i in range(1, item.Attachments.Count + 1):
            nom = str(item.Attachments.Item(i).FileName or "")
            if not RE_ADJ_IGNORA.search(nom):
                n += 1
    except Exception:
        pass
    return n


# ------------------------------------------------------- heuristica central

def analitzar(cos, n_adjunts):
    """Torna (senyals_persona, detall). Cada senyal es una raó legible."""
    senyals = []
    net = cos_net(cos)
    nc = len(net)

    if RE_PIC.search(net):
        senyals.append("pics de llista")
    if n_adjunts > 0:
        senyals.append(f"{n_adjunts} adjunt(s) real(s)")
    if nc < CURT_CARACTERS and not RE_DESISTIMENT.search(net):
        senyals.append(f"molt curt ({nc} car.)")
    if RE_EXCLAMA.search(net):
        senyals.append("exclamacions/interrogants repetits")
    if RE_EMOJI.search(net):
        senyals.append("emoticones")
    if len(RE_CRIDAT.findall(net)) >= 3:
        senyals.append("paraules en majuscules")

    return senyals, nc


def classificar(senyals, bessons):
    """PERSONA si hi ha qualsevol senyal huma. IA si no en te cap.
    'bessons' (quants altres enviats s'hi assemblen) nomes reforça la IA."""
    if senyals:
        return "PERSONA"
    if bessons >= 2:
        return "IA"
    return "IA (feble)"   # cap senyal huma pero tampoc text repetit


# ------------------------------------------------------------------ Outlook

NOMS_ENVIATS = ("elementos enviados", "sent items", "enviados",
                "elements enviats", "elementos enviados ")


def _buscar_enviats(arrel):
    """Busca la carpeta d'enviats dins d'una arrel. None si no hi es."""
    for j in range(1, arrel.Folders.Count + 1):
        if str(arrel.Folders.Item(j).Name).lower().strip() in NOMS_ENVIATS:
            return arrel.Folders.Item(j)
    return None


def escollir_arrel(arrels, nom):
    """Mateixa prioritat que escoger_buzon() del projecte:
    nom EXACTE -> COMENCA per -> CONTE.

    Sense aixo, 'info@recuperatudinero.com' tambe casa amb
    'Online Archive - info@recuperatudinero.com', i com que l'arxiu surt
    abans a la llista guanyava ell: 52.330 enviats vells i cap de recent.
    """
    n = (nom or "").strip().lower()
    noms = [str(getattr(a, "Name", "") or "") for a in arrels]
    for prova in (lambda x: x.lower().strip() == n,
                  lambda x: x.lower().strip().startswith(n)):
        tri = [a for a, x in zip(arrels, noms) if prova(x)]
        if tri:
            return tri[0], len(tri)
    conte = [a for a, x in zip(arrels, noms) if n in x.lower()]
    if not conte:
        return None, 0
    return conte[0], len(conte)


def carpeta_enviats(buzon):
    """Troba Elements enviats de la bustia demanada.

    OJO amb les busties compartides: segons com estigui el perfil, les
    respostes enviades DES de la bustia compartida es desen a la carpeta
    d'enviats PERSONAL de qui les envia, no a la de la compartida.
    """
    ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    if not buzon:
        return ol.GetDefaultFolder(5)   # olFolderSentMail

    arrels = [ol.Folders.Item(i) for i in range(1, ol.Folders.Count + 1)]
    arrel, quants = escollir_arrel(arrels, buzon)
    if arrel is None:
        print("  Magatzems disponibles al perfil:")
        for a in arrels:
            print(f"     - {a.Name}")
        raise SystemExit(f"No trobo cap bustia que contingui '{buzon}'")

    print(f"    Magatzem triat: '{arrel.Name}'"
          + (f"   (ATENCIO: {quants} candidats, pot ser ambigu)" if quants > 1 else ""))

    c = _buscar_enviats(arrel)
    if c:
        return c
    print(f"  Carpetes dins de '{arrel.Name}':")
    for j in range(1, arrel.Folders.Count + 1):
        print(f"     - {arrel.Folders.Item(j).Name}")
    raise SystemExit("No hi ha cap carpeta d'enviats reconeguda.")


def totes_les_carpetes_enviats():
    """Totes les carpetes d'enviats del perfil, per si la compartida es buida."""
    ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    fora = []
    for i in range(1, ol.Folders.Count + 1):
        arrel = ol.Folders.Item(i)
        try:
            c = _buscar_enviats(arrel)
            if c:
                fora.append((str(arrel.Name), c))
        except Exception:
            continue
    return fora


def llegir_enviats(carpeta, desde, maxim):
    """Llegeix els enviats a partir d'una data. Nomes lectura."""
    items = carpeta.Items
    total_carpeta = 0
    try:
        total_carpeta = items.Count
    except Exception:
        pass
    print(f"    Carpeta '{carpeta.Name}': {total_carpeta} elements en total.")
    if total_carpeta == 0:
        return []

    try:
        items.Sort("[SentOn]", True)
    except Exception:
        pass

    # La mitjanit en format 12h es '12:00 AM', NO '00:00 AM' (aixo tornava 0).
    filtrats = None
    for patro in ("[SentOn] >= '%m/%d/%Y 12:00 AM'",
                  "[SentOn] >= '%d/%m/%Y 12:00 AM'"):
        try:
            prova = items.Restrict(desde.strftime(patro))
            if prova.Count > 0:
                filtrats = prova
                break
        except Exception:
            continue
    if filtrats is None:
        print("    (el filtre de data no ha donat res: es filtra aqui, mes lent)")
        filtrats = items

    fora, errors, fora_rang = [], 0, 0
    for it in filtrats:
        if len(fora) >= maxim:
            break
        try:
            enviat = getattr(it, "SentOn", None)
            if enviat is not None:
                enviat = enviat.replace(tzinfo=None)
                if enviat < desde:
                    fora_rang += 1
                    continue
            fora.append({
                "data": enviat,
                "para": str(getattr(it, "To", "") or ""),
                "asunto": str(getattr(it, "Subject", "") or ""),
                "cos": str(getattr(it, "Body", "") or ""),
                "adj": adjunts_reals(it),
            })
        except Exception:
            errors += 1
            continue

    if fora_rang:
        print(f"    {fora_rang} elements descartats per ser anteriors a la data.")
    if errors:
        print(f"    ATENCIO: {errors} elements no s'han pogut llegir"
              " (cites, convocatories o elements no sincronitzats).")
    return fora


# ------------------------------------------------------------- el registre

def carregar_registre():
    """Textos d'esborrany que l'agent va deixar. Per contrastar l'heuristica."""
    if not os.path.exists(RUTA_REGISTRE):
        return []
    try:
        con = sqlite3.connect(RUTA_REGISTRE)
        files = con.execute(
            "SELECT ts, remitente, asunto, respuesta FROM borradores"
            " WHERE respuesta IS NOT NULL AND respuesta <> ''").fetchall()
        con.close()
        return [{"ts": r[0], "rem": (r[1] or "").lower(),
                 "asunto": r[2] or "", "txt": r[3] or ""} for r in files]
    except Exception as e:
        print(f"    (no s'ha pogut llegir el registre: {str(e)[:60]})")
        return []


def te_esborrany(env, registre):
    """L'agent havia redactat alguna cosa per aquest fil? Torna la similitud."""
    para = (env["para"] or "").lower()
    assum = re.sub(r"^\s*(re|rv|fwd?)\s*:\s*", "", env["asunto"], flags=re.I).lower()
    millor = 0.0
    for r in registre:
        if r["rem"] and r["rem"] not in para:
            continue
        ra = re.sub(r"^\s*(re|rv|fwd?)\s*:\s*", "", r["asunto"], flags=re.I).lower()
        if assum and ra and assum[:40] != ra[:40]:
            continue
        s = similitud(paraules(cos_net(env["cos"])), paraules(r["txt"]))
        millor = max(millor, s)
    return millor


# ----------------------------------------------------------------- informe

def main(buzon, desde, maxim, ruta_csv):
    print(f"\n  Llegint enviats de: {buzon or '(bustia per defecte)'}")
    print(f"  Des de: {desde:%Y-%m-%d}   (maxim {maxim} elements)\n")

    carpeta = carpeta_enviats(buzon)
    enviats = llegir_enviats(carpeta, desde, maxim)
    print(f"  {len(enviats)} respostes enviades al periode.\n")

    if not enviats:
        print("  Cap resultat aqui. Mirant la resta de carpetes d'enviats"
              " del perfil...\n")
        for nom, c in totes_les_carpetes_enviats():
            if c.EntryID == carpeta.EntryID:
                continue
            if "archive" in nom.lower() or "archivo" in nom.lower():
                continue   # els arxius no porten correu recent
            altres = llegir_enviats(c, desde, maxim)
            if altres:
                print(f"\n  >>> N'hi ha {len(altres)} a '{nom}'."
                      f"\n      Les respostes de la bustia compartida es desen aqui."
                      f"\n      Torna a llancar amb:  --outlook \"{nom}\"\n")
                return
        print("  Tampoc n'hi ha a cap altra carpeta d'enviats.\n"
              "  Possibles causes:\n"
              "   - la finestra de sincronitzacio no cobreix aquests dies\n"
              "   - les respostes s'envien des de l'OWA i no es baixen a l'.ost\n"
              "   - realment no s'ha contestat res en aquest periode\n")
        return

    # Bosses de paraules un sol cop.
    for e in enviats:
        e["net"] = cos_net(e["cos"])
        e["bossa"] = paraules(e["net"])

    # Bessons: quants altres enviats diuen practicament el mateix.
    for i, e in enumerate(enviats):
        n = 0
        for j, altre in enumerate(enviats):
            if i != j and similitud(e["bossa"], altre["bossa"]) >= LLINDAR_BESSONS:
                n += 1
        e["bessons"] = n

    registre = carregar_registre()
    if registre:
        print(f"  Registre d'esborranys: {len(registre)} textos per contrastar.\n")
    else:
        print("  Sense registre d'esborranys: no es podra mesurar l'encert.\n")

    for e in enviats:
        e["senyals"], e["nchar"] = analitzar(e["cos"], e["adj"])
        e["veredicte"] = classificar(e["senyals"], e["bessons"])
        e["sim_registre"] = te_esborrany(e, registre) if registre else None

    # -------- recompte
    recompte = Counter(e["veredicte"] for e in enviats)
    total = len(enviats)
    print("  " + "=" * 56)
    print("  AUTORIA DE LES RESPOSTES ENVIADES")
    print("  " + "=" * 56)
    for etiqueta in ("PERSONA", "IA", "IA (feble)"):
        n = recompte.get(etiqueta, 0)
        print(f"    {etiqueta:12}  {n:5}   {100*n/total:5.1f} %")
    ia = recompte.get("IA", 0) + recompte.get("IA (feble)", 0)
    print(f"    {'-'*40}")
    print(f"    {'IA (total)':12}  {ia:5}   {100*ia/total:5.1f} %")

    # -------- per que s'ha dit PERSONA
    raons = Counter()
    for e in enviats:
        for s in e["senyals"]:
            raons[re.sub(r"\(.*\)", "", s).strip()] += 1
    if raons:
        print("\n  SENYALS DE PERSONA (un mail en pot tenir varis):")
        for r, n in raons.most_common():
            print(f"    {n:5}  {r}")

    # -------- contrast amb el registre: aixo es el que diu si et pots fiar
    if registre:
        amb = [e for e in enviats if (e["sim_registre"] or 0) >= 0.60]
        print(f"\n  CONTRAST AMB EL REGISTRE D'ESBORRANYS")
        print(f"    {len(amb)} enviats coincideixen amb un esborrany de l'agent.")
        if amb:
            encerts = sum(1 for e in amb if e["veredicte"].startswith("IA"))
            print(f"    L'heuristica n'encerta {encerts}/{len(amb)}"
                  f"  ({100*encerts/len(amb):.0f} %)")
            if encerts / len(amb) < 0.80:
                print("    ATENCIO: encert baix. El % global NO es fiable.")
            quasi = sum(1 for e in amb if (e["sim_registre"] or 0) >= 0.95)
            print(f"    Enviats TAL QUAL (sense retocar): {quasi}/{len(amb)}")
            print(f"    Retocats abans d'enviar:          {len(amb)-quasi}/{len(amb)}")

    # -------- csv
    if ruta_csv:
        with open(ruta_csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["data", "para", "asunto", "veredicte", "senyals",
                        "n_caracters", "bessons", "similitud_registre", "text"])
            for e in sorted(enviats, key=lambda x: x["data"] or datetime.min):
                w.writerow([
                    e["data"].strftime("%Y-%m-%d %H:%M") if e["data"] else "",
                    e["para"][:80], e["asunto"][:90], e["veredicte"],
                    " | ".join(e["senyals"]), e["nchar"], e["bessons"],
                    f"{e['sim_registre']:.2f}" if e["sim_registre"] is not None else "",
                    e["net"][:1500].replace("\n", " ")])
        print(f"\n  CSV: {ruta_csv}")

    print("\n  Recorda: l'heuristica mira el TEXT. Revisa una mostra del CSV"
          "\n  abans de donar el percentatge per bo.\n")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--outlook", metavar="BUSTIA", default=None)
    ap.add_argument("--desde", metavar="AAAA-MM-DD", default=None,
                    help="data d'arrencada de l'agent (obligatori si no hi ha --dias)")
    ap.add_argument("--dias", type=int, default=0, help="alternativa a --desde")
    ap.add_argument("--max", type=int, default=2000)
    ap.add_argument("--csv", metavar="FITXER", default="")
    a = ap.parse_args()

    if a.desde:
        d = datetime.strptime(a.desde, "%Y-%m-%d")
    elif a.dias:
        d = datetime.now() - timedelta(days=a.dias)
    else:
        raise SystemExit("Cal --desde AAAA-MM-DD o --dias N "
                         "(sense data comptaries com a IA tot el correu "
                         "anterior a l'agent)")
    main(a.outlook, d, a.max, a.csv)