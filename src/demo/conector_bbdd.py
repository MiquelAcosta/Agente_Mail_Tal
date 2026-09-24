# -*- coding: utf-8 -*-
"""Connector NOMES LECTURA a la BBDD real (taula wp_base_xam).

Busca la fitxa d'un client per email (a les 5 columnes d'email de la taula)
i la tradueix al format que el pipeline ja coneix (com l'escenario.json).

Configuracio: src\\demo\\credenciales.json  (MAI al GitHub) amb:
{
  "bbdd": {"host": "...", "puerto": 3306, "esquema": "bitnami_wordpress",
            "usuario": "...", "password": "..."}
}

Prova rapida des de l'arrel del repo:
  python src\\demo\\conector_bbdd.py correo@de.prueba.com

Requisit:  pip install mysql-connector-python
"""
import json, os, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RUTA_CRED = os.path.join(AQUI, "credenciales.json")

COLUMNAS_EMAIL = ["email_propietario", "email_registro", "email_alta",
                  "email_coche", "email_firma"]

CAMPOS = """matricula, marca, modelo, nombre, apellido1, apellido2, empresa,
tipo_titular, estado_vehiculo_ayp, estado_propietario_ayp, procedimiento,
apto, coche_cerrado, estado_perito, factura, contrato_compra, modelo_576, solicitud_mod_576,
denegacion_modelo_576, herencia, poderes_pleitos, ha_vendido_coche,
envio_fyg, estado_fyg, envio_minsait, estado_minsait,
total_perito, Diligencias_Preliminares, Solicitando_576, Denegado_576,
fecha_aviso_ya_esta"""


def _importe_perito(valor):
    """Import de la indemnitzacio en format espanyol (1.886,81 €), o None.

    Retorna None si el camp es buit o no es un numero: en aquest cas la resposta
    NO es pot automatitzar i el correu ha d'anar a una persona. Mai s'envia una
    plantilla amb el buit sense omplir."""
    if valor is None:
        return None
    t = str(valor).strip().replace("€", "").replace(" ", "")
    if not t:
        return None
    t = t.replace(".", "").replace(",", ".") if ("," in t and "." in t) else t.replace(",", ".")
    try:
        n = float(t)
    except Exception:
        return None
    if n <= 0:
        return None
    ent, dec = f"{n:,.2f}".split(".")
    return ent.replace(",", ".") + "," + dec + " €"


def _si(v):
    """True si el camp marca afirmatiu (Si/Sí/S/1/X/True...)."""
    t = str(v or "").strip().lower()
    return t in ("si", "sí", "s", "1", "x", "true", "yes", "ok")


def _componer_estado(d):
    """Estat de l'expedient segons l'arbre de respostes de l'equip.

    Retorna (codi, frase_per_al_client, detall_intern). El CODI es el que fa
    servir l'arbre per triar la plantilla; la frase es per al redactor.

      E1  Completo SIN informe pericial
      E2  Completo CON informe pericial
      E3  Incompleto, NO diligencies preliminars
      E3.1 Pendent de documentacio (se sap quina falta)
      E4  Diligencies preliminars, model 576 SOL.LICITAT
      E5  Diligencies preliminars, model 576 NO sol.licitat
      E6  Vehicle no elegible / no apte
      E7  (el codi no el pot donar: client no identificat)
    """
    interno = " | ".join(f"{k}={d[k]}" for k in
                         ("estado_vehiculo_ayp", "apto", "coche_cerrado", "estado_perito",
                          "Diligencias_Preliminares", "Solicitando_576", "Denegado_576",
                          "envio_fyg", "estado_fyg", "estado_minsait", "procedimiento")
                         if d.get(k))
    ev = str(d.get("estado_vehiculo_ayp") or "").upper()

    # --- No elegible: mana sobre tota la resta -----------------------------
    if ev.startswith("OUT_") or ev in ("NO AFECTADO", "CAMION", "ADQUIRIDO FUERA DE ESPAÑA") \
       or str(d.get("apto") or "").strip().lower() == "no":
        return "E6", "vehículo no elegible según la revisión", interno

    # --- Expedient tancat (revisio completa) -------------------------------
    if _si(d.get("coche_cerrado")):
        if str(d.get("estado_perito") or "").strip():
            return "E2", ("completo, con informe pericial recibido, "
                          "pendiente de interposición de la demanda"), interno
        return "E1", "completo y revisado, en espera del informe pericial", interno

    # --- Diligencies preliminars -------------------------------------------
    if _si(d.get("Diligencias_Preliminares")):
        if _si(d.get("Solicitando_576")) or str(d.get("solicitud_mod_576") or "").strip():
            return "E4", ("en diligencias preliminares, con el modelo 576 solicitado"), interno
        return "E5", "en diligencias preliminares, sin el modelo 576 solicitado", interno

    # --- Falta documentacio concreta ---------------------------------------
    if ev.startswith("PEDIDO_"):
        que = ev.replace("PEDIDO_", "").replace("_", " ").strip().lower()
        return "E3.1", f"pendiente de documentación: {que}", interno
    faltan = [n for n, k in (("factura", "factura"),
                             ("contrato de compra", "contrato_compra"),
                             ("modelo 576", "modelo_576"),
                             ("poderes", "poderes_pleitos"))
              if not str(d.get(k) or "").strip()]
    if faltan and not (d.get("envio_fyg") or d.get("estado_fyg")):
        return "E3.1", "pendiente de documentación: " + ", ".join(faltan), interno

    # --- Incomplet, sense diligencies --------------------------------------
    return "E3", "documentación en revisión, expediente en preparación", interno


def _traducir(fila):
    """De la fila crua de wp_base_xam al format de fitxa del pipeline."""
    d = {k: _texto(v) for k, v in fila.items()}
    titular = " ".join(x for x in (d["nombre"], d["apellido1"], d["apellido2"]) if x)
    if not titular and d["empresa"]:
        titular = d["empresa"] + " (empresa)"
    codi, estado, detalle = _componer_estado(d)
    return {
        "matricula": d["matricula"] or "(sin matrícula)",
        "titular": titular or "(sin nombre)",
        "tipo_titular": d["tipo_titular"],
        "estado_codigo": codi,
        "estado": estado,
        "detalle_interno": detalle,
        "importe_perito": _importe_perito(d.get("total_perito")),
        "diligencias_preliminares": _texto(d.get("Diligencias_Preliminares")),
        "solicitando_576": _texto(d.get("Solicitando_576")),
        "denegado_576": _texto(d.get("Denegado_576")),
        "avisado_ya_esta": bool(_texto(d.get("fecha_aviso_ya_esta"))),
        "documentacion": {
            "factura": _hay_doc(d["factura"]),
            "contrato_compra": _hay_doc(d["contrato_compra"]),
            "modelo_576": _hay_doc(d["modelo_576"]),
            "poderes": _hay_doc(d["poderes_pleitos"]),
        },
        "vehiculo": " ".join(x for x in (d["marca"], d["modelo"]) if x),
        "sospecha_sucesion_bbdd": _texto(fila.get("herencia")) != "",
        "vendido": d["ha_vendido_coche"],
    }


def _conexion():
    import mysql.connector
    cred = json.load(open(RUTA_CRED, encoding="utf-8-sig"))["bbdd"]
    return mysql.connector.connect(
        host=cred["host"], port=int(cred.get("puerto", 3306)),
        database=cred["esquema"], user=cred["usuario"], password=cred["password"],
        charset="utf8mb4", connection_timeout=10)


def _texto(v):
    if v is None:
        return ""
    return str(v).strip()


def _hay_doc(v):
    """A la BBDD, els camps de document son NULL (no hi ha) o IDs/arrays (hi ha)."""
    return "Sí" if _texto(v) not in ("", "0") else "No"


def fichas_por_email(email, max_fichas=5):
    """Retorna la llista de fitxes (expedients) associades a un email. [] si cap."""
    email = (email or "").strip().lower()
    if not email:
        return []
    cond = " OR ".join(f"LOWER(TRIM({c})) = %s" for c in COLUMNAS_EMAIL)
    sql = (f"SELECT {CAMPOS} FROM wp_base_xam WHERE {cond} LIMIT {int(max_fichas)}")
    cn = _conexion()
    try:
        cur = cn.cursor(dictionary=True)
        cur.execute(sql, [email] * len(COLUMNAS_EMAIL))
        return [_traducir(f) for f in cur.fetchall()]
    finally:
        cn.close()


def fichas_por_matricula(matricula, max_fichas=3):
    """Cerca per matricula (normalitzant espais i majuscules). Us: LOCALITZAR la
    fitxa quan l'email falla — MAI com a autenticacio per a envio automatic."""
    mat = (matricula or "").replace(" ", "").replace("-", "").upper()
    if not mat:
        return []
    sql = (f"SELECT {CAMPOS} FROM wp_base_xam "
           f"WHERE UPPER(REPLACE(REPLACE(TRIM(matricula),' ',''),'-','')) = %s "
           f"LIMIT {int(max_fichas)}")
    cn = _conexion()
    try:
        cur = cn.cursor(dictionary=True)
        cur.execute(sql, [mat])
        return [_traducir(f) for f in cur.fetchall()]
    finally:
        cn.close()


_COLS_TEL = None


def _columnas_telefono(cn):
    global _COLS_TEL
    if _COLS_TEL is None:
        cur = cn.cursor()
        cur.execute("SELECT COLUMN_NAME FROM information_schema.COLUMNS "
                    "WHERE TABLE_NAME='wp_base_xam' AND COLUMN_NAME LIKE '%telefono%'")
        _COLS_TEL = [r[0] for r in cur.fetchall()]
    return _COLS_TEL


def fichas_por_telefono(telefono, max_fichas=3):
    """Cerca per telefon (nomes digits). LOCALITZADOR: mai autenticacio d'enviament."""
    tel = "".join(ch for ch in str(telefono) if ch.isdigit())
    if len(tel) < 9:
        return []
    tel = tel[-9:]
    cn = _conexion()
    try:
        cols = _columnas_telefono(cn)
        if not cols:
            return []
        cond = " OR ".join(
            f"REPLACE(REPLACE(REPLACE(TRIM({c}),' ',''),'-',''),'.','') LIKE %s" for c in cols)
        cur = cn.cursor(dictionary=True)
        cur.execute(f"SELECT {CAMPOS} FROM wp_base_xam WHERE {cond} LIMIT {int(max_fichas)}",
                    ["%" + tel] * len(cols))
        return [_traducir(f) for f in cur.fetchall()]
    finally:
        cn.close()


def ficha_por_telefono(telefono):
    fichas = fichas_por_telefono(telefono, max_fichas=1)
    return fichas[0] if fichas else None


def ficha_por_matricula(matricula):
    fichas = fichas_por_matricula(matricula, max_fichas=1)
    return fichas[0] if fichas else None


def ficha_por_email(email):
    """Compatibilitat amb el pipeline actual: una sola fitxa (la primera) o None."""
    fichas = fichas_por_email(email, max_fichas=1)
    return fichas[0] if fichas else None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Us:  python src\\demo\\conector_bbdd.py correo@a.buscar")
        sys.exit(1)
    try:
        arg = sys.argv[1]
        fichas = fichas_por_email(arg) if "@" in arg else fichas_por_matricula(arg)
    except FileNotFoundError:
        print(f"Falta {RUTA_CRED} amb les credencials."); sys.exit(1)
    except Exception as e:
        print(f"Error de connexio o consulta: {e}"); sys.exit(1)
    if not fichas:
        print("Cap fitxa amb aquest email (el triatge ho enviaria a HUMANO per sin_ficha_bbdd).")
    else:
        print(f"{len(fichas)} fitxa/es trobades:")
        for i, f in enumerate(fichas, 1):
            print(f"\n--- Fitxa {i} ---")
            for k, v in f.items():
                print(f"  {k}: {v}")