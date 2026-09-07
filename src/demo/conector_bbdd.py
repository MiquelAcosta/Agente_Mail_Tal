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
apto, coche_cerrado, factura, contrato_compra, modelo_576, solicitud_mod_576,
denegacion_modelo_576, herencia, poderes_pleitos, ha_vendido_coche,
envio_fyg, estado_fyg, envio_minsait, estado_minsait"""


def _componer_estado(d):
    """Compon l'estat de cara al client a partir del cicle de vida real.
    Retorna (frase_apta_per_a_client, detall_intern)."""
    interno = " | ".join(f"{k}={d[k]}" for k in
                         ("estado_vehiculo_ayp", "apto", "envio_fyg", "estado_fyg",
                          "estado_minsait", "procedimiento", "coche_cerrado") if d.get(k))
    ev = d["estado_vehiculo_ayp"].upper()
    if d["coche_cerrado"]:
        return "expediente cerrado", interno
    if d["procedimiento"]:
        return "en fase de procedimiento judicial", interno
    if d["envio_fyg"] or d["estado_fyg"]:
        return "documentación revisada y expediente remitido para su tramitación en la demanda", interno
    if ev.startswith("PEDIDO_"):
        que = ev.replace("PEDIDO_", "").replace("_", " ").strip().lower()
        return f"pendiente de documentación: {que}", interno
    if ev.startswith("OUT_") or ev in ("NO AFECTADO", "CAMION", "ADQUIRIDO FUERA DE ESPAÑA"):
        return "vehículo no elegible según la revisión (derivar a persona para explicación)", interno
    if d["apto"] == "No":
        return "revisado como no apto (derivar a persona para explicación)", interno
    if ev == "OK" or d["apto"] == "Si":
        return "documentación en revisión, expediente en preparación", interno
    return "sin información de estado", interno


def _traducir(fila):
    """De la fila crua de wp_base_xam al format de fitxa del pipeline."""
    d = {k: _texto(v) for k, v in fila.items()}
    titular = " ".join(x for x in (d["nombre"], d["apellido1"], d["apellido2"]) if x)
    if not titular and d["empresa"]:
        titular = d["empresa"] + " (empresa)"
    estado, detalle = _componer_estado(d)
    return {
        "matricula": d["matricula"] or "(sin matrícula)",
        "titular": titular or "(sin nombre)",
        "tipo_titular": d["tipo_titular"],
        "estado": estado,
        "detalle_interno": detalle,
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


def ficha_por_email(email):
    """Compatibilitat amb el pipeline actual: una sola fitxa (la primera) o None."""
    fichas = fichas_por_email(email, max_fichas=1)
    return fichas[0] if fichas else None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Us:  python src\\demo\\conector_bbdd.py correo@a.buscar")
        sys.exit(1)
    try:
        fichas = fichas_por_email(sys.argv[1])
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
