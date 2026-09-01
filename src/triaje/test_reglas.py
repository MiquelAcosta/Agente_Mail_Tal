# -*- coding: utf-8 -*-
from reglas import correo_sistema, adjunto_real, sospecha_sucesion, triaje, REGEX_MATRICULA

HIST = {"cliente@mail.com"}
BBDD = {"cliente@mail.com": {"matricula": "2717GJD", "titular": "Cliente Prueba",
                             "estado": "En fase de demanda"}}
BASE = {"remitente": "cliente@mail.com", "asunto": "Estado", "cuerpo": "¿Cómo va lo mío?", "adjuntos": []}

def test_filtro0_sistema():
    assert triaje({**BASE, "remitente": "mailer-daemon@x.com"}, HIST, BBDD)[0] == "SISTEMA"

def test_logo_no_es_adjunto():
    assert not adjunto_real([{"nombre": "image001.png", "is_inline": False, "bytes": 8000}])

def test_orden_filtros():
    assert triaje(BASE, HIST, BBDD)[0] == "CIRCUITO"
    assert triaje({**BASE, "adjuntos": [{"nombre": "factura.pdf", "is_inline": False, "bytes": 99999}]},
                  HIST, BBDD)[0] == "HUMANO"
    assert triaje({**BASE, "remitente": "nuevo@x.com"}, HIST, BBDD)[0] == "HUMANO"
    assert triaje(BASE, HIST, {})[0:2] == ("HUMANO", "sin_ficha_bbdd")
    assert triaje({**BASE, "cuerpo": "mi padre, titular, ha fallecido"}, HIST, BBDD)[0] == "HUMANO_SUCESION"

def test_ficha_devuelta():
    destino, motivo, ficha = triaje(BASE, HIST, BBDD)
    assert ficha and ficha["matricula"] == "2717GJD"

def test_matricula_es_utilidad():
    assert REGEX_MATRICULA.search("mi coche es el 2717 GJD")

if __name__ == "__main__":
    for n, f in list(globals().items()):
        if n.startswith("test_"): f(); print("OK ", n)
    print("Todos los tests pasan.")
