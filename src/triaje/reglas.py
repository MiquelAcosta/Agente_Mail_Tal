# -*- coding: utf-8 -*-
"""Filtros deterministas del triaje — v2 (alineado con doc 03 v2)."""
import re

REMITENTE_SISTEMA = re.compile(r"(mailer-daemon|postmaster|no[-_.]?reply|notificacion|bounce)", re.I)
NOMBRES_FALSO_ADJUNTO = re.compile(r"^(image\d{3}|logo|icon|firma|banner)", re.I)
UMBRAL_BYTES_IMG_FIRMA = 30 * 1024
EXT_IMAGEN = (".png", ".gif", ".jpg", ".jpeg", ".bmp")

# Utilidad secundaria (NO filtro): verificar una matrícula si el cliente la escribe.
REGEX_MATRICULA = re.compile(r"\b\d{4}\s?[B-DF-HJ-NP-TV-Z]{3}\b")  # formato español moderno


def correo_sistema(remitente):
    """Filtro 0: rebotes y notificaciones automáticas. ~16% del tráfico real."""
    return bool(REMITENTE_SISTEMA.search(remitente or ""))


def adjunto_real(adjuntos):
    for a in adjuntos or []:
        if a.get("is_inline"):
            continue
        nombre = a.get("nombre", "")
        es_img = nombre.lower().endswith(EXT_IMAGEN)
        if es_img and (NOMBRES_FALSO_ADJUNTO.match(nombre)
                       or a.get("bytes", 0) < UMBRAL_BYTES_IMG_FIRMA):
            continue
        return True
    return False


PALABRAS_SUCESION = ("fallecimiento", "fallecido", "fallecida", "defunción",
                     "defuncion", "herencia", "herederos", "heredero", "testamento",
                     "viudedad", "viudo", "viuda", "era cliente", "ha muerto")


def sospecha_sucesion(texto):
    t = (texto or "").lower()
    return any(p in t for p in PALABRAS_SUCESION)


def triaje(mail, historial_remitentes, bbdd_por_email):
    """Filtros 0-4 en orden (doc 03 v2). Devuelve (destino, motivo, ficha|None).
    bbdd_por_email: dict email→ficha (mock de la BBDD; en real, consulta solo-lectura).
    """
    rem = (mail.get("remitente") or "").lower()
    if correo_sistema(rem):
        return "SISTEMA", "correo_sistema", None
    if adjunto_real(mail.get("adjuntos")):
        return "HUMANO", "adjunto_real", None
    if rem not in historial_remitentes:
        return "HUMANO", "remitente_sin_historial", None
    ficha = bbdd_por_email.get(rem)
    if ficha is None:
        return "HUMANO", "sin_ficha_bbdd", None
    texto = f"{mail.get('asunto','')} {mail.get('cuerpo','')}"
    if sospecha_sucesion(texto):
        return "HUMANO_SUCESION", "sospecha_sucesion", ficha
    return "CIRCUITO", "pasa_filtros", ficha
