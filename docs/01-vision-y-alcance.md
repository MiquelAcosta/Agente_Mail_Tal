# 01 · Visión y alcance

## Qué es

Un agente de IA conectado al buzón compartido de Outlook que:

1. **Clasifica** todo mail entrante.
2. **Responde automáticamente** los casos simples de clientes conocidos, adjuntando
   cuando corresponda uno de los ~5 documentos oficiales.
3. **Deriva al equipo humano** todo lo demás, acompañado de un resumen del caso
   (qué pide, por qué se derivó, antecedentes del hilo).

El objetivo no es eliminar el trabajo humano sino quitarle la parte mecánica: el equipo
pasa de procesar correo a gestionar casos y controlar calidad.

## Alcance de la versión 1 (lo que SÍ hace)

- Procesa **solo mails de texto** (sin adjuntos entrantes).
- Responde solo a **remitentes con historial** en el buzón.
- Usa la **BBDD interna por matrícula** (solo lectura) para responder con datos verificados.
- Adjunta documentos de salida de una **lista cerrada de ~5** (carpeta controlada).
- Deja **log de cada decisión** (auditable).
- Modo de envío configurable por categoría: automático o **borrador que aprueba un humano**.

## Fuera de alcance de la v1 (roadmap futuro)

- **Leer/escanear adjuntos entrantes** → v2.
- **Atender a remitentes nuevos** (sin historial) → v3, requiere verificación de identidad.
- **Escribir en la BBDD** / cerrar gestiones → v3+.
- Sucesiones y herencias: **el agente jamás las responde**; en la práctica casi siempre
  llevan adjunto (filtro 1) y el caso residual se cubre con la detección del clasificador
  (ver doc 03).

## Criterio rector

> Ante la duda, derivar a humano. El coste de derivar de más es bajo;
> el de responder mal en automático puede ser alto.
