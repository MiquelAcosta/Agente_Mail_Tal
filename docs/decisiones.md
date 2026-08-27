# Registro de decisiones

Formato: fecha · decisión · motivo. Añadir arriba las nuevas.

- **2026-08-27 · Repo y estructura creados.** Documentación inicial volcada desde la
  sesión de diseño del proyecto.
- **2026-08 · Sucesiones/herencias: el agente jamás responde.** En la práctica casi
  siempre entran con adjunto (filtro 1); el residual lo detecta el clasificador y solo
  cabe respuesta de petición de documentación. Hilo marcado queda marcado.
- **2026-08 · Filtros duros v1: adjunto real → humano · remitente sin historial → humano.**
- **2026-08 · v1 sin lectura de adjuntos entrantes** (pasa a v2). Adjuntos de salida SÍ,
  con lista cerrada de ~5 documentos.
- **2026-08 · Integración BBDD por matrícula, solo lectura**, con cruce
  remitente↔matrícula como control anti-suplantación.
- **2026-08 · Un único sistema con pasos internos** (clasificador/redactor/verificador),
  no múltiples agentes. Polling en vez de webhooks. Corre en el Windows Server propio.
- **2026-08 · Ficha persistente de cliente pospuesta**; en v1, búsqueda de historial
  al vuelo en el buzón. Log de decisiones sí, desde el día 1, innegociable.
