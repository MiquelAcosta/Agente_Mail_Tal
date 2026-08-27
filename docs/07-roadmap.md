# 07 · Roadmap

## Fases de la v1

| Fase | Contenido | Duración | Criterio de salida |
|---|---|---|---|
| **0. Reglas** | Clasificar muestra ~300 mails, definir categorías, docs de salida, plantillas. DPO avisado, DPIA en marcha. | 1-2 semanas | Doc 03 completado con datos reales |
| **1. Clasificador en sombra** | El agente solo clasifica y mueve a carpetas. No responde nada. | 1-2 sem. desarrollo + rodaje | Precisión de triaje medida y aceptable por categoría |
| **2. Borradores** | Redacta + elige adjunto de salida, deja **borrador**; humano revisa y envía. | 2-3 sem. desarrollo + rodaje | % de borradores enviados sin retoques, por categoría |
| **3. Automático gradual** | Categorías con ≥95% de acierto pasan a envío automático, empezando por las inocuas. El resto sigue en borrador. | continuo | Revisión periódica del log |

Quedarse permanentemente en modo borrador es una opción válida (riesgo ≈ 0,
conserva la mayor parte del ahorro).

## Versiones futuras (fuera de v1)

- **v2 — Lectura de adjuntos entrantes**: OCR/visión, validación de completitud,
  extracción de datos. Reabre el mail «difícil» a la automatización parcial.
  Requiere ampliar la DPIA.
- **v3 — Remitentes nuevos**: exige verificación de identidad (se caen las tres
  anclas: historial, matrícula, BBDD). Diseñar con legal.
- **v3+ — Escritura en BBDD / cierre de gestiones**: el agente deja de solo responder
  y pasa a gestionar. Solo cuando la confianza y los controles lo permitan.

## Checklist de arranque (qué falta conseguir)

- [ ] Muestra de mails exportada y clasificada (Fase 0)
- [ ] App registrada en Entra ID + permisos acotados al buzón (IT)
- [ ] Regla de salida del firewall hacia Graph y la API de IA (IT)
- [ ] Cuenta empresarial del proveedor de IA (residencia UE + DPA) (compras/legal)
- [ ] Usuario solo-lectura de la BBDD interna con campos mínimos (admin BBDD)
- [ ] Carpeta SharePoint con los ~5 documentos de salida + dueño asignado (negocio)
- [ ] DPO involucrado / DPIA iniciada (legal)
- [ ] Desarrollador asignado o contratado
