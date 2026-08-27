# 06 · Costes estimados

Estimaciones de agosto 2026 con el alcance v1 (sin lectura de adjuntos entrantes).
Volumen de referencia: 100-200 mails/día. **Verificar precios vigentes de las APIs
antes de presupuestar formalmente.**

## Operación mensual

| Concepto | Estimación |
|---|---|
| Tokens (pipeline Haiku/Sonnet, con prompt caching) | 20-40 € |
| Infraestructura | ~0 € (Windows Server propio ya existente) |
| Licencias Microsoft extra | 0 € (Graph incluido en M365) |
| **Total operativo** | **~20-40 €/mes** |

Orden de magnitud por mail procesado: 2-4 céntimos. El coste de tokens es marginal:
no condiciona ninguna decisión de diseño.

## Construcción (inversión única)

- Esfuerzo estimado con este alcance: **~1 a 1,5 meses de un desarrollador**
  (Fase 1: 1-2 semanas · Fase 2: 2-3 semanas · integración BBDD y ajustes: resto).
- Coste según quién lo haga (órdenes de magnitud, cotizar 2-3 opciones):
  interno = nómina del período · freelance senior (España) ≈ 8.000-15.000 € ·
  consultora ≈ 15.000-30.000 €.

## Retorno

- Situación actual: correspondencia ≈ carga de trabajo de ~3 personas.
- Escenario realista tras rodaje: 30-50% de mails resueltos en automático + el resto
  pre-procesado (resumen + borrador) → ~40-60% del tiempo total liberado.
- Con eso, la construcción se amortiza en pocos meses; el operativo es despreciable.
- Beneficios no monetarios: respuesta en minutos (también fuera de horario),
  consistencia, nada se pierde en la bandeja, trazabilidad completa.
