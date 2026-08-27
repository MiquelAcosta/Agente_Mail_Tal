# 03 · Filtros y reglas de negocio

**Este documento es la fuente de verdad del comportamiento del agente.**
Cualquier cambio aquí debe reflejarse en los prompts/código en el mismo PR.

## El filtro, en orden de evaluación

| # | Regla | Evaluación | Destino |
|---|---|---|---|
| 1 | Trae **adjunto real** (excluye logos de firma) | Metadata Graph, determinista | Humano |
| 2 | **Remitente sin historial** en el buzón | Búsqueda Graph, determinista | Humano |
| 3 | **Señales de sucesión/herencia** en mail o hilo (fallecimiento, herederos, testamento, viudedad, defunción, «era cliente»...) — ante la duda, SÍ | Clasificador, sesgo deliberado a derivar | Humano + hilo marcado permanente |
| 4 | **Complejo / ambiguo / repregunta** a una respuesta automática previa | Clasificador | Humano |
| 5 | **Sin matrícula identificable**, o la matrícula **no cuadra** con el remitente según BBDD/historial | Regex + consulta BBDD | Humano (o respuesta genérica pidiendo el dato) |
| 6 | Todo lo demás | — | Circuito automático |

Consecuencia: toda respuesta automática es a **cliente conocido + solo texto +
matrícula coherente + tema simple y no sucesorio**.

## Reglas duras del circuito automático

1. Datos concretos (fechas, estados, importes) **solo** si provienen de la BBDD o de un
   mail real recuperado. Nunca deducidos.
2. Información patrimonial o de terceros **jamás** se envía en automático.
3. Si no se recuperó historial, la respuesta no puede referirse a interacciones pasadas
   (lo chequea el verificador).
4. Adjuntos de salida: solo de la lista cerrada (catálogo en doc 07 cuando se defina).
5. Toda respuesta automática se identifica como tal (firma) — pendiente de texto final.
6. Todo queda en el log: mail, decisión, motivo, respuesta, consultas a BBDD.

## Pendiente de completar con la Fase 0

- [ ] Lista definitiva de categorías y su volumen (%).
- [ ] Qué categorías arrancan en modo borrador y cuáles podrán pasar a automático.
- [ ] Catálogo de los ~5 documentos de salida y la regla de cuándo va cada uno.
- [ ] Plantillas de respuesta tipo por categoría.
- [ ] Formato exacto de la matrícula (regex) y campo(s) de la BBDD para el cruce
      remitente↔matrícula.
