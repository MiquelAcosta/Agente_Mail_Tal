# 05 · Legal — RGPD y AI Act

**Estado: pendiente de validar con el DPO.** Este doc resume lo acordado como punto
de partida; la palabra final es del DPO / legal.

## Naturaleza del tratamiento

- El proveedor de IA actúa como **encargado del tratamiento** (como Microsoft con M365).
- Base jurídica: la misma con la que hoy se responde el correo (ejecución de
  contrato / interés legítimo). Cambia el *cómo*, no el *por qué*.
- Datos: comerciales/patrimoniales (categoría general). Las sucesiones pueden rozar
  categoría especial (p.ej. certificados de defunción) → **quedan fuera del circuito
  automático por diseño**, lo que simplifica el análisis.

## Requisitos acordados

- [ ] **DPIA** antes de la puesta en marcha. Involucrar al DPO desde la Fase 0.
- [ ] **DPA** con el proveedor de IA + **residencia de datos UE** + sin uso para
      entrenamiento (Anthropic API o Azure OpenAI lo ofrecen; si la casa ya está en
      Azure, Azure OpenAI = menor fricción).
- [ ] **Minimización**: cada paso recibe solo los campos necesarios. Valorar
      seudonimización (sustituir nombre/mail/matrícula por marcadores antes de llamar
      al modelo y reinsertar al final).
- [ ] **Art. 22 RGPD**: el agente responde y gestiona, pero no toma *decisiones* con
      efectos jurídicos/significativos — cualquier decisión va a humano (ya cubierto
      por los filtros del doc 03).
- [ ] **Transparencia (AI Act + buenas prácticas)**: firma de «respuesta generada
      automáticamente» + mención en la política de privacidad.
- [ ] **Registro de actividades de tratamiento** actualizado.
- [ ] Retención definida para el log de decisiones: ______ meses.
- [ ] Acceso a la BBDD interna: usuario técnico **solo lectura**, campos mínimos,
      consultas logueadas.
- [ ] Seguridad anti-suplantación: información de terceros/patrimonial nunca en
      automático; cruce remitente↔matrícula obligatorio.

## Argumentos clave para la conversación con el DPO

1. Perímetro automático mínimo y defendible: cliente conocido, solo texto,
   identidad cruzada con BBDD, temas simples.
2. Sucesiones y adjuntos fuera del pipeline de IA → los documentos delicados ni
   siquiera pasan por el modelo.
3. Log completo = auditabilidad total.
4. Modo borrador disponible: supervisión humana del 100% mientras se quiera.
