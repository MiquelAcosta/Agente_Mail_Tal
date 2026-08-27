# 02 · Arquitectura

## Visión general

Un **único sistema** (no varios agentes) corriendo en el Windows Server de la empresa,
con pasos internos especializados. Flujo por cada mail:

```
Mail llega al buzón (Exchange Online)
        │  (polling vía Microsoft Graph cada 1-5 min)
        ▼
[TRIAJE DETERMINISTA]  — sin IA, solo metadata y patrones
  ├─ ¿Adjunto real? ──────────────► carpeta HUMANO (+ resumen)
  ├─ ¿Remitente sin historial? ──► carpeta HUMANO (+ resumen)
  └─ pasa ▼
[CLASIFICADOR]  (modelo económico, p.ej. Haiku)
  ├─ ¿Sucesión / complejo / repregunta? ──► HUMANO (+ resumen)
  └─ categoría simple ▼
[CONSULTA BBDD]  — extrae matrícula, consulta solo-lectura
  ├─ ¿Sin matrícula o no cuadra con el remitente? ──► HUMANO / respuesta genérica
  └─ datos verificados ▼
[REDACTOR]  (modelo capaz, p.ej. Sonnet) — redacta + elige doc. de salida (lista cerrada)
        ▼
[VERIFICADOR]  (modelo económico) — checklist: ¿afirma solo datos de BBDD/mails reales?
  ├─ no pasa ──► HUMANO
  └─ pasa ▼
ENVÍO (o BORRADOR según configuración de la categoría) + LOG
```

## Componentes

| Pieza | Tecnología | Notas |
|---|---|---|
| Conexión al buzón | Microsoft Graph API | App registrada en Entra ID; permisos Mail.ReadWrite + Mail.Send **acotados al buzón** (application access policy). Polling, no webhooks (el servidor no necesita ser accesible desde internet). |
| Runtime | Python (o Node) como servicio de Windows / tarea programada | Sobrevive reinicios. |
| IA | API de Anthropic o Azure OpenAI | Residencia UE + DPA + sin entrenamiento. Prompt caching para la parte fija. |
| BBDD interna | Usuario técnico **solo lectura**, campos mínimos | Cada consulta se loguea. |
| Log de decisiones | SQLite o SQL Server | Qué mail, qué decisión, por qué, qué respondió. Imprescindible. |
| Documentos de salida | Carpeta SharePoint controlada (~5 ficheros) | El negocio los actualiza sin tocar código. |
| Historial del cliente | Búsqueda en el buzón al vuelo (Graph) | Sin ficha persistente en v1. |

## Decisiones de diseño ya tomadas

- **Polling y no webhooks**: más simple, no expone el servidor.
- **Lista cerrada de adjuntos de salida**: el agente elige de un catálogo; imposible
  adjuntar nada fuera de él.
- **El modelo nunca inventa datos**: todo dato concreto sale de la BBDD o de un mail
  real recuperado; lo verifica el paso final.
- **Filtro inline-attachments**: los logos de firma NO cuentan como adjunto
  (propiedad `isInline` + tipo + tamaño).
- **Hilo marcado**: un hilo derivado por sucesión (u otra marca) queda derivado;
  las repreguntas a respuestas automáticas van a humano.
