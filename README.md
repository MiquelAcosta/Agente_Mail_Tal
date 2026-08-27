# Agente de correo — triaje y respuesta automática del buzón compartido

Sistema que clasifica los mails entrantes del buzón compartido, responde automáticamente
los casos simples de clientes conocidos (adjuntando documentos oficiales cuando toca) y
deriva el resto al equipo con un resumen del caso ya preparado.

**Estado:** Fase 0 — clasificación de la muestra de mails y definición de reglas.

## Mapa del repositorio

| Carpeta | Contenido |
|---|---|
| `docs/` | Toda la documentación del proyecto (alcance, arquitectura, reglas, legal, costes, roadmap) |
| `docs/plantillas/` | Plantilla Excel de la Fase 0 |
| `src/` | Código del agente (vacío hasta Fase 1; estructura ya definida) |
| `config/` | Plantillas de configuración. **Nunca subir credenciales reales** |
| `scripts/` | Utilidades sueltas (exportación de muestras, etc.) |

## Por dónde empezar

1. Leer `docs/01-vision-y-alcance.md` (10 min) — qué hace y qué NO hace el sistema.
2. Leer `docs/03-filtros-y-reglas-negocio.md` — el corazón del proyecto.
3. Si vas a clasificar mails de la muestra: `docs/04-fase0-clasificacion.md`.
4. Si vas a programar: `docs/02-arquitectura.md` y el README de `src/`.

## Reglas del repositorio

- **Repo PRIVADO siempre.** Contiene lógica de negocio interna.
- **Prohibido subir datos de clientes**: ni mails reales, ni exportaciones, ni la plantilla
  de Fase 0 rellenada. El `.gitignore` bloquea lo obvio, pero la responsabilidad es de quien
  hace commit.
- **Prohibido subir credenciales** (API keys, secretos de la app de Graph, conexiones a BBDD).
  Solo se versionan los `*.example`.
- Forma de trabajo del equipo: ver `CONTRIBUTING.md`.
