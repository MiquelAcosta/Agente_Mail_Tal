# src/ — código del agente (Fase 1 en adelante)

Estructura prevista, alineada con `docs/02-arquitectura.md`:

| Carpeta | Responsabilidad |
|---|---|
| `triaje/` | Filtros deterministas: adjunto real (excluye inline), historial del remitente, extracción de matrícula (regex) |
| `clasificador/` | Llamada al modelo económico: categoría + señales (sucesión, complejidad, repregunta) |
| `conectores/` | Cliente de Microsoft Graph (polling, lectura, carpetas, borradores, envío) · cliente BBDD solo-lectura · cliente API de IA |
| `redactor/` | Redacción con datos verificados + elección de documento de salida (lista cerrada) |
| `verificador/` | Checklist final antes de enviar/guardar borrador |

Transversal: log de decisiones (cada módulo escribe en él).

Convenciones (a fijar por quien desarrolle): lenguaje propuesto Python 3.11+,
configuración por variables de entorno (ver `config/settings.example.env`),
sin credenciales en código, tests mínimos por filtro determinista.
