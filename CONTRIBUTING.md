# Cómo trabajamos en este repo (equipo de 3)

Reglas mínimas para no pisarnos, sin burocracia de más.

## Ramas

- `main`: siempre estable. Nadie hace commit directo a `main` (proteger la rama en
  GitHub: Settings → Branches → branch protection → require pull request).
- Una rama por tarea, con nombre descriptivo: `docs/reglas-sucesiones`,
  `feat/triaje-adjuntos`, `fix/regex-matricula`.
- Al terminar: Pull Request → la revisa otra persona del equipo → merge → borrar la rama.
  Con 3 personas, una aprobación basta.

## Issues = lista de tareas

- Cada tarea pendiente es un Issue de GitHub. Antes de empezar algo, **asígnate el Issue**:
  esa es la señal de "esto lo estoy haciendo yo" y lo que evita solaparse.
- Usar etiquetas: `fase-0`, `fase-1`, `docs`, `codigo`, `legal`, `bloqueado`.
- Lo que no está en un Issue no existe. Las decisiones tomadas fuera (reuniones, chats)
  se apuntan en el Issue correspondiente o en `docs/decisiones.md`.

## Documentación

- Los `docs/` son la fuente de verdad del proyecto. Si una regla de negocio cambia,
  se cambia el doc en el mismo PR que el código que la implementa.
- Cambios de documentación también van por PR, aunque sean pequeños: así todos ven
  qué cambió.

## Datos y secretos — tolerancia cero

- Ningún mail real, exportación, ni Excel rellenado entra al repo. Trabajadlos en
  SharePoint/OneDrive corporativo.
- Ninguna credencial en el código. Configuración sensible → variables de entorno o
  fichero local ignorado por git (partir de `config/settings.example.env`).
- Si algo sensible entra por error: avisad inmediatamente; hay que reescribir el
  historial, no basta con borrarlo en el siguiente commit.
