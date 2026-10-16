# APA Fácil

Herramienta web que aplica automáticamente las normas APA 7 (versión de estudiante) a trabajos académicos en Word. Público: estudiantes hispanohablantes. Toda la interfaz y los mensajes, en español.

## Plan de desarrollo

El plan completo, paso a paso, está en `ROADMAP.md`. Léelo al empezar cada sesión.

- Trabaja solo en el paso que se te indique. No adelantes pasos de fases posteriores.
- Al terminar un paso: márcalo con `[x]` en ROADMAP.md, añade una línea al "Registro de avances" y haz un commit.
- La especificación APA 7 completa está al final de ROADMAP.md.

## Tecnología

- Núcleo: Python 3 con `python-docx`, y `lxml` para lo que python-docx no cubre. Pruebas con `pytest`.
- Más adelante: API con FastAPI y web con Next.js.

## Reglas de trabajo (obligatorias)

- Nunca modificar el texto del usuario, solo el formato. Cualquier excepción tiene que estar prevista en ROADMAP.md y quedar registrada en el informe.
- Nunca sobrescribir el archivo original: generar siempre una copia nueva en `salida/`.
- Conservar tablas, imágenes, notas al pie, títulos y secciones.
- No inventar fuentes ni datos de referencias. Si falta información, señalarla, no rellenarla.
- Antes de escribir código, explicar el plan paso a paso. El propietario del proyecto quiere entender cómo funciona cada parte.
- Probar cada cambio con los archivos de `pruebas/` e informar de qué funcionó y qué no en cada uno.
- Hacer un commit de git tras cada avance que funcione.

## Estructura de carpetas

- `apa_facil/`: el núcleo (motor de formato).
- `herramientas/`: inspector y verificador.
- `tests/`: pruebas automáticas.
- `pruebas/`: documentos Word para probar. No modificar.
- `salida/`: documentos generados.
