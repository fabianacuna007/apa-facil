# APA Fácil

Herramienta web que aplica automáticamente las normas APA 7 (versión de estudiante) a trabajos académicos en Word. Público: estudiantes hispanohablantes. Toda la interfaz y los mensajes, en español.

## Objetivo actual (fase 1)

Un script en Python que reciba un .docx y devuelva una copia con formato APA 7. Sin web y sin IA todavía.

Flujo: subir .docx → aplicar formato APA 7 → guardar `nombre_APA.docx`

## Tecnología

- Python 3 con `python-docx` para leer y escribir Word.
- Más adelante: frontend web (probablemente Next.js) y un backend que use este script.

## Reglas de trabajo (obligatorias)

- Nunca modificar el texto del usuario: solo el formato.
- Nunca sobrescribir el archivo original: generar siempre una copia nueva.
- Conservar tablas, imágenes, notas al pie y encabezados existentes.
- No inventar fuentes ni datos de referencias. Si falta información, señalarla, no rellenarla.
- Antes de escribir código, explicar el plan paso a paso. El propietario del proyecto quiere entender cómo funciona cada parte.
- Probar cada cambio con los archivos de `pruebas/` e informar de qué funcionó y qué no en cada uno.
- Hacer un commit de git tras cada avance que funcione.

## Formato APA 7 de la fase 1

- Márgenes: 2,54 cm en los cuatro lados.
- Fuente: Times New Roman 12 por defecto. También se admiten Calibri 11, Arial 11, Georgia 11 y Lucida Sans Unicode 10.
- Interlineado doble en todo el documento.
- Sangría de primera línea de 1,27 cm en los párrafos de texto.
- Alineación a la izquierda (sin justificar).
- Número de página arriba a la derecha, en todas las páginas.

## Fases siguientes (no empezar aún)

1. Portada de estudiante (título, autor, universidad, asignatura, profesor, fecha).
2. Los cinco niveles de encabezado APA.
3. Lista de referencias: título "Referencias", orden alfabético y sangría francesa.
4. Revisión de citas en el texto frente a la lista de referencias.
5. Asistencia con IA para explicar errores, sin modificar el contenido.
6. Interfaz web para subir, configurar y descargar.

## Estructura de carpetas

- `pruebas/`: documentos Word reales para probar. No modificar.
- `salida/`: documentos generados por el script.
