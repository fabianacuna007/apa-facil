# APA Fácil

Aplica automáticamente el formato APA 7 (versión de estudiante) a trabajos académicos en Word (.docx). Nunca cambia el texto: solo el formato, y siempre en una copia nueva.

El plan de desarrollo está en [ROADMAP.md](ROADMAP.md).

## Instalación (Windows)

Hace falta Python 3.11 o más reciente. Desde la carpeta del proyecto:

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

El primer comando crea un entorno aislado para el proyecto, el segundo lo activa y el tercero instala lo que necesita. A partir de ahí, basta con activar el entorno (segundo comando) cada vez que se abra una terminal nueva.

## Uso

Con el entorno activado, desde la carpeta del proyecto:

```
python -m apa_facil pruebas\01_ensayo_basico.docx
```

La copia se guarda en `salida/` con el nombre `nombre_APA.docx`. El original nunca se modifica. Se pueden procesar varios documentos a la vez (`python -m apa_facil pruebas\*.docx`).

Para elegir la fuente de todo el trabajo:

```
python -m apa_facil pruebas\01_ensayo_basico.docx --fuente "Arial"
```

Las fuentes recomendadas por APA, con su tamaño, son Aptos 12, Calibri 11, Arial 11, Lucida Sans Unicode 10, Times New Roman 12, Georgia 11 y Computer Modern 10. Si no se indica `--fuente`, se mantiene la del documento si es una de ellas, o se usa la más parecida.

Para quitar la numeración de los títulos («1.1. Objetivos» → «Objetivos») o pasar a minúsculas los títulos escritos en mayúsculas («RESUMEN» → «Resumen»), que APA no usa:

```
python -m apa_facil pruebas\05_tfg_completo.docx --quitar-numeracion --quitar-mayusculas
```

Con `--quitar-introduccion` se quita el rótulo «Introducción» del primer apartado, que APA no usa (el título del trabajo hace sus veces), si el apartado no tiene subapartados.

Las tres opciones son independientes. Sin ellas solo se avisa, porque cambian el texto y muchas universidades exigen numerar los títulos. Al quitar las mayúsculas se respetan los nombres propios y las siglas que el trabajo escribe así en el texto.

### Datos de la portada

Los datos de la portada de estudiante (título, autores, departamento, universidad, código y nombre de la asignatura, profesor y fecha de entrega) se dan en un archivo JSON con `--datos`. Hay un ejemplo en `docs\datos_ejemplo.json`:

```
python -m apa_facil pruebas\01_ensayo_basico.docx --datos docs\datos_ejemplo.json
```

También se pueden dar con opciones sueltas, que mandan sobre el archivo dato a dato: `--titulo`, `--autor` (una vez por autor), `--departamento`, `--universidad`, `--codigo-asignatura`, `--asignatura`, `--profesor` (una vez por profesor) y `--fecha`. La fecha se escribe como `16/10/2026`, `2026-10-16` o `16 de octubre de 2026`.

Los datos se usan tal cual se escriben: solo se quitan los espacios de los extremos. Lo que falte no se rellena, ni siquiera la fecha: se avisa en el informe. Si el archivo está mal escrito (una coma de más, una clave desconocida, una fecha que no existe), no se procesa nada y se explica qué corregir. El informe JSON solo dice qué datos se han dado, no sus valores.

Con esos datos se genera la **portada APA de estudiante**, como la describe la web oficial de APA:
- el título, en negrita y centrado, tres líneas por debajo del margen superior;
- una línea en blanco;
- debajo, centrados: los autores («Ana Pérez y Luis Gómez»), la afiliación («Departamento…, Universidad…»), la asignatura («HIS 201: Historia…»), el profesor y la fecha («16 de octubre de 2026»).

Todo va a doble espacio, y el texto empieza en una página nueva con el mismo título. El título lleva mayúscula inicial en las palabras principales, como pide APA: los artículos y las conjunciones y preposiciones de tres letras o menos van en minúscula. Con `--titulo-tal-cual` se deja como lo escribes. «que», «lo» y «vía» pueden ir de las dos formas según la frase: se dejan como estén y se avisa. De momento la portada solo se añade a documentos que no tienen ninguna; si ya hay una, se deja como está y se avisa.

### Qué hace (formato básico, títulos y limpieza)

- **Márgenes** de 2,54 cm en todas las secciones; las páginas horizontales siguen horizontales.
- **Una sola fuente** recomendada por APA en todo el trabajo, con su tamaño (salvo el código de programación, los símbolos y las ecuaciones).
- **Interlineado doble** sin espacio extra entre párrafos (en las tablas se respeta el sencillo, 1,5 o doble; las notas al pie no se tocan).
- **Texto alineado a la izquierda**, sin justificar ni guiones automáticos.
- **Sangría de 1,27 cm** en la primera línea de los párrafos de texto (no en títulos, listas, referencias, citas en bloque…).
- **Título del trabajo** repetido al principio del texto, centrado y en negrita: se copia tal cual de la portada.
- **Títulos** con los estilos de Word «Título 1» a «Título 5» y el formato APA de cada nivel (así aparecen en el panel de navegación), sin formato puesto a mano que lo contradiga; los de nivel 4 y 5, en la misma línea que su texto y terminados en punto (con el separador de estilo de Word: en LibreOffice se ven en su propia línea). Avisa de los apartados con un solo subapartado y de los títulos numerados o en mayúsculas.
- **Número de página** arriba a la derecha en todas las páginas, también en la portada, y nada más en el encabezado.
- **Limpieza**: quita los párrafos vacíos usados para dejar espacio, los tabuladores o espacios usados como sangría y el color, el resaltado y el subrayado del texto (los enlaces se conservan). No toca la portada ni las tablas.
- **Avisos** de lo que conviene revisar a mano sin que la herramienta lo cambie: saltos de línea manuales, viñetas escritas a mano, dobles espacios entre palabras y texto que estaba en otra fuente (suele venir de pegarlo de una web).

Al terminar, muestra para cada documento qué ha cambiado y qué conviene revisar a mano, y guarda ese mismo informe en JSON junto a la copia (`salida\nombre_APA.json`), con la regla de la que viene cada mensaje. Antes de entregar la copia comprueba que el texto, las tablas, las imágenes, las notas y las secciones son exactamente los del original; si no, no la entrega.

## Ver el formato de un documento

El inspector muestra el formato que tiene ahora un Word (márgenes, fuentes, interlineado, sangrías, alineación, estilos, encabezados y pies), sin cambiar nada:

```
python herramientas\inspeccionar.py pruebas\04_formato_caotico.docx
```

## Ver cómo se entiende un documento (diagnóstico)

Antes de aplicar títulos, sangrías o referencias, la herramienta decide qué es cada párrafo: portada, título (con su nivel), texto, lista, leyenda, cita en bloque, referencia… El modo diagnóstico muestra esa decisión en una tabla, sin generar ninguna copia:

```
python -m apa_facil pruebas\04_formato_caotico.docx --diagnostico
```

Cada fila lleva el número de párrafo, el tipo, la confianza (alta, media o baja) y el principio del texto; una línea separa cada parte del trabajo (portada, resumen, índice, cuerpo, referencias, anexos). En los párrafos dudosos (confianza media o baja) se explica debajo por qué se han clasificado así. Al final, cuántos párrafos hay de cada tipo y con cada confianza.

## Comprobar si un documento cumple APA

El verificador comprueba las reglas de formato básico (márgenes, fuente, interlineado, alineación, sangría y número de página) y los títulos, y lista cada incumplimiento con su sección, su número de párrafo y un extracto del texto. Lo que la herramienta deja así a propósito (la portada de la universidad, el código de programación, los títulos numerados…) aparece aparte, como excepción admitida. Termina con código 0 si todo cumple y 1 si no:

```
python herramientas\verificar.py salida\04_formato_caotico_APA.docx
```

## Revisar en PDF

Convierte documentos a PDF para revisarlos sin abrir Word. Usa Word si está instalado (el resultado es el que verá el estudiante) y, si no, LibreOffice. Los PDF se guardan siempre en `salida\`, aunque el documento venga de `pruebas\`, y los documentos no se modifican:

```
python herramientas\a_pdf.py salida\*.docx
```

## Pruebas automáticas

```
pytest
```
