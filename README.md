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

Los títulos llevan mayúscula inicial en las palabras principales, como pide APA: «Marco teórico» → «Marco Teórico», «RESUMEN» → «Resumen». Se respetan los nombres propios y las siglas que el trabajo escribe así en el texto, y la numeración del principio. Con `--titulos-tal-cual`, los títulos de los apartados se quedan como están; con `--titulo-tal-cual`, el título del trabajo.

Para quitar la numeración de los títulos («1.1. Objetivos» → «Objetivos»), que APA no usa:

```
python -m apa_facil pruebas\05_tfg_completo.docx --quitar-numeracion
```

Con `--quitar-introduccion` se quita el rótulo «Introducción» del primer apartado, que APA no usa (el título del trabajo hace sus veces), si el apartado no tiene subapartados.

Sin `--quitar-numeracion` ni `--quitar-introduccion` solo se avisa, porque cambian el texto y muchas universidades exigen numerar los títulos.

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

Todo va a doble espacio, y el texto empieza en una página nueva con el mismo título. El título lleva mayúscula inicial en las palabras principales, como pide APA: los artículos y las conjunciones y preposiciones de tres letras o menos van en minúscula. Con `--titulo-tal-cual` se deja como lo escribes. «que», «lo» y «vía» pueden ir de las dos formas según la frase: se dejan como estén y se avisa.

Si el documento ya tiene una portada (la de la universidad), con tus datos **se sustituye** por la portada APA, logo incluido, y el informe dice qué líneas se han quitado. Con `--mantener-portada` se deja como está, con su diseño, y solo se hace que el texto empiece en una página nueva; muchas universidades exigen su propia portada, y APA lo admite. Sin datos de portada, también se mantiene. Tampoco se sustituye, y se avisa, si la portada lleva algo que podría ser del trabajo: una tabla, una nota, un campo o un párrafo largo.

### Resumen y palabras clave

Si el trabajo tiene resumen, se le da el formato de la guía oficial de APA: el rótulo «Resumen» centrado y en negrita al principio de su página, el texto sin sangría, y la línea de palabras clave con sangría, con el rótulo «*Palabras clave:*» en cursiva y las palabras en letra normal y sin punto al final. Lo que sigue al resumen empieza en una página nueva. Si el resumen pasa de 250 palabras, tiene varios párrafos o las palabras clave no son de tres a cinco o empiezan con mayúscula sin ser nombres propios, solo se avisa.

### Índice

APA no usa el índice en los trabajos de estudiante, pero muchas universidades lo exigen, así que **se mantiene** y se avisa. Para quitarlo:

```
python -m apa_facil pruebas\05_tfg_completo.docx --quitar-indice
```

Se quita el índice automático de Word o el escrito a mano (el rótulo «Índice» y sus líneas con número de página), sin dejar una página en blanco ni perder secciones. Si algo de lo que parece índice no tiene forma de línea de índice, no se quita nada y se avisa.

### Tablas y figuras

Como las describe la web oficial de APA (las tablas y las figuras se presentan igual):
- encima, «**Tabla 1**» o «**Figura 1**» en negrita y, en la línea siguiente, el título en *cursiva* y con mayúscula inicial en las palabras principales, los dos a la izquierda y sin sangría. Si la leyenda está debajo, sube encima; si el número y el título van en la misma línea («Tabla 1: Título»), se separan;
- APA llama «Tabla» a todas las tablas y «Figura» a todo lo demás: «Cuadro 1» pasa a «Tabla 1», y «Gráfico 1», «Imagen 2» o «Esquema 3» pasan a «Figura 1», «Figura 2» y «Figura 3»;
- debajo, la nota, que empieza por «*Nota.*» en cursiva. A una nota que empieza por «Fuente:» se le pone «Nota.» delante;
- si ninguna tabla (o figura) tiene número, se añade «Tabla 1», «Tabla 2»… a las que tienen título (si unas lo tienen y otras no, solo se avisa);
- en las tablas, la fila de encabezados se repite en cada página si la tabla ocupa más de una, y solo hay líneas horizontales (arriba, debajo de los encabezados y abajo), sin bordes verticales ni colores de fondo. Con `--tablas-tal-cual` se dejan los bordes y los colores como están.

Una figura es un párrafo con solo una imagen; la imagen no se toca. Si falta la línea en blanco entre la tabla o figura y el texto, si no tiene título, si la numeración no va seguida o si el título lleva «(elaboración propia)» (que APA pone en una nota), solo se avisa. Si una tabla es demasiado ancha, el aviso propone ponerla en una página horizontal.

### Citas en bloque

Una cita de 40 palabras o más que va en su propio párrafo con sangría a la izquierda se pone como pide la web oficial de APA: todo el bloque sangrado 1,27 cm desde el margen izquierdo, sin sangría a la derecha ni en la primera línea (si tiene varios párrafos, la primera línea del segundo y los siguientes lleva 1,27 cm más), a doble espacio y sin la cursiva que vaya en todo el texto (la de una parte, como el título de un libro, se respeta).

Lo que cambiaría el texto solo se avisa: las comillas (APA no las pone en una cita en bloque), un punto después del paréntesis de la fuente, un párrafo con sangría de menos de 40 palabras (si es una cita, va dentro del texto y entre comillas) y una cita de 40 palabras o más escrita entre comillas dentro de un párrafo.

### Lista de referencias

Como la describe la web oficial de APA: empieza en una página nueva, con el rótulo «Referencias» («Referencia» si solo hay una) centrado y en negrita (con el estilo «Título 1» de Word), y cada referencia con sangría francesa de 1,27 cm. Los enlaces (DOI y URL) se dejan como estén.

APA no admite «Bibliografía», «Webgrafía» ni «Lista de referencias», así que el rótulo pasa a «Referencias». Si tu universidad pide otro, déjalo como está con:

```
python -m apa_facil pruebas_tfg_completo.docx --rotulo-tal-cual
```

Las referencias se ordenan alfabéticamente, como pide APA: por los apellidos del primer autor (las obras sin autor, por el título), luego por el año (sin fecha, primero) y luego por el título. Las tildes no cuentan y la «ñ» va después de la «n». Una línea que solo es una dirección web va al final, con un aviso. Para dejar el orden que tenían, usa `--orden-tal-cual`.

Si hay varias listas (por ejemplo, una bibliografía y una webgrafía), no se cambia ningún rótulo ni se ordena nada, y se avisa: únelas en una sola.

### Apéndices

Cada apéndice empieza en una página nueva, con el rótulo y, en la línea siguiente, su título, los dos centrados y en negrita, como pide la web oficial de APA. Si van en la misma línea («Anexo I. Guion de las entrevistas»), se separan; el título lleva las mayúsculas de APA.

APA los llama «Apéndice» («Apéndice A», «Apéndice B»… si hay varios), pero muchas universidades piden «Anexo», así que por defecto se deja y se avisa. Para cambiarlo:

```
python -m apa_facil pruebas_tfg_completo.docx --rotulo-apendice
```

Si las tablas o figuras de un apéndice no llevan su letra («Tabla A1»), o si los apéndices van antes de las referencias, solo se avisa.

### Qué hace (formato básico, títulos y limpieza)

- **Márgenes** de 2,54 cm en todas las secciones; las páginas horizontales siguen horizontales.
- **Una sola fuente** recomendada por APA en todo el trabajo, con su tamaño (salvo el código de programación, los símbolos y las ecuaciones).
- **Interlineado doble** sin espacio extra entre párrafos (en las tablas se respeta el sencillo, 1,5 o doble; las notas al pie no se tocan).
- **Texto alineado a la izquierda**, sin justificar ni guiones automáticos.
- **Sangría de 1,27 cm** en la primera línea de los párrafos de texto (no en títulos, listas, referencias, citas en bloque…).
- **Título del trabajo** repetido al principio del texto, centrado y en negrita: se copia tal cual de la portada.
- **Títulos** con los estilos de Word «Título 1» a «Título 5» y el formato APA de cada nivel (así aparecen en el panel de navegación), sin formato puesto a mano que lo contradiga; los de nivel 4 y 5, en la misma línea que su texto y terminados en punto (con el separador de estilo de Word: en LibreOffice se ven en su propia línea). Con mayúscula inicial en las palabras principales. Avisa de los apartados con un solo subapartado y de los títulos numerados.
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

## Servidor

```
uvicorn servidor.app:app
```

En http://localhost:8000/docs se puede subir un documento a mano. `POST /procesar` recibe el Word y
`POST /procesar-texto` el trabajo pegado como texto (un párrafo por línea, empezando por el título), con las
mismas opciones. Para guardar
lo que devuelve (copia corregida, informe y resumen) en `salida/`:

```
python herramientas/probar_servidor.py pruebas/01_ensayo_basico.docx
```

Con Docker (solo el núcleo y el servidor):

```
docker build -t apa-facil .
docker run --rm -p 8000:8000 apa-facil
```

## Web

Hace falta Node.js 20 o más reciente. La primera vez, desde la carpeta del proyecto:

```
cd web
npm install
```

Con el servidor en marcha (`uvicorn servidor.app:app`), en otra terminal:

```
cd web
npm run dev
```

La web queda en http://localhost:3000 y llama al servidor en http://localhost:8000. Tiene la página de
inicio, la herramienta (`/formatear`: subir el Word o pegar el texto, rellenar la portada, elegir las opciones y
descargar el resultado con su informe) y las páginas de privacidad y condiciones de uso. Si el servidor está en otra
dirección, cópiala en `web/.env.local` (hay un ejemplo en `web/.env.example`). El servidor solo responde a las webs
de la variable `APA_ORIGENES` (separadas por comas; por defecto, `http://localhost:3000`).
