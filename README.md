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

### Qué hace (formato básico y limpieza)

- **Márgenes** de 2,54 cm en todas las secciones; las páginas horizontales siguen horizontales.
- **Una sola fuente** recomendada por APA en todo el trabajo, con su tamaño (salvo el código de programación, los símbolos y las ecuaciones).
- **Interlineado doble** sin espacio extra entre párrafos (en las tablas se respeta el sencillo, 1,5 o doble; las notas al pie no se tocan).
- **Texto alineado a la izquierda**, sin justificar ni guiones automáticos.
- **Sangría de 1,27 cm** en la primera línea de los párrafos de texto (no en títulos, listas, referencias, citas en bloque…).
- **Número de página** arriba a la derecha en todas las páginas, también en la portada, y nada más en el encabezado.
- **Limpieza**: quita los párrafos vacíos usados para dejar espacio, los tabuladores o espacios usados como sangría y el color, el resaltado y el subrayado del texto (los enlaces se conservan). No toca la portada ni las tablas.

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

El verificador comprueba las reglas de formato básico (márgenes, fuente, interlineado, alineación, sangría y número de página) y lista cada incumplimiento con su sección, su número de párrafo y un extracto del texto. Lo que la herramienta deja así a propósito (la portada de la universidad, el código de programación…) aparece aparte, como excepción admitida. Termina con código 0 si todo cumple y 1 si no:

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
