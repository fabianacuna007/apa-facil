# APA Fácil: hoja de ruta de desarrollo

Este archivo es el plan completo del proyecto. Se trabaja **un paso cada vez**, en orden.

## Cómo usar este archivo

**Para Claude Code:**
- Trabaja solo en el paso que se te indique. No adelantes pasos de fases posteriores.
- Antes de programar, explica el plan del paso y espera confirmación.
- Un paso está terminado cuando cumple su criterio **"Hecho cuando"** y las pruebas pasan.
- Al terminar un paso: márcalo con `[x]`, añade una línea al **Registro de avances** (al final) y haz un commit.
- Si durante un paso descubres un problema que afecta a otro paso, anótalo en **Problemas abiertos**. No lo arregles por tu cuenta.

**Para empezar cada sesión**, el mensaje tipo es:

> Lee ROADMAP.md. Vamos con el paso X.Y. Explícame el plan antes de escribir código.

Conviene empezar una conversación nueva en Claude Code al cambiar de fase, para que el contexto no se llene de detalles antiguos.

---

## Arquitectura

### Idea central

El **núcleo** (el motor que formatea el Word) es independiente de la web. La web y la IA llegan después y solo llaman al núcleo. Así cada parte se puede probar por separado.

```
Word original
   │
   ▼
1. Abrir y diagnosticar      (qué formato tiene ahora)
   │
   ▼
2. Detectar la estructura    (qué es cada párrafo: título, cuerpo, lista, referencia…)
   │
   ▼
3. Aplicar reglas APA        (una regla = un módulo)
   │
   ▼
4. Verificar                 (¿el texto sigue intacto? ¿se cumplen las reglas?)
   │
   ▼
Word APA + informe de cambios y avisos
```

### Estructura de carpetas objetivo

```
apa-facil/
├── CLAUDE.md
├── ROADMAP.md
├── requirements.txt
├── apa_facil/                 # el núcleo
│   ├── __main__.py            # línea de comandos: python -m apa_facil archivo.docx
│   ├── pipeline.py            # decide el orden en que se aplican las reglas
│   ├── documento.py           # abrir, guardar y utilidades de XML
│   ├── estructura.py          # clasificador de párrafos (fase 3)
│   ├── informe.py             # registro de cambios y avisos
│   ├── reglas/                # una regla APA por archivo
│   │   ├── margenes.py
│   │   ├── fuente.py
│   │   ├── interlineado.py
│   │   ├── alineacion.py
│   │   ├── sangria.py
│   │   └── paginacion.py
│   └── revision/              # análisis de citas (fase 7)
├── herramientas/
│   ├── inspeccionar.py        # muestra el formato actual de un Word
│   └── verificar.py           # comprueba si un Word cumple APA
├── tests/                     # pruebas automáticas (pytest)
├── docs/                      # documentación de apoyo (comparativa de normas APA)
├── pruebas/                   # documentos de prueba: NO MODIFICAR
└── salida/                    # documentos generados
```

### Convenciones

- Cada regla es una función `aplicar(doc, opciones, informe, estructura=None)`. Recibe el documento, las opciones del usuario, el informe donde anota lo que hizo y la estructura: la clasificación de los párrafos (`estructura.clasificar`), que el pipeline hace **una vez sobre el original, antes de aplicar ninguna regla** (paso 3.4: después de la regla de fuente ya no se ve, por ejemplo, qué títulos tenían la letra más grande). Las reglas que no la necesitan la ignoran; si una que la necesita se llama sin ella (en sus pruebas), clasifica el documento tal como lo recibe.
- El informe distingue tres cosas:
  - **cambios**: lo que la herramienta modificó;
  - **avisos**: lo que el usuario debe revisar a mano;
  - **errores**: lo que no se pudo procesar.
- Todos los mensajes del informe van en español y están pensados para un estudiante, no para un programador.
- `python-docx` no cubre todo (campos de número de página, notas al pie, algunos ajustes de sección). Para eso se trabaja directamente con el XML mediante `lxml`, siempre encapsulado en `documento.py`.

---

## Fase 0: Preparación

**Por qué primero:** antes de cambiar nada hay que demostrar que abrir y guardar un Word con nuestras herramientas no lo estropea. Si eso falla, todo lo demás falla.

- [x] **0.1 Repositorio.** `git init`, `.gitignore` (salida/, entornos virtuales, cachés) y primer commit.
- [x] **0.2 Entorno.** Entorno virtual de Python, `requirements.txt` con `python-docx`, `lxml` y `pytest`, e instrucciones de instalación en un `README.md` corto.
- [x] **0.3 Esqueleto.** Crear la estructura de carpetas de arriba con archivos vacíos o mínimos.
- [x] **0.4 Ida y vuelta.** Línea de comandos que abre un .docx y lo guarda en `salida/` **sin cambios**.
- [x] **0.5 Prueba de integridad.** Una función que extrae todo el texto de un documento (párrafos, tablas y notas al pie), en orden. Una prueba compara el original con la copia del paso 0.4: el texto debe ser idéntico, y el número de tablas, imágenes, notas al pie y secciones también.
- [x] **0.6 Inspector.** `herramientas/inspeccionar.py`. Muestra, por documento:
  - márgenes por sección;
  - fuentes usadas y cuántas veces;
  - interlineados;
  - sangrías;
  - estilos de párrafo;
  - qué hay en encabezados y pies.

  Sirve para diagnosticar los documentos de prueba.

**Hecho cuando:** los 5 documentos de `pruebas/` hacen la ida y vuelta, pasan la prueba de integridad, y el inspector muestra un resumen legible de cada uno.

---

## Fase 1: Formato básico

**Por qué:** es el núcleo del producto y lo que el usuario nota primero. Son reglas deterministas: se pueden programar y comprobar con exactitud.

- [x] **1.1 Márgenes.** 2,54 cm en los cuatro lados, en **todas** las secciones. Las páginas en horizontal conservan su orientación.
- [x] **1.2 Fuente.**
  - Fuentes recomendadas por APA, cada una con su tamaño: Aptos 12, Calibri 11, Arial 11, Lucida Sans Unicode 10, Times New Roman 12, Georgia 11 y Computer Modern 10.
  - Una sola fuente para todo el documento, títulos y portada incluidos. Se elige así (decisión del propietario):
    - si la fuente principal del documento (la que tiene más letras) es una de las recomendadas, **se mantiene** y solo se ajusta el tamaño;
    - si no, se cambia por la recomendada **más parecida**, con su tamaño;
    - la opción `--fuente` (paso 1.7) tiene prioridad sobre las dos reglas anteriores.
  - Aplicarla en los estilos (Normal y los demás) **y** en el formato directo de cada fragmento de texto, incluidos los atributos de fuente para otros alfabetos (`ascii`, `hAnsi`, `eastAsia`, `cs`).
  - Tablas: misma fuente y **mismo tamaño** que el resto.
  - Excepciones que no se tocan:
    - código de programación en fuente monoespaciada (APA pide, por ejemplo, Courier New 10 o Lucida Console 10);
    - fuentes de símbolos y ecuaciones;
    - el tamaño de las notas al pie (formato por defecto de Word).
  - El informe no dice que la fuente original "no está permitida" (APA presenta su lista como recomendaciones), sino que se ha cambiado por la recomendada más parecida.
- [x] **1.3 Interlineado.** Doble, con 0 pt de espacio antes y después del párrafo. Excepciones:
  - dentro de tablas: se respeta el interlineado que tengan si es sencillo, 1,5 o doble (APA admite los tres); si es otro, se pone sencillo. El número, el título y la nota de la tabla van a doble espacio;
  - notas al pie: formato por defecto de Word.
- [x] **1.4 Alineación.** Izquierda, sin justificar. Las tablas y las imágenes centradas se quedan como están.
- [x] **1.5 Sangría (versión simple).** 1,27 cm en la primera línea de los párrafos de cuerpo de texto, no en títulos, listas, tablas ni párrafos centrados. En la fase 3 se refina con el detector de estructura.
- [x] **1.6 Número de página.**
  - Campo PAGE alineado a la derecha en el encabezado de todas las secciones, incluida la primera página.
  - Hay que tener en cuenta la opción "primera página diferente" y los encabezados de páginas pares.
  - Quitar los números de página que ya existan en pies y encabezados.
  - Si había otro texto en el encabezado, quitarlo y registrar un aviso, porque APA de estudiante no lleva encabezado de página.
- [x] **1.7 Línea de comandos completa.** `python -m apa_facil entrada.docx --fuente "Arial"` genera `salida/entrada_APA.docx` y muestra un resumen del informe.

**Hecho cuando:** los 5 documentos se procesan sin errores, pasan la prueba de integridad, y al abrirlos en Word las reglas 1.1 a 1.6 se cumplen a simple vista.

---

## Fase 2: Verificación e informe

**Por qué:** sin un verificador automático, cada cambio obliga a abrir los Word a mano para ver si algo se ha roto. Con él, las regresiones se detectan solas.

- [x] **2.1 Verificador APA.** `herramientas/verificar.py` lee un .docx y comprueba cada regla de la fase 1. Devuelve la lista de incumplimientos con su ubicación (sección, número de párrafo y un extracto del texto).
- [x] **2.2 Pruebas automáticas.** `pytest` que, para cada documento de `pruebas/`, hace tres cosas:
  1. lo procesa;
  2. comprueba la integridad del texto y de los elementos;
  3. pasa el verificador sin incumplimientos.
- [x] **2.3 Informe estructurado.** Informe en JSON (para la web futura) y en texto legible en la terminal. Ejemplo de aviso: *"El encabezado contenía el texto 'Facultad de Ciencias…'. Se ha eliminado porque APA de estudiante no lo usa."*
- [x] **2.4 Revisión visual (opcional).** Si LibreOffice está instalado, convertir la salida a PDF para revisarla sin abrir Word.

**Hecho cuando:** `pytest` pasa en verde con los 5 documentos y cada procesado genera su informe.

---

## Fase 3: Detección de estructura y limpieza

**Por qué:** es la fase más difícil y la que diferencia una herramienta buena de una mala. Para aplicar bien la sangría, los títulos o las referencias, primero hay que saber qué es cada párrafo, aunque el estudiante no haya usado los estilos de Word.

- [x] **3.1 Clasificador de párrafos.** `estructura.py` asigna a cada párrafo un tipo y una confianza (alta, media o baja). Tipos:
  - `portada`, `indice`, `titulo` (con su nivel), `cuerpo`, `lista`;
  - `leyenda_tabla`, `leyenda_figura`, `cita_bloque`;
  - `titulo_referencias`, `referencia`, `anexo`, `vacio`;
  - `imagen`: párrafo que solo contiene una imagen (decisión del propietario: así la limpieza de 3.5 no lo borra como vacío).

  Cada párrafo lleva también su **zona** (decisión del propietario): `portada`, `indice`, `resumen`, `cuerpo`, `referencias` o `anexos`. Así el párrafo del resumen y la línea de palabras clave son `cuerpo` en la zona "resumen", y el paso 6.7 sabe qué tablas están en un anexo.
- [x] **3.2 Títulos falsos.** Detectar como título un párrafo corto, en negrita o en mayúsculas, sin punto final, a veces con numeración ("1.", "1.1."), aunque use el estilo Normal. Deducir su nivel por la numeración, el tamaño y el formato.
- [x] **3.3 Modo diagnóstico.** La opción `--diagnostico` imprime una tabla con párrafo, tipo, confianza y extracto. Sirve para depurar el clasificador con los documentos de prueba.
- [x] **3.4 Sangría definitiva.** Rehacer la regla 1.5 usando el clasificador: solo los párrafos `cuerpo` llevan sangría de primera línea.
- [x] **3.5 Limpieza de formato.** Todo lo que se haga aquí queda registrado en el informe.
  - Eliminar los párrafos vacíos usados como espaciado. **Excepción:** la línea en blanco entre el título y el autor de la portada es obligatoria y no se borra.
  - Eliminar los tabuladores y espacios al inicio de párrafo usados como sangría.
  - Quitar colores de texto, resaltados y subrayados decorativos. Los hipervínculos se mantienen.
- [x] **3.6 Problemas que solo se avisan, sin corregirlos.**
  - Saltos de línea manuales dentro de un párrafo.
  - Viñetas escritas a mano ("•", "-").
  - Dobles espacios entre palabras.
  - Texto en otra fuente que parezca pegado de internet.

**Hecho cuando:** el clasificador acierta el tipo de al menos el 95 % de los párrafos de los 5 documentos (revisado con el modo diagnóstico), y los avisos de 3.6 aparecen en el informe del documento 04.

---

## Fase 4: Títulos APA (cinco niveles)

**Por qué:** los títulos son lo más visible después de la fuente, y dependen del clasificador de la fase 3.

- [x] **4.1 Estilos de título.** Modificar los estilos de título de Word (Título 1 a Título 5) con el formato APA y asignar cada título detectado a su estilo. Usar estilos en vez de formato directo mantiene el panel de navegación y el índice de Word, y es lo que recomienda APA por accesibilidad. Avisar cuando un apartado tenga un solo subapartado, porque APA lo desaconseja.
- [x] **4.2 Formato por nivel:**
  - Nivel 1: centrado, negrita.
  - Nivel 2: a la izquierda, negrita.
  - Nivel 3: a la izquierda, negrita y cursiva.
  - Nivel 4: con sangría, negrita, terminado en punto, y el texto sigue en la misma línea.
  - Nivel 5: con sangría, negrita y cursiva, terminado en punto, y el texto sigue en la misma línea.
- [x] **4.3 Títulos en mayúsculas y numeración.** Cambiar un texto de MAYÚSCULAS a mayúscula inicial o quitar la numeración ("1. Introducción" → "Introducción") **modifica el texto**. Por eso:
  - por defecto, solo se avisa;
  - con una opción explícita del usuario, se cambia y se registra.
- [ ] **4.4 Título del trabajo.** En la primera página del cuerpo va el título del trabajo, centrado y en negrita, como nivel 1. El primer apartado no lleva el rótulo "Introducción".
- [ ] **4.5 Niveles 4 y 5 en línea.** Unir el título con el párrafo siguiente es delicado. Implementarlo al final de la fase, con una prueba específica.

**Hecho cuando:** los títulos de los documentos 02, 03 y 05 quedan con el formato de su nivel y el panel de navegación de Word sigue funcionando.

---

## Fase 5: Portada de estudiante y secciones iniciales

- [ ] **5.1 Datos de portada.** Los datos llegan por opciones o por un archivo `datos.json`: título, autor o autores, departamento y universidad, código y nombre de la asignatura, profesor y fecha de entrega.
- [ ] **5.2 Generar la portada APA.**
  - El título, centrado, en negrita y con mayúscula inicial en las palabras principales, se coloca 3 o 4 líneas por debajo del margen superior.
  - Después, una línea en blanco a doble espacio (obligatoria; ver la excepción del paso 3.5).
  - Debajo, centrados y a doble espacio: autor, afiliación, asignatura, profesor y fecha.
  - La fecha va en formato español: "16 de octubre de 2026".
  - La portada termina con un salto de página.
- [ ] **5.3 Portada existente.** Detectarla (texto centrado antes del primer salto de página o en la primera sección) y ofrecer dos opciones:
  - **sustituirla** por la portada APA (opción por defecto);
  - **mantenerla**, porque muchas universidades exigen su propia portada con logo.
- [ ] **5.4 Índice.** APA de estudiante no lleva índice, pero muchas universidades españolas lo exigen. Por defecto se mantiene, y se puede quitar con una opción.
- [ ] **5.5 Resumen.** Si existe, el rótulo "Resumen" va centrado y en negrita, y el párrafo sin sangría. Las palabras clave van en cursiva en una línea con sangría.

**Hecho cuando:** el documento 01 recibe una portada APA completa a partir de `datos.json`, y en el documento 05 se pueden probar las dos opciones de portada.

---

## Fase 6: Tablas, figuras, citas en bloque y referencias

- [ ] **6.1 Tablas.**
  - Encima de la tabla: "Tabla N" en negrita y, en la línea siguiente, el título en cursiva.
  - Debajo de la tabla, si hay nota: "*Nota.*" seguido del texto.
  - Detectar las leyendas existentes, aunque estén debajo de la tabla, y recolocarlas.
  - Opcional: estilo de tabla APA, con bordes solo horizontales y sin colores de fondo.
  - Si una tabla ocupa más de una página, repetir la fila de encabezados en cada página.
  - El aviso de tabla demasiado ancha del paso 1.1 debe sugerir ponerla en una página horizontal, que es lo que recomienda APA.
  - Avisar (sin cambiar nada) si falta la línea en blanco entre la tabla y el texto, o si la tabla queda en mitad de una página en vez de arriba o abajo.
- [ ] **6.2 Figuras.** El mismo esquema: "Figura N" en negrita y el título en cursiva encima de la imagen. Los mismos avisos de colocación que en las tablas.
- [ ] **6.3 Citas en bloque** (40 palabras o más).
  - Sangría izquierda de 1,27 cm en todo el bloque, sin sangría de primera línea y con doble espacio.
  - Quitar la cursiva decorativa.
  - Si la cita lleva comillas, solo se avisa: quitarlas modificaría el texto.
- [ ] **6.4 Lista de referencias.**
  - Detectar la sección aunque se llame "Bibliografía", "Webgrafía" o "Fuentes".
  - Empieza en página nueva, con el rótulo "Referencias" centrado y en negrita ("Referencia" si solo hay una). APA no admite "Bibliografía", "Obras citadas" ni "Lista de referencias".
  - Sangría francesa de 1,27 cm y doble espacio.
  - Los enlaces (DOI y URL) se dejan como estén: APA admite tanto el azul subrayado como el texto normal.
- [ ] **6.5 Orden alfabético.** Reordenar las referencias es una opción activada por defecto y queda registrada. Si hay varias listas (por ejemplo, una bibliografía y una webgrafía), se avisa y se propone unirlas.
- [ ] **6.6 Notas al pie.** Misma fuente que el documento; el resto del formato no se toca.
- [ ] **6.7 Apéndices.**
  - Cada apéndice empieza en página nueva, al final del trabajo.
  - Rótulo centrado y en negrita: "Apéndice" si solo hay uno; "Apéndice A", "Apéndice B"… si hay varios.
  - En la línea siguiente, el título del apéndice, también centrado y en negrita.
  - Sus tablas y figuras se numeran con la letra del apéndice, sin punto: "Tabla A1", "Figura B2".
  - Detectar la sección aunque se llame "Anexo" o "Anexos". Cambiar el rótulo modifica el texto, así que se trata como en el paso 4.3: por defecto solo se avisa.

**Hecho cuando:** las tablas de los documentos 02, 04 y 05, las figuras del 03, la cita en bloque del 05, las referencias de los cinco documentos y los apéndices que haya quedan en formato APA, y la prueba de integridad sigue pasando (permitiendo el cambio de rótulo y el nuevo orden).

---

## Fase 7: Revisión de citas y referencias (solo análisis)

**Por qué:** es la segunda función del producto. Aquí la herramienta **no modifica nada**: solo detecta y explica.

**Ojo:** los ejemplos de normas-apa.org **no sirven** como respuestas correctas para las pruebas, porque mezclan la 6.ª y la 7.ª edición y tienen errores (ver `docs/comparativa_normas_apa.md`). Usar los ejemplos de la web oficial de APA.

- [ ] **7.1 Detector de citas en el texto.** Formatos a reconocer:
  - `(Autor, año)` y `Autor (año)`;
  - dos autores con "y" y varios autores con "et al.";
  - varias citas en un paréntesis separadas por ";";
  - páginas con "p." y "pp.";
  - "s. f." para obras sin fecha;
  - autores que son organizaciones.
- [ ] **7.2 Lector de referencias.** Extraer de cada entrada: autores, año, título, fuente, volumen, número, páginas y DOI o URL. Si no se puede leer, se marca como "no reconocida", nunca se adivina.
- [ ] **7.3 Cruce de citas y referencias.**
  - Citas sin referencia en la lista.
  - Referencias que no se citan en el texto.
  - Años que no coinciden.
  - Autores escritos de forma distinta.
- [ ] **7.4 Errores de forma.** Ejemplos:
  - falta la coma antes del año;
  - "et. al" en lugar de "et al.";
  - "s.f" en lugar de "s. f.";
  - "&" en lugar de "y" (en APA en español);
  - DOI en formato antiguo ("doi:") en lugar de `https://doi.org/`;
  - ciudad de edición, que APA 7 ya no incluye;
  - "Recuperado de" innecesario;
  - falta el año;
  - falta el espacio en "p.45".
- [ ] **7.5 Avisos con ubicación.** Cada aviso indica el párrafo, un extracto, qué está mal y cómo se corrige.
- [ ] **7.6 Comentarios en Word (opcional).** Insertar los avisos como comentarios en el documento de salida, para que el estudiante los vea en su sitio.

**Casos esperados en los documentos de prueba** (sirven como prueba automática):

| Documento | Errores que debe detectar |
|---|---|
| 01 | Castell (2011) se cita pero no está en las referencias. "Lefèvre, A. 2009." tiene el año sin paréntesis. La ciudad de edición aparece en Ramírez Soto y en Lefèvre. Hay una URL suelta sin datos. Las referencias no están en orden alfabético. |
| 02 | Ninguno grave: es el caso de control. |
| 03 | Campos y Vidal usan "&" en lugar de "y". |
| 04 | "(Martínez 2019)": falta la coma. "Pérez et. al": mal escrito. "Gómez (s.f)": formato incorrecto. Gómez no está en las referencias. La Webgrafía tiene URLs sueltas y referencias incompletas. |
| 05 | Herranz no tiene año. Ruiz Montero (2023) se cita pero no está en las referencias. Torres (2019) está en las referencias pero no se cita. "doi:" está en formato antiguo. Falta el espacio en "p.45". Molina está fuera del orden alfabético. |

**Hecho cuando:** se detectan todos los casos esperados de la tabla y no aparecen avisos falsos en el documento 02.

---

## Fase 8: API (servidor)

**Por qué:** convierte el núcleo en un servicio que la web puede usar.

- [ ] **8.1 Servidor FastAPI.** Endpoint `POST /procesar`: recibe el archivo y las opciones, y devuelve el Word procesado y el informe en JSON.
- [ ] **8.2 Seguridad de entrada.**
  - Solo se aceptan .docx, y se comprueba que el archivo sea un ZIP válido aunque tenga esa extensión.
  - Tamaño máximo de 10 MB.
  - Protección contra archivos comprimidos maliciosos (bombas ZIP).
  - Tiempo máximo de procesado.
- [ ] **8.3 Privacidad.**
  - Los archivos se guardan solo en una carpeta temporal y se borran al terminar.
  - Nunca se guarda el contenido de los documentos en los registros (logs).
- [ ] **8.4 Pruebas del endpoint** con los 5 documentos y con archivos inválidos (un PDF renombrado, un archivo vacío, uno demasiado grande).
- [ ] **8.5 Dockerfile** para poder desplegarlo.

**Hecho cuando:** se puede subir un documento al endpoint en local y descargar el resultado, y las pruebas con archivos inválidos devuelven errores claros en español.

---

## Fase 9: Web

- [ ] **9.1 Proyecto Next.js** con Tailwind CSS, conectado a la API.
- [ ] **9.2 Estética.** Definir colores, tipografías y componentes a partir de la imagen de referencia del proyecto (pendiente de añadir a la carpeta `diseno/`).
- [ ] **9.3 Página de inicio:** qué hace la herramienta, en una frase, y un botón para empezar.
- [ ] **9.4 Subida del documento:** arrastrar y soltar, o seleccionar archivo.
- [ ] **9.5 Formulario:** datos de portada y opciones (fuente, mantener portada, mantener índice, ordenar referencias).
- [ ] **9.6 Pantalla de procesado y resultado:** botón de descarga e informe agrupado en "Lo que hemos cambiado" y "Lo que debes revisar tú".
- [ ] **9.7 Modo "pegar texto":** convierte el texto pegado en un .docx y lo procesa con el mismo núcleo.
- [ ] **9.8 Diseño adaptable a móvil, accesibilidad básica** (contraste, teclado, etiquetas) **y páginas legales.**

**Hecho cuando:** un estudiante puede, desde el navegador, subir su trabajo, rellenar la portada y descargar el resultado sin ayuda.

---

## Fase 10: Asistencia con IA

**Regla absoluta:** la IA nunca inventa autores, años, títulos, DOI ni datos de ninguna fuente, y nunca modifica el texto del estudiante.

- [ ] **10.1 Explicaciones.** Para cada aviso de la fase 7, una explicación sencilla de qué está mal y por qué, con un ejemplo correcto.
- [ ] **10.2 Ayuda para completar referencias.** La IA pide al estudiante los datos que faltan y le devuelve la referencia bien formateada **solo con los datos que él aporte**.
- [ ] **10.3 Búsqueda de datos reales (opcional).** Consultar un DOI en Crossref, una base de datos pública de publicaciones académicas. El resultado se muestra como sugerencia y el estudiante lo confirma.
- [ ] **10.4 Control de la salida.** Validar que la respuesta de la IA no contiene datos que no estaban en la entrada o en Crossref.
- [ ] **10.5 Costes.** Medir el coste por documento, fijar límites por usuario y guardar en caché las explicaciones repetidas.

**Hecho cuando:** los avisos de los documentos 01, 04 y 05 tienen explicaciones correctas, y en ninguna prueba la IA introduce datos inventados.

---

## Fase 11: Lanzamiento

- [ ] **11.1 Despliegue.** Web en Vercel, API en Render o Railway, dominio propio y HTTPS.
- [ ] **11.2 Cumplimiento legal (RGPD).**
  - Política de privacidad y condiciones de uso.
  - Lista de proveedores que tratan datos (hosting, proveedor de IA).
  - Plazo de borrado de archivos.
  - Aviso de cookies si se usa analítica.
- [ ] **11.3 Monitorización** de errores (por ejemplo, Sentry) sin enviar el contenido de los documentos, y analítica respetuosa con la privacidad (por ejemplo, Plausible).
- [ ] **11.4 Prueba con 10 a 20 estudiantes reales:** recoger sus documentos (con permiso), los fallos y las opiniones. Cada documento problemático se añade a `pruebas/`.

---

## Fase 12: Monetización (solo después de validar)

- [ ] **12.1 Límites de la versión gratuita** (documentos por día o por mes). Requiere cuentas de usuario o algún otro control.
- [ ] **12.2 Pagos con Stripe:** pago por documento y suscripción prémium.
- [ ] **12.3 Anuncios.** Valorar si compensan: pueden restar confianza en una herramienta que maneja trabajos personales.

---

## Fuera de alcance por ahora

- Archivos .doc antiguos, PDF, Google Docs y LaTeX.
- APA en su versión profesional, con encabezado corto (*running head*).
- Otras normas (MLA, Chicago, Vancouver, ISO 690). Podrían ser una ampliación futura.

## Decisiones pendientes (las toma el propietario del proyecto)

- [x] Fuente por defecto. **Decidido:** se mantiene la del documento si es una de las recomendadas por APA; si no, se cambia por la recomendada más parecida (ver el paso 1.2).
- [ ] Si la portada original se sustituye o se mantiene por defecto.
- [ ] Si el orden alfabético de las referencias se aplica automáticamente o solo se avisa.
- [x] Si los títulos en mayúsculas o numerados se corrigen por defecto. **Decidido:** por defecto solo se avisa; se corrigen con dos opciones independientes, `--quitar-numeracion` y `--quitar-mayusculas` (paso 4.3).
- [ ] Nombre de dominio.

---

## Especificación APA 7 de estudiante (referencia)

Fuente: la web oficial de APA (<https://apastyle.apa.org/>), que manda sobre cualquier otra. La comparación con normas-apa.org está en `docs/comparativa_normas_apa.md`.

| Elemento | Regla |
|---|---|
| Márgenes | 2,54 cm en los cuatro lados |
| Fuentes recomendadas | Aptos 12, Calibri 11, Arial 11, Lucida Sans Unicode 10, Times New Roman 12, Georgia 11, Computer Modern 10. La misma en todo el trabajo, también en títulos y tablas. |
| Excepciones de fuente | Código de programación en fuente monoespaciada (p. ej., Courier New 10 o Lucida Console 10); texto dentro de las imágenes de figuras (sin serifa, de 8 a 14 pt); notas al pie con el formato por defecto |
| Interlineado | Doble en todo el documento, incluidas la portada y las referencias. Sin espacio extra antes ni después de los párrafos. |
| Excepciones de interlineado | Interior de tablas y figuras (sencillo, 1,5 o doble; su número, título y nota, a doble espacio), notas al pie (formato por defecto), una línea en blanco entre el título y el autor en la portada, y más espacio alrededor de ecuaciones |
| Alineación | Izquierda, con el borde derecho irregular. Sin justificar y sin guiones automáticos. |
| Sangría | 1,27 cm en la primera línea de los párrafos de cuerpo |
| Sin sangría | Portada, rótulos de sección, títulos de nivel 1 a 3, párrafo del resumen, títulos y notas de tablas y figuras |
| Número de página | Arriba a la derecha, en todas las páginas, incluida la portada |
| Encabezado de página | No se usa en trabajos de estudiante |
| Portada | Título en negrita centrado, 3 o 4 líneas por debajo del margen superior; después, una línea en blanco a doble espacio. Debajo, centrados: autor, afiliación, asignatura, profesor y fecha. |
| Título nivel 1 | Centrado, negrita |
| Título nivel 2 | A la izquierda, negrita |
| Título nivel 3 | A la izquierda, negrita y cursiva |
| Título nivel 4 | Con sangría, negrita, terminado en punto, el texto sigue en la misma línea |
| Título nivel 5 | Con sangría, negrita y cursiva, terminado en punto, el texto sigue en la misma línea |
| Tablas | "Tabla N" en negrita encima; título en cursiva en la línea siguiente; "*Nota.*" debajo |
| Figuras | "Figura N" en negrita encima; título en cursiva en la línea siguiente; "*Nota.*" debajo |
| Cita en bloque | 40 palabras o más. Sangría izquierda de 1,27 cm en todo el bloque, sin comillas; la cita va después del punto final. |
| Referencias | Página nueva; rótulo "Referencias" (o "Referencia" si solo hay una) centrado y en negrita; orden alfabético; sangría francesa de 1,27 cm; doble espacio |
| Apéndices | Cada uno en página nueva; rótulo "Apéndice" (o "Apéndice A", "B"…) centrado y en negrita; título en la línea siguiente, centrado y en negrita; tablas y figuras numeradas "Tabla A1", "Figura B2" |
| Títulos | Sin números ni letras delante; sin líneas en blanco antes ni después |
| APA en español | "y" entre autores (no "&"); "s. f." para obras sin fecha; "et al."; "p." y "pp."; DOI con formato `https://doi.org/...` |

---

## Problemas abiertos

_(Claude Code anota aquí lo que descubra y afecte a otros pasos.)_

- **2.4 Revisión visual.** LibreOffice no está instalado en este equipo. Word sí, y se puede convertir a PDF por COM desde PowerShell, pero solo con llamadas dinámicas (`InvokeMember`): el acceso normal falla con el error `E_NOINTERFACE`. Funciona `[System.__ComObject].InvokeMember(...)`, sin pasar los objetos de Word por funciones de PowerShell (si no, `InvokeMember` no encuentra el método). Así se contaron las páginas en el paso 1.1. Ojo también: PowerShell no distingue mayúsculas en los nombres de variable (`$t` y `$T` son la misma). Y si un script falla con un documento abierto, el Word invisible se queda en segundo plano y bloquea la copia de `salida/` (el programa dice "Seguramente lo tienes abierto en Word"): hay que cerrar los documentos y llamar a `Quit` sin guardar en un `finally`, y comprobar con `Get-CimInstance Win32_Process` que no quedan procesos `WINWORD.EXE /Automation`. **Resuelto en 2.4:** `herramientas/a_pdf.py` usa Word así (y LibreOffice si estuviera instalado). Ojo: el número de páginas que da Word (`ComputeStatistics`) salía una página más alto que el PDF en 4 de los 5 documentos; las páginas se cuentan en el PDF.
- **2.1 Verificador y la fuente que informa Word.** Al pedir a Word la fuente de todo el documento (`Content.Font`), devuelve "mezclada" en los documentos 02, 04 y 05 aunque todos sus caracteres, párrafos y celdas tengan la misma fuente y el mismo tamaño (comprobado carácter a carácter en el paso 1.2). Pasa con las tablas de esos documentos, no con tablas creadas por python-docx. El verificador no debe fiarse de la fuente de rangos grandes: hay que comprobar párrafo a párrafo o leer el XML. Además, Word cuenta como párrafos las marcas de fin de fila de las tablas (una por fila): no tienen texto y no hay que tenerlas en cuenta. **Resuelto en 2.1:** el verificador lee el XML fragmento a fragmento, sin usar Word.
- **3.1, 3.2 y 3.4 Clasificador.** **Resuelto en 3.4:** la regla de sangría usa el clasificador y se han borrado de `sangria.py` sus detectores propios. Queda `estructura.py` usando `_salto_antes` y `_salto_en` de `reglas/alineacion.py` (podrían pasar a `documento.py`), y la regla de alineación (1.4) sigue con su detector sencillo de portada, `parece_portada` (ver 5.3).
- **5.4 Índice.** Si el documento pide a Word que actualice los campos al abrirse (`w:updateFields`), Word rehace el índice con sus estilos "toc 1" a "toc 9" ("TDC" en español). El paso 1.3 los define a doble espacio cuando hay un índice; el 5.4 debe contar con ello.
- **5.3 Portada existente.** `reglas/alineacion.parece_portada` ya detecta, de forma sencilla, lo que parece una portada: los párrafos del principio hasta el primer salto de página o de sección, o hasta el primer párrafo de 25 palabras o más (en el paso 1.5 se bajó de 40 a 25 porque en el 04 la portada se "comía" el primer párrafo de texto). Se usa para no tocar lo alineado a la derecha en ella ni ponerle sangría. El paso 5.3 puede partir de ahí, y cuando trate la portada debe revisar esos párrafos (en el 05, las líneas "Autora", "Tutor" y "Convocatoria").
- **3.2 Títulos falsos (desde 3.1).** **Resuelto en 3.2:** el nivel sale de la numeración o del formato, y "Factores de riesgo:" del 04 ya es nivel 2 (`PENDIENTES` de `tests/test_estructura.py` queda vacío). Límites conocidos, para la fase 4: los dos puntos finales ("Factores de riesgo:") cuentan como final de frase, así que ese título se reconoce solo por la negrita; una frase corta en negrita que acaba en punto se toma por título con confianza baja (los títulos APA de nivel 4 y 5 acaban en punto, pero van en la misma línea que el texto: paso 4.5). Los títulos de Word no se tocan aunque salten un nivel (de 1 a 3): el paso 4.1 debería avisar. **Resuelto en 4.1:** se avisa del salto de nivel.
- **3.5 Párrafos vacíos con saltos (desde 3.1).** **Resuelto en 3.5:** no se borran los vacíos que llevan un salto de página o de sección, una imagen, un marcador, una nota o un campo, ni los que están en la portada o junto a una tabla o imagen. **Para 6.4 y 6.7:** un párrafo vacío que solo lleva un salto de página (en el 05, antes de la Bibliografía y del Anexo) deja una línea en blanco al principio de la página siguiente. Al hacer que las referencias y los anexos empiecen en página nueva, conviene sustituirlo por "salto de página anterior" en el rótulo y borrar el vacío (anotándolo como cambio previsto).
- **Verificador y clasificador (2.1 → fase 4).** **Resuelto en 3.4** para la sangría: el verificador decide qué párrafos llevan sangría con el clasificador (a través de `sangria.motivos_sin_sangria`), y la numeración es una sola (desde 1 y contando los párrafos de las tablas: `Parrafo.numero` en el clasificador y el diagnóstico). **Pendiente para la fase 4:** el verificador solo tiene la copia y la clasifica a ella, mientras que el programa clasificó el original. En el 04 no coinciden: sin la letra más grande (la regla de fuente la iguala), «¿Qué son las redes sociales?» parece portada y los niveles de sus títulos se invierten. Para la sangría no importa en los 5 documentos (coinciden los párrafos de texto, y si no coincidieran el verificador daría incumplimientos en `test_cinco_documentos`), pero sí importará para comprobar títulos. El paso 4.1 lo resuelve al dar a cada título su estilo de Word: la copia llevará su estructura escrita. **Resuelto en 4.1:** la copia del 04 se clasifica igual que el original (lo comprueba `test_cinco_documentos.py`). El verificador también usa `parece_portada` y `es_codigo` de las reglas.
- **Comodines en el verificador (2.1).** `python herramientaserificar.py salida\*_APA.docx` no expande el comodín en cmd ni en PowerShell (dice que no encuentra «*_APA.docx»); la línea de comandos principal sí lo hace con `expandir_comodines`. Visto en el paso 3.4; arreglo de una línea, para cuando se decida.
- **Resultado del verificador en el informe (2.3 → fase 8 o 9).** El JSON no dice si la copia cumple APA, aunque sería útil para la web ("tu documento cumple"). El verificador está en `herramientas/`, y el núcleo no debe depender de las herramientas: incluirlo obliga a mover el verificador a `apa_facil/`. Lo decide el propietario.
- **Privacidad del informe (2.3 → 8.3).** El JSON contiene extractos del texto del estudiante (por ejemplo, el texto que había en el encabezado). En el servidor no debe acabar en los registros (logs) y debe borrarse con el documento.
- **4.1 Títulos que se quedan solos al final de la página.** En el PDF del 03 (paso 2.4), el título "2.1. Fuentes de energía" queda al final de la página 1 y su texto empieza en la página 2. Al definir los estilos de título, conviene activar "Conservar con el siguiente" (`w:keepNext`), como hacen los títulos de Word por defecto. **Resuelto en 4.1:** los estilos Título 1 a 5 lo llevan, y Word confirma que en ninguna de las 5 copias queda un título solo al final de una página.
- **Integridad y cambios previstos (0.5).** **Resuelto en 3.5 para borrar y recortar:** las reglas anotan cada cambio de contenido con `informe.cambio_previsto(numero, accion, quitado)` (número del párrafo en el original) y `documento.comparar` acepta exactamente esos: borrar un párrafo sin texto y quitar espacio en blanco del principio; cualquier otra diferencia sigue impidiendo entregar la copia. Los cambios previstos no van en el JSON. **Resuelto en 4.3** para los títulos: acciones `quitar_numeracion` y `cambiar_mayusculas`. **Pendiente:** los pasos 6.4 (rótulo "Referencias") y 6.5 (orden de las referencias) necesitarán acciones nuevas en `CambioPrevisto` (cambiar un texto por otro concreto, reordenar párrafos), cada una con su comprobación estricta.
- **Limpieza: lo que no se toca (3.5 → 4.1, 5.3, 6.1).** Decisión del propietario: la limpieza no toca la portada (vacíos ni colores: el 04 conserva su título rojo y subrayado y sus datos en gris; lo decide el 5.3), los colores de las tablas (los encabezados del 02 son blancos sobre fondo oscuro: van con sus fondos, en el 6.1), ni los estilos (los títulos azules del 02, 03 y 05 vienen de los estilos «Título 1» a «Título 4», y del estilo «Title» en la portada del 02: los rehace el 4.1). Tampoco toca los cuadros de texto, el código, las notas al pie (6.6) ni los encabezados y pies.

- **Títulos: lo que queda para 4.2 y 6.4 (desde 4.1).** La regla 4.1 pone el formato en los estilos, pero no quita el formato directo de los párrafos de título, que manda sobre el estilo: los títulos falsos del 01, 04 y 05 conservan su negrita puesta a mano (no estorba) y, si un documento trajera una cursiva, un tamaño o un centrado puestos a mano en un título, se verían. El 4.2 debe quitar lo que contradiga el formato del nivel y comprobarlo párrafo a párrafo en el verificador (ahora solo comprueba los estilos y que cada título tenga el suyo). **Resuelto en 4.2.** Los rótulos de referencias escritos a mano («BIBLIOGRAFÍA» del 01 y del 05, «Webgrafía:» del 04) no tienen estilo y no salen en el panel de navegación; los que ya usaban Título 1 («Referencias» del 02, «Bibliografía» del 03, el anexo del 05) sí, y quedan centrados y en negrita. El 6.4 y el 6.7 deberían darles a todos el mismo estilo.
- **Avisos de estructura (4.1).** Los avisos de «un solo subapartado» y de salto de nivel usan el nivel que da el clasificador; si se equivoca en un título falso, el aviso también. Los rótulos de referencias y de anexos cierran los apartados como un nivel 1.

- **Punto final de los niveles 4 y 5 (4.2 → 4.5).** Decisión del propietario: en el 4.2 solo se avisa («a) Evolución demográfica» y «b) Percepción de los empresarios» del 05). El 4.5, al unir el título con su párrafo, debe añadir el punto, con una acción nueva en `CambioPrevisto` («añadir al final») y su comprobación estricta, y quitar el aviso.
- **Cursiva de un estilo de carácter en un título (4.2).** La regla quita la cursiva puesta a mano en todo un título, pero no la que venga de un estilo de carácter (por ejemplo, «Énfasis») aplicado a todo el título. No pasa en los 5 documentos; si pasara, el verificador lo marcaría como incumplimiento («va en cursiva»).

- **Mensajes con opciones de la línea de comandos (4.3 → 9.5).** Los avisos de títulos numerados y en mayúsculas dicen «usa la opción --quitar-numeracion» / «--quitar-mayusculas». En la web, esas opciones serán casillas del formulario: el mensaje tendrá que nombrarlas como allí (o el informe tendrá que saber desde dónde se usa la herramienta).
- **Mayúsculas y nombres propios (4.3).** Al quitar las mayúsculas, una palabra se escribe como aparece a mitad de frase en el resto del trabajo («España», «UE»); una que no salga en el texto queda en minúscula, aunque sea un nombre propio. Por eso el informe pide revisar los títulos cambiados. Los números romanos («SIGLO XXI») se conservan, salvo «MI», «DI», «VI» y «LI», que son también palabras.

## Registro de avances

_(Una línea por paso terminado: fecha · paso · qué se hizo · commit.)_

- 2026-10-16 · 0.1 · `git init`, `.gitignore` (salida/, .venv/, \_\_pycache\_\_/, archivos temporales de Word) y primer commit · 02a823f
- 2026-10-16 · 0.2 · `requirements.txt` con versiones exactas (python-docx 1.2.0, lxml 6.1.3, pytest 9.1.1), `README.md` con instalación y uso, `.pytest_cache/` en `.gitignore`. Instalación probada desde cero en un entorno nuevo · 0bffc16
- 2026-10-16 · 0.3 · Estructura de carpetas creada: `apa_facil/` (con `reglas/` y `revision/`), `herramientas/` y `tests/`, con archivos mínimos que explican su función. Las reglas sin programar dan error si se usan. `pytest.ini` y primera prueba (18 comprobaciones en verde) · 6407ffc
- 2026-10-16 · 0.4 · `python -m apa_facil archivo.docx` abre el documento y guarda una copia sin cambios en `salida/nombre_APA.docx` (`documento.py`: abrir, ruta_salida, guardar; `pipeline.procesar`). Errores explicados en español (archivo inexistente, .doc, no Word, dañado, temporal de Word, copia abierta en Word). Los 5 documentos hacen la ida y vuelta y Word abre las copias con las mismas páginas que el original. 32 pruebas en verde · ff9b588
- 2026-10-16 · 0.5 · `documento.radiografia` (texto en orden de cuerpo, tablas y notas, y recuento de tablas, imágenes, notas y secciones) y `documento.comparar` (diferencias en español con su ubicación). `python -m apa_facil` verifica cada copia y no la entrega si el contenido cambia. 51 pruebas en verde · 3303625
- 2026-10-16 · 0.6 · `herramientas/inspeccionar.py` muestra márgenes por sección, fuentes (% de letras), interlineado, espaciado, sangrías, alineación, estilos, y encabezados y pies con los campos en palabras. Formato efectivo (valores por defecto → estilos → formato directo, fuentes del tema incluidas) en `documento.Estilos`, para reutilizarlo en la fase 1. 67 pruebas en verde. **Fase 0 terminada** · 2f1fe47
- 2026-10-16 · 1.1 · `reglas/margenes.py`: 2,54 cm en los cuatro lados de cada sección y encuadernación a 0, sin tocar tamaño, orientación ni distancia del encabezado. Avisa (sin cambiar tamaños) de tablas, imágenes y gráficos más anchos que el nuevo espacio para el texto. Primer `Informe` (cambios, avisos, errores); `pipeline.procesar` aplica las reglas y devuelve la copia y el informe, y la línea de comandos lo muestra (provisional hasta 1.7). Los 5 documentos quedan a 2,54 cm (la sección horizontal del 04 sigue horizontal), Word los abre con las mismas páginas que el original y nada se sale de los márgenes. 86 pruebas en verde · 1902040
- 2026-10-16 · 1.2 · `reglas/fuente.py`: una sola fuente por documento. Se mantiene la principal si es recomendada por APA (Aptos 12, Calibri 11, Arial 11, Lucida Sans Unicode 10, Times New Roman 12, Georgia 11, Computer Modern 10; con sus variantes de familia), y si no se cambia por la más parecida (tabla de equivalencias; las desconocidas, con aviso). Se aplica en valores por defecto, estilos (incluidos los de tabla), numeración de listas, cuerpo, tablas, cuadros de texto, encabezados, pies y notas, quitando las fuentes del tema. Respeta símbolos, ecuaciones, el tamaño de las notas al pie y el código de programación (estilo de código, o fuente monoespaciada que además parece código; se deja a 10 pt y se avisa). Utilidades de fuente y edición de partes (también las notas, que python-docx no carga como XML) en `documento.py`. Resultado: 01 Arial 11, 02 y 04 Calibri 11, 03 Georgia 11, 05 Times New Roman 12; Word lo confirma carácter a carácter. 135 pruebas en verde · ee1c87a
- 2026-10-16 · 1.3 · `reglas/interlineado.py`: doble espacio y 0 pt antes y después en valores por defecto, estilos y párrafos del cuerpo (portada, títulos, listas e índice incluidos), quitando el interlineado exacto o mínimo, las medidas en líneas y el espaciado automático. Tablas: se mantiene el sencillo, 1,5 o doble de cada párrafo y lo demás pasa a sencillo, sin espacio entre párrafos. Notas, encabezados, pies y cuadros de texto conservan su espaciado (se fija en sus párrafos y estilos para que no hereden el doble). Si hay índice, se definen los estilos "toc 1" a "toc 9", porque Word lo rehace al abrir el documento. Resultado: cuerpo a doble en los 5; tablas del 02 de 1,15 a sencillo, del 04 sencillas y del 05 a 1,5; notas, encabezados y pies sin cambios; Word lo confirma párrafo a párrafo. 160 pruebas en verde · 2f8b601
- 2026-10-16 · 1.4 · `reglas/alineacion.py`: el justificado (todas sus variantes) pasa a la izquierda en valores por defecto, estilos, cuerpo, tablas y notas; lo centrado no se toca; lo alineado a la derecha pasa a la izquierda salvo en lo que parece una portada (antes del primer salto o del primer párrafo de 40 palabras o más), que se deja con un aviso. Se desactivan los guiones automáticos y se avisa de los guiones opcionales sin quitarlos. Encabezados y pies conservan su alineación. Resultado: 35 párrafos justificados pasan a la izquierda (01, 03, 04 y 05), la frase alineada a la derecha del 04 también, y en el 05 se respetan las 3 líneas de la portada con aviso; Word lo confirma. 179 pruebas en verde · 6a08dbb
- 2026-10-16 · 1.5 · `reglas/sangria.py`: 1,27 cm en la primera línea de los párrafos de texto, puesta en cada párrafo; la sangría de primera línea de los estilos se quita para que títulos, listas y tablas no la hereden. Sin sangría (reglas sencillas, hasta el clasificador de la fase 3): vacíos, portada, centrados, títulos de Word y falsos, referencias, leyendas y notas, listas de Word y a mano, citas en bloque, primer párrafo del resumen y párrafos que empiezan con tabulador o espacios (estos, con aviso). La portada del paso 1.4 acaba ahora en el primer párrafo de 25 palabras (antes 40). Resultado: 8, 8, 6, 3 y 13 párrafos con sangría en los documentos 01 a 05, aviso por los 3 párrafos con tabulador del 04; sangrías francesas de listas y referencias intactas; Word lo confirma. 211 pruebas en verde · bc6b702
- 2026-10-16 · 1.6 · `reglas/paginacion.py`: un único encabezado (el de la primera sección; las demás se enlazan a él) con el campo PAGE a la derecha, en la fuente del documento, a interlineado sencillo y sin sangría. Se desactivan "primera página diferente" y "pares e impares", y la numeración pasa a ser continua y arábiga (con aviso si había reinicios o romanos). Se elimina lo que hubiera en los encabezados (texto, imágenes, tablas) con aviso; en los pies se quitan los números de página (el párrafo entero si solo numeraba, como "Página X de Y"; si hay otro texto, se conserva y se avisa). Resultado: los 5 documentos muestran el número arriba a la derecha en todas las páginas, también en la portada y el índice del 05; se quitan los números del pie de 02, 03 y 05 y el texto del encabezado de 03 y 05, con aviso; Word confirma que cada página muestra su número. Las pruebas de 1.3 y 1.4 ya no comprueban encabezados y pies (los rehace este paso). 233 pruebas en verde · 5ec58f5
- 2026-10-16 · 1.7 · `python -m apa_facil archivo.docx [--fuente "Arial"]`: la opción `--fuente` admite las 7 fuentes recomendadas (sin distinguir mayúsculas y con sus variantes) y, si no es una de ellas, no procesa nada y lo explica. Resumen del informe por grupos con su número (cambios, "Revisa tú", "No se ha podido hacer"). Arreglada la salida a `NUL` o a un archivo (se cambia a UTF-8 siempre que la salida no lo use). Los comodines (`pruebas\*.docx`) se expanden también en cmd y PowerShell. Borrado el prototipo `apa.py`; README con el uso real. Los 5 documentos se procesan sin errores desde PowerShell y pasan la prueba de integridad. 248 pruebas en verde. **Fase 1 terminada** · 90e171e
- 2026-10-16 · Fase 1 · El propietario revisa las cinco copias en Word: las reglas 1.1 a 1.6 se cumplen a simple vista. **Fase 1 cerrada.**
- 2026-10-17 · 2.1 · `herramientas/verificar.py`: comprueba las reglas 1.1 a 1.6 leyendo el XML con el formato efectivo (valores por defecto → estilos → formato directo), sin usar Word. Márgenes y encuadernación por sección; una sola fuente con su tamaño en cuerpo, tablas, cuadros de texto, encabezados, pies y notas (en estas, sin tamaño); doble espacio y 0 pt (también espaciado en líneas y automático), y sencillo, 1,5 o doble en tablas; nada justificado ni a la derecha y sin guiones automáticos; 1,27 cm en los párrafos de texto y ninguna sangría de primera línea en los demás; número de página a la derecha en el encabezado que se ve en cada sección, sin primera página diferente, pares e impares, reinicios, romanos, otro texto en el encabezado ni números en los pies. Cada incumplimiento lleva sección, párrafo y extracto; las excepciones de las reglas (portada a la derecha, código, párrafos con tabulador, texto del pie) se listan aparte. Termina con código 1 si algo no cumple. Resultado: las 5 copias de `salida/` cumplen todo; los originales dan 64, 134, 73, 94 y 140 incumplimientos. README con el uso. 287 pruebas en verde (39 nuevas: fallos conocidos de los originales y un fallo introducido en la copia del 01 por cada comprobación) · ef1f778
- 2026-10-17 · 2.2 · `tests/test_cinco_documentos.py`: cada documento de `pruebas/` se procesa una vez en una carpeta temporal y se comprueba que no hay errores, que el contenido de la copia (leída desde el disco) es idéntico al del original, que el verificador no encuentra incumplimientos, que el original no cambia (huella SHA-256) y que las excepciones admitidas son exactamente las esperadas (04: párrafos 7, 8 y 61 con tabulador; 05: párrafos 8, 9 y 10 de la portada; los demás, ninguna). Probado a la contra: sin la regla de sangría fallan las 5 copias en el verificador; con un detector de tabuladores demasiado permisivo el verificador da las 5 por buenas, pero la lista de excepciones lo detecta en las 5. Un documento nuevo en `pruebas/` obliga a revisar y apuntar sus excepciones. 313 pruebas en verde · 5cdd6b0
- 2026-10-17 · 2.3 · `informe.py`: cada entrada guarda su tipo, su regla y su mensaje; el pipeline anota qué regla se aplica antes de llamarla, así que las reglas no cambian y `cambios`, `avisos` y `errores` siguen dando las listas de mensajes. `pipeline.procesar` guarda `salida/nombre_APA.json` junto a cada copia entregada (versión del formato, archivo, copia, fecha, opciones, resumen y las entradas con regla y nombre legible; UTF-8 con tildes) y lo borra si la copia no se entrega. El texto de la terminal sale de `Informe.como_texto()`, con los grupos "Lo que hemos cambiado", "Lo que debes revisar tú" y "Lo que no se ha podido hacer" (pruebas de la línea de comandos y de márgenes adaptadas a los títulos nuevos), y la línea de comandos indica dónde está el JSON. Los 5 documentos generan su informe y siguen cumpliendo el verificador. 328 pruebas en verde · f738bdc
- 2026-10-17 · 2.4 · `herramientas/a_pdf.py`: convierte documentos a PDF en `salida/` (nunca en `pruebas/`) con Word si está instalado o, si no, con LibreOffice; sin ninguno, explica qué instalar. Word se maneja desde PowerShell con llamadas dinámicas, en solo lectura, sin avisos (no se queda esperando ante un cuadro de diálogo) y siempre se cierra; si se cuelga, solo se cierran los Word abiertos por la conversión, nunca el del usuario. Las páginas se cuentan en el propio PDF. Resultado: los 5 originales y las 5 copias se convierten en unos 12 s (copias: 2, 2, 2, 4 y 8 páginas) y no queda ningún Word abierto. Revisión de los PDF: las reglas de la fase 1 se ven bien en los cinco; lo pendiente es de fases posteriores (colores, títulos, leyendas, referencias) salvo un título que queda solo a final de página en el 03 (anotado para 4.1). 343 pruebas en verde (15 nuevas: conversores simulados y una conversión real con Word, que se omite si no hay Word). **Fase 2 terminada** · 68d2daa
- 2026-10-17 · 3.1 · `estructura.clasificar(doc)`: cada párrafo del cuerpo (sin tablas ni cuadros de texto) recibe tipo, nivel (en los títulos), confianza, zona y un motivo en español. Tres pasos: rasgos de cada párrafo (estilo y nivel de esquema de Word, listas, negrita, mayúsculas, tamaño, sangrías, imágenes, saltos, campo o control de índice, y si tiene una tabla o una imagen al lado), portada (hasta el primer salto, título de Word, rótulo o párrafo largo, sin el título que va justo antes del texto) y recorrido en orden con las zonas. Tipo nuevo `imagen` y dato `zona` (decisiones del propietario). Corrige lo que el detector sencillo hacía mal: «Introducción» del 01 y «¿Qué son las redes sociales?» del 04 ya no son portada, «Resumen» del 02 tampoco, las imágenes del 03 no son vacíos, el índice del 05 se reconoce y «Datos de uso por red social» del 04 es la leyenda de su tabla. Clasificación correcta de los 169 párrafos escrita a mano en `tests/clasificacion_esperada.py`: aciertan los 169 tipos y todas las zonas; un nivel queda pendiente para 3.2. Probado a la contra (sin detectar imágenes o sin recortar la portada, la comparación falla). Aún no lo usa ninguna regla (paso 3.4): las copias de salida no cambian. 405 pruebas en verde (62 nuevas) · 636b2b7
- 2026-10-17 · 3.2 · Títulos falsos por indicios: fuertes (negrita, mayúsculas, letra más grande que el texto normal —el de los párrafos largos fuera de la portada—, numeración de apartado seguida de texto o de su subapartado) y débiles (sin punto final, cursiva, centrado, seguido de un párrafo de texto); confianza alta con 3 indicios o más, media con menos y baja si solo hay indicios débiles o acaba en punto. No son títulos: enlaces, viñetas y líneas numeradas seguidas que empiezan igual (listas, como el anexo del 05). Nivel: por la numeración («1.1.» nivel 2, «I.» nivel 1 y entonces «1.» nivel 2, «a)» un nivel por debajo del último título sin letra) o por el formato (los títulos de Word y los numerados dicen qué formato tiene cada nivel; los demás formatos se ordenan por tamaño, centrado, mayúsculas, negrita y cursiva); un título falso no salta niveles respecto al anterior (se corrige con confianza baja). El motivo explica los indicios y de dónde sale el nivel. Resultado: los 169 párrafos coinciden con la clasificación esperada, también «Factores de riesgo:» del 04 (nivel 2); los títulos falsos del 01 y del 04 pasan a confianza alta. Probado a la contra (sin nivel por formato falla el 04). Las copias de salida no cambian. 437 pruebas en verde (33 nuevas; la de pendientes se omite porque ya no queda ninguno) · cc9d4e7
- 2026-10-17 · 3.3 · `python -m apa_facil archivo.docx --diagnostico`: no genera copia ni informe; muestra por documento una tabla con número de párrafo (el del clasificador y de `clasificacion_esperada.py`), tipo en español con su nivel, confianza y extracto recortado al ancho de la terminal (100 si no se sabe), una línea separadora por zona, el motivo debajo de los párrafos de confianza media o baja (decisión del propietario) y los recuentos de confianza y de tipos. Admite varios archivos y comodines; avisa si se combina con `--fuente`; errores de apertura en español. La tabla sale de `estructura.como_texto()`. README con el uso. Resultado con los 5 documentos: 169 párrafos (143 de confianza alta, 26 media, 0 baja) y los 169 tipos coinciden con la clasificación esperada (criterio de la fase 3: 95 %); `salida/` no cambia. 455 pruebas en verde (18 nuevas) · 18f0499
- 2026-10-17 · 3.4 · `pipeline.procesar` clasifica el documento una vez, antes de las reglas, y pasa la clasificación a cada una (`aplicar(doc, opciones, informe, estructura=None)`; convención actualizada). `reglas/sangria.py` rehecha sobre el clasificador: llevan 1,27 cm solo los párrafos `cuerpo`, salvo el primero del resumen, los centrados o a la derecha y los que empiezan con tabulador o espacios (con aviso); los motivos son los nombres de tipo del clasificador; borrados sus detectores propios. Un párrafo con sangría francesa fuera de las referencias se clasifica como referencia (baja si no empieza por autor y año) y no se toca, como antes. Numeración única: `Parrafo.numero` (desde 1 y contando las tablas, como el verificador), que muestra el diagnóstico; comentarios de `clasificacion_esperada.py` renumerados (datos sin cambios). El verificador usa el clasificador sobre la copia. Resultado: mismas copias que antes salvo «Más información en: www…» del 04, que es texto y ahora lleva sangría (el detector antiguo lo tomaba por título): 8, 8, 6, 4 y 13 párrafos con sangría; los 5 cumplen el verificador con las mismas excepciones. Probado a la contra (clasificando después de la regla de fuente, falla la prueba de la estructura del original). 465 pruebas en verde (10 nuevas) · f482358
- 2026-10-18 · 3.5 · `reglas/limpieza.py`, la primera regla del pipeline, con la estructura del original: borra los párrafos `vacio` de relleno (fuera de la portada, sin saltos, imágenes, marcadores, notas ni campos, y no junto a una tabla o imagen) y los quita de la estructura; quita los tabuladores y espacios iniciales de los párrafos `cuerpo` (también repartidos en varios fragmentos), que luego reciben la sangría; quita color, resaltado, fondo y subrayado del formato directo, salvo en enlaces (hipervínculos, estilo «Hipervínculo» o URL escrita a mano), portada, tablas, cuadros de texto y código. Un mensaje por tipo de limpieza en el informe («Limpieza de formato»). Integridad: `documento.CambioPrevisto` («borrar» o «recortar»), anotado con `informe.cambio_previsto` y aceptado por `comparar` solo si es exactamente eso. Resultado: 04, 6 vacíos borrados, 3 párrafos sin tabulador (ahora con sangría: 7 en total), color quitado en 2 fragmentos y un fondo amarillo en otro; 01, subrayado de «Introducción»; 02, color de las dos leyendas; 03 y 05 sin cambios. Los 5 cumplen el verificador (el 04 ya sin excepciones) y el PDF del 04 sigue con 4 páginas. Probado a la contra (sin mirar saltos y marcadores, sin enlaces, sin recortar y sin anotar los cambios previstos, fallan las pruebas). 495 pruebas en verde (30 nuevas) · bdc13db
- 2026-10-18 · 3.6 · `reglas/avisos.py` («Revisión del texto»), la primera regla del pipeline (antes de que la regla de fuente lo iguale todo), solo avisa y no cambia nada; no mira la portada ni las tablas. Un aviso por tipo con cuántos casos hay y hasta 2 ejemplos (sin números de párrafo, que cambian en la copia): saltos de línea manuales (no los de página o columna), viñetas escritas a mano (y si se mezclan símbolos; las listas numeradas a mano no, decisión del propietario), dos o más espacios entre palabras (no al principio ni tras una viñeta), y párrafos de texto, listas, citas o referencias con 20 letras o más en una fuente distinta de la principal (las variantes de una familia cuentan como la misma; no títulos, leyendas ni código). Resultado: el 04 recibe los 4 avisos (1 párrafo con saltos, 4 viñetas con «•» y «-», 1 doble espacio, 7 párrafos en otra fuente); el 02, el de otra fuente (un párrafo en Times New Roman); 01, 03 y 05, ninguno. Los 5 cumplen el verificador. Probado a la contra (con la regla detrás de la de fuente desaparece el aviso; contando saltos de página, distinguiendo variantes o sin umbral de letras, fallan las pruebas). 512 pruebas en verde (17 nuevas). **Fase 3 terminada:** el clasificador acierta los 169 tipos (100 %, criterio: 95 %) y los avisos de 3.6 están en el informe del 04 · 4ae8ba5
- 2026-10-18 · 4.1 · `reglas/titulos.py` («Títulos»), detrás de la sangría (que quita la sangría de primera línea de todos los estilos) y antes de la paginación. Busca los estilos Título 1 a 5 por su nombre interno («heading N», aunque el identificador sea «Ttulo1»), deja una sola definición si están duplicados (en el 02, 03 y 05 había dos de cada uno, una azul y otra en Calibri Light, Georgia o con mayúsculas) y crea, basados en Normal, los que falten. Formato APA de cada nivel en el estilo (negrita; centrado el 1; cursiva el 3 y el 5; sangría de 1,27 cm el 4 y el 5), sin color, subrayado, mayúsculas forzadas, bordes ni sombreado (también en el estilo de carácter vinculado), con «Conservar con el siguiente» y su nivel de esquema. Cada título del clasificador recibe el estilo de su nivel, salvo los de confianza baja, que solo se avisan (decisión del propietario: recomendación B); la numeración automática de un estilo, si la hubiera, solo se avisa. Avisos de apartado con un solo subapartado y de salto de nivel. El texto no cambia. Verificador: nueva comprobación «4.1 Estilos de título» (formato de cada estilo, duplicados y títulos sin su estilo; los dudosos, como excepción), y la sangría de los títulos 4 y 5 ya no es incumplimiento; el resultado dice ahora «cumple todas las reglas comprobadas». Resultado: 01, 5 títulos falsos pasan a Título 1; 04, 4 (3 de nivel 1 y «Factores de riesgo:» de nivel 2); 05, «RESUMEN»; el 02 y el 03 ya los usaban. Avisos de un solo subapartado en el 03 («2. Datos y resultados»), el 04 («Efectos sobre la autoestima») y el 05 («1. Introducción»). La copia del 04 se clasifica ya igual que el original. Word confirma en las 5 copias el estilo, el nivel de esquema (panel de navegación), la negrita, la cursiva, la alineación, el color automático y «Conservar con el siguiente» de cada título, y que ninguno queda solo al final de una página; los PDF tienen las mismas páginas (2, 2, 2, 4 y 8). Los originales dan ahora también incumplimientos de títulos. Probado a la contra (sin quitar duplicados, sin «Conservar con el siguiente», con la regla antes de la sangría, poniendo estilo a los dudosos o sin asignar estilos, fallan las pruebas). Pruebas de 1.4 y 1.5 adaptadas (lo único que se centra son los títulos de nivel 1; los títulos 4 y 5 llevan sangría). 544 pruebas en verde (32 nuevas) · a6cc22e
- 2026-10-18 · 4.2 · `reglas/titulos.py` quita del párrafo de cada título (con su estilo ya puesto) el formato puesto a mano que contradice su nivel, y solo ese: una alineación o una sangría distintas, bordes, sombreado, «Conservar con el siguiente» desactivado, la negrita desactivada, «Todo mayúsculas» y versalitas (formato: el texto no cambia) y la cursiva cuando va en todo el título al revés que su nivel; en una parte del título (un libro, una especie) se respeta. El color, el subrayado y el resaltado ya los quita la limpieza (3.5). Un cambio en el informe con qué se ha quitado y dónde (nada si no había nada). Avisos nuevos: numeración automática de Word en un título (se deja, paso 4.3) y títulos de nivel 4 o 5 sin punto final (decisión del propietario: el punto se añade en el 4.5). Verificador: comprobación «4.2 Formato de los títulos», párrafo a párrafo con el formato final (alineación, negrita en todo el título, cursiva del nivel salvo la parcial, sangría izquierda y francesa en todos y de primera línea en los niveles 4 y 5 —la de los niveles 1 a 3 ya la mira 1.5—, «Conservar con el siguiente», sin mayúsculas forzadas, color ni subrayado salvo en enlaces). Resultado: en los 5 documentos no quedaba formato a mano que contradijera el nivel (los títulos falsos solo conservan su negrita, que coincide), así que las copias no cambian; aviso del punto final en el 05 («a)» y «b)», nivel 4). Los originales dan incumplimientos de 4.2 en el 02, 03 y 05 (el «2.1.» del 03 iba en cursiva por su estilo duplicado). Word confirma los 38 títulos de las copias: los 26 de nivel 1, centrados; los 8 de nivel 2 y los 2 de nivel 3, a la izquierda (los de nivel 3, en cursiva); los 2 de nivel 4, con 36 pt (1,27 cm) de sangría; todos en negrita, sin color ni mayúsculas y con «Conservar con el siguiente»; los PDF, con las mismas páginas. Probado a la contra (quitando también la cursiva parcial, sin quitar el formato a mano, sin quitar la alineación o sin el aviso del punto, fallan las pruebas). 554 pruebas en verde (10 nuevas) · 9b75cd2
- 2026-10-18 · 4.3 · Títulos numerados y en mayúsculas (`reglas/titulos.py`). Por defecto solo se avisa, un aviso por tipo con ejemplos (decisión del propietario); con `--quitar-numeracion` se quita la numeración de apartado («1.», «1.1», «2.3.1.», «IV.», «a)», con el espacio que la sigue; un número sin punto ni paréntesis, como «20 años de…», no), y con `--quitar-mayusculas` los títulos escritos enteros en mayúsculas pasan a mayúscula inicial, con los nombres propios y las siglas como los escribe el resto del trabajo y los números romanos como están. Las dos opciones son independientes. El texto se cambia dentro de cada fragmento, sin mover el formato. Integridad: `CambioPrevisto` con dos acciones nuevas, cada una con su comprobación estricta: `quitar_numeracion` (lo quitado tiene que ser una numeración de apartado al principio y dejar texto detrás) y `cambiar_mayusculas` (el texto nuevo solo puede diferir en mayúsculas y minúsculas de uno escrito en mayúsculas). Avisos al cambiar: remisiones a apartados por su número («apartado 2.1») y un índice que Word no actualiza solo. Informe JSON, versión 2: las opciones de los títulos. Verificador: «4.3 Numeración y mayúsculas de los títulos», que no da incumplimientos sino excepciones admitidas (01: 3; 03: 5; 05: 12 numerados y «RESUMEN»). Resultado sin opciones: avisos en el 01, 03 y 05, y las copias no cambian. Con las dos opciones: 3, 5 y 13 títulos cambiados, los 5 documentos pasan la integridad y el verificador sin excepciones de títulos; Word confirma los títulos nuevos con su estilo, nivel de esquema y formato, y al actualizar el índice del 05 muestra los títulos sin número y «Resumen». Probado a la contra (aceptando quitar cualquier cosa o cualquier texto nuevo en la integridad, sin anotar el cambio, quitando números sin punto, sin buscar los nombres propios o quitando la numeración por defecto, fallan las pruebas). 578 pruebas en verde (24 nuevas) · e903440
