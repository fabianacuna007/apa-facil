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
├── pruebas/                   # documentos de prueba: NO MODIFICAR
└── salida/                    # documentos generados
```

### Convenciones

- Cada regla es una función `aplicar(doc, opciones, informe)`. Recibe el documento, las opciones del usuario y el informe donde anota lo que hizo.
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
- [ ] **0.4 Ida y vuelta.** Línea de comandos que abre un .docx y lo guarda en `salida/` **sin cambios**.
- [ ] **0.5 Prueba de integridad.** Una función que extrae todo el texto de un documento (párrafos, tablas y notas al pie), en orden. Una prueba compara el original con la copia del paso 0.4: el texto debe ser idéntico, y el número de tablas, imágenes, notas al pie y secciones también.
- [ ] **0.6 Inspector.** `herramientas/inspeccionar.py`. Muestra, por documento:
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

- [ ] **1.1 Márgenes.** 2,54 cm en los cuatro lados, en **todas** las secciones. Las páginas en horizontal conservan su orientación.
- [ ] **1.2 Fuente.**
  - Aplicarla en los estilos (Normal y los demás) **y** en el formato directo de cada fragmento de texto, incluidos los atributos de fuente para otros alfabetos (`ascii`, `hAnsi`, `eastAsia`, `cs`).
  - Fuente elegible por opción: Times New Roman 12 (por defecto), Calibri 11, Arial 11, Georgia 11, Lucida Sans Unicode 10 o Computer Modern 10.
  - Tablas: misma fuente (el tamaño puede ser algo menor).
- [ ] **1.3 Interlineado.** Doble, con 0 pt de espacio antes y después del párrafo. Excepciones:
  - dentro de tablas: sencillo;
  - notas al pie: formato por defecto de Word.
- [ ] **1.4 Alineación.** Izquierda, sin justificar. Las tablas y las imágenes centradas se quedan como están.
- [ ] **1.5 Sangría (versión simple).** 1,27 cm en la primera línea de los párrafos de cuerpo de texto, no en títulos, listas, tablas ni párrafos centrados. En la fase 3 se refina con el detector de estructura.
- [ ] **1.6 Número de página.**
  - Campo PAGE alineado a la derecha en el encabezado de todas las secciones, incluida la primera página.
  - Hay que tener en cuenta la opción "primera página diferente" y los encabezados de páginas pares.
  - Quitar los números de página que ya existan en pies y encabezados.
  - Si había otro texto en el encabezado, quitarlo y registrar un aviso, porque APA de estudiante no lleva encabezado de página.
- [ ] **1.7 Línea de comandos completa.** `python -m apa_facil entrada.docx --fuente "Arial"` genera `salida/entrada_APA.docx` y muestra un resumen del informe.

**Hecho cuando:** los 5 documentos se procesan sin errores, pasan la prueba de integridad, y al abrirlos en Word las reglas 1.1 a 1.6 se cumplen a simple vista.

---

## Fase 2: Verificación e informe

**Por qué:** sin un verificador automático, cada cambio obliga a abrir los Word a mano para ver si algo se ha roto. Con él, las regresiones se detectan solas.

- [ ] **2.1 Verificador APA.** `herramientas/verificar.py` lee un .docx y comprueba cada regla de la fase 1. Devuelve la lista de incumplimientos con su ubicación (sección, número de párrafo y un extracto del texto).
- [ ] **2.2 Pruebas automáticas.** `pytest` que, para cada documento de `pruebas/`, hace tres cosas:
  1. lo procesa;
  2. comprueba la integridad del texto y de los elementos;
  3. pasa el verificador sin incumplimientos.
- [ ] **2.3 Informe estructurado.** Informe en JSON (para la web futura) y en texto legible en la terminal. Ejemplo de aviso: *"El encabezado contenía el texto 'Facultad de Ciencias…'. Se ha eliminado porque APA de estudiante no lo usa."*
- [ ] **2.4 Revisión visual (opcional).** Si LibreOffice está instalado, convertir la salida a PDF para revisarla sin abrir Word.

**Hecho cuando:** `pytest` pasa en verde con los 5 documentos y cada procesado genera su informe.

---

## Fase 3: Detección de estructura y limpieza

**Por qué:** es la fase más difícil y la que diferencia una herramienta buena de una mala. Para aplicar bien la sangría, los títulos o las referencias, primero hay que saber qué es cada párrafo, aunque el estudiante no haya usado los estilos de Word.

- [ ] **3.1 Clasificador de párrafos.** `estructura.py` asigna a cada párrafo un tipo y una confianza (alta, media o baja). Tipos:
  - `portada`, `indice`, `titulo` (con su nivel), `cuerpo`, `lista`;
  - `leyenda_tabla`, `leyenda_figura`, `cita_bloque`;
  - `titulo_referencias`, `referencia`, `anexo`, `vacio`.
- [ ] **3.2 Títulos falsos.** Detectar como título un párrafo corto, en negrita o en mayúsculas, sin punto final, a veces con numeración ("1.", "1.1."), aunque use el estilo Normal. Deducir su nivel por la numeración, el tamaño y el formato.
- [ ] **3.3 Modo diagnóstico.** La opción `--diagnostico` imprime una tabla con párrafo, tipo, confianza y extracto. Sirve para depurar el clasificador con los documentos de prueba.
- [ ] **3.4 Sangría definitiva.** Rehacer la regla 1.5 usando el clasificador: solo los párrafos `cuerpo` llevan sangría de primera línea.
- [ ] **3.5 Limpieza de formato.** Todo lo que se haga aquí queda registrado en el informe.
  - Eliminar los párrafos vacíos usados como espaciado.
  - Eliminar los tabuladores y espacios al inicio de párrafo usados como sangría.
  - Quitar colores de texto, resaltados y subrayados decorativos. Los hipervínculos se mantienen.
- [ ] **3.6 Problemas que solo se avisan, sin corregirlos.**
  - Saltos de línea manuales dentro de un párrafo.
  - Viñetas escritas a mano ("•", "-").
  - Dobles espacios entre palabras.
  - Texto en otra fuente que parezca pegado de internet.

**Hecho cuando:** el clasificador acierta el tipo de al menos el 95 % de los párrafos de los 5 documentos (revisado con el modo diagnóstico), y los avisos de 3.6 aparecen en el informe del documento 04.

---

## Fase 4: Títulos APA (cinco niveles)

**Por qué:** los títulos son lo más visible después de la fuente, y dependen del clasificador de la fase 3.

- [ ] **4.1 Estilos de título.** Modificar los estilos de título de Word (Título 1 a Título 5) con el formato APA y asignar cada título detectado a su estilo. Usar estilos en vez de formato directo mantiene el panel de navegación y el índice de Word.
- [ ] **4.2 Formato por nivel:**
  - Nivel 1: centrado, negrita.
  - Nivel 2: a la izquierda, negrita.
  - Nivel 3: a la izquierda, negrita y cursiva.
  - Nivel 4: con sangría, negrita, terminado en punto, y el texto sigue en la misma línea.
  - Nivel 5: con sangría, negrita y cursiva, terminado en punto, y el texto sigue en la misma línea.
- [ ] **4.3 Títulos en mayúsculas y numeración.** Cambiar un texto de MAYÚSCULAS a mayúscula inicial o quitar la numeración ("1. Introducción" → "Introducción") **modifica el texto**. Por eso:
  - por defecto, solo se avisa;
  - con una opción explícita del usuario, se cambia y se registra.
- [ ] **4.4 Título del trabajo.** En la primera página del cuerpo va el título del trabajo, centrado y en negrita, como nivel 1. El primer apartado no lleva el rótulo "Introducción".
- [ ] **4.5 Niveles 4 y 5 en línea.** Unir el título con el párrafo siguiente es delicado. Implementarlo al final de la fase, con una prueba específica.

**Hecho cuando:** los títulos de los documentos 02, 03 y 05 quedan con el formato de su nivel y el panel de navegación de Word sigue funcionando.

---

## Fase 5: Portada de estudiante y secciones iniciales

- [ ] **5.1 Datos de portada.** Los datos llegan por opciones o por un archivo `datos.json`: título, autor o autores, departamento y universidad, código y nombre de la asignatura, profesor y fecha de entrega.
- [ ] **5.2 Generar la portada APA.**
  - El título, centrado y en negrita, se coloca 3 o 4 líneas por debajo del margen superior.
  - Después, una línea en blanco.
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
- [ ] **6.2 Figuras.** El mismo esquema: "Figura N" en negrita y el título en cursiva encima de la imagen.
- [ ] **6.3 Citas en bloque** (40 palabras o más).
  - Sangría izquierda de 1,27 cm en todo el bloque, sin sangría de primera línea y con doble espacio.
  - Quitar la cursiva decorativa.
  - Si la cita lleva comillas, solo se avisa: quitarlas modificaría el texto.
- [ ] **6.4 Lista de referencias.**
  - Detectar la sección aunque se llame "Bibliografía", "Webgrafía" o "Fuentes".
  - Empieza en página nueva, con el rótulo "Referencias" centrado y en negrita.
  - Sangría francesa de 1,27 cm y doble espacio.
- [ ] **6.5 Orden alfabético.** Reordenar las referencias es una opción activada por defecto y queda registrada. Si hay varias listas (por ejemplo, una bibliografía y una webgrafía), se avisa y se propone unirlas.
- [ ] **6.6 Notas al pie.** Misma fuente que el documento; el resto del formato no se toca.

**Hecho cuando:** las tablas de los documentos 02, 04 y 05, las figuras del 03, la cita en bloque del 05 y las referencias de los cinco documentos quedan en formato APA, y la prueba de integridad sigue pasando (permitiendo el cambio de rótulo y el nuevo orden).

---

## Fase 7: Revisión de citas y referencias (solo análisis)

**Por qué:** es la segunda función del producto. Aquí la herramienta **no modifica nada**: solo detecta y explica.

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

- [ ] Fuente por defecto: Times New Roman 12 o Calibri 11.
- [ ] Si la portada original se sustituye o se mantiene por defecto.
- [ ] Si el orden alfabético de las referencias se aplica automáticamente o solo se avisa.
- [ ] Si los títulos en mayúsculas o numerados se corrigen por defecto.
- [ ] Nombre de dominio.

---

## Especificación APA 7 de estudiante (referencia)

| Elemento | Regla |
|---|---|
| Márgenes | 2,54 cm en los cuatro lados |
| Fuentes admitidas | Times New Roman 12, Calibri 11, Arial 11, Georgia 11, Lucida Sans Unicode 10, Computer Modern 10 |
| Interlineado | Doble en todo el documento, incluidas la portada y las referencias. Sin espacio extra antes ni después de los párrafos. |
| Excepciones de interlineado | Interior de tablas y figuras (sencillo, 1,5 o doble) y notas al pie |
| Alineación | Izquierda, con el borde derecho irregular. Sin justificar y sin guiones automáticos. |
| Sangría | 1,27 cm en la primera línea de los párrafos de cuerpo |
| Sin sangría | Portada, rótulos de sección, títulos de nivel 1 a 3, párrafo del resumen, títulos y notas de tablas y figuras |
| Número de página | Arriba a la derecha, en todas las páginas, incluida la portada |
| Encabezado de página | No se usa en trabajos de estudiante |
| Portada | Título en negrita centrado, 3 o 4 líneas por debajo del margen superior. Debajo, centrados: autor, afiliación, asignatura, profesor y fecha. |
| Título nivel 1 | Centrado, negrita |
| Título nivel 2 | A la izquierda, negrita |
| Título nivel 3 | A la izquierda, negrita y cursiva |
| Título nivel 4 | Con sangría, negrita, terminado en punto, el texto sigue en la misma línea |
| Título nivel 5 | Con sangría, negrita y cursiva, terminado en punto, el texto sigue en la misma línea |
| Tablas | "Tabla N" en negrita encima; título en cursiva en la línea siguiente; "*Nota.*" debajo |
| Figuras | "Figura N" en negrita encima; título en cursiva en la línea siguiente; "*Nota.*" debajo |
| Cita en bloque | 40 palabras o más. Sangría izquierda de 1,27 cm en todo el bloque, sin comillas; la cita va después del punto final. |
| Referencias | Página nueva; rótulo "Referencias" centrado y en negrita; orden alfabético; sangría francesa de 1,27 cm; doble espacio |
| APA en español | "y" entre autores (no "&"); "s. f." para obras sin fecha; "et al."; "p." y "pp."; DOI con formato `https://doi.org/...` |

---

## Problemas abiertos

_(Claude Code anota aquí lo que descubra y afecte a otros pasos.)_

- **Prototipo `apa.py` (anterior a este ROADMAP).** Ya aplica casi toda la fase 1 en un solo archivo, y tiene una comprobación de integridad (`foto`/`comparar`) que no cuenta las secciones. Su código se reutilizará al repartirlo en `apa_facil/` (pasos 0.4, 0.5 y 1.1–1.7). Se borra cuando 1.7 esté hecho. Diferencias con este ROADMAP que habrá que corregir al pasarlo:
  - 1.2: falta Computer Modern 10.
  - 1.3: pone interlineado doble dentro de las tablas (debe ser sencillo) y fuerza interlineado sencillo en las notas al pie (debe quedar el formato por defecto de Word).
  - 1.6: conserva el texto que había en el encabezado y deja los números de página que ya hubiera en él (debe quitar las dos cosas y avisar).
- **2.4 Revisión visual.** LibreOffice no está instalado en este equipo. Word sí, y se puede convertir a PDF por COM desde PowerShell, pero solo con llamadas dinámicas (`InvokeMember`): el acceso normal falla con el error `E_NOINTERFACE`.

## Registro de avances

_(Una línea por paso terminado: fecha · paso · qué se hizo · commit.)_

- 2026-10-16 · 0.1 · `git init`, `.gitignore` (salida/, .venv/, \_\_pycache\_\_/, archivos temporales de Word) y primer commit · 02a823f
- 2026-10-16 · 0.2 · `requirements.txt` con versiones exactas (python-docx 1.2.0, lxml 6.1.3, pytest 9.1.1), `README.md` con instalación y uso, `.pytest_cache/` en `.gitignore`. Instalación probada desde cero en un entorno nuevo · 0bffc16
- 2026-10-16 · 0.3 · Estructura de carpetas creada: `apa_facil/` (con `reglas/` y `revision/`), `herramientas/` y `tests/`, con archivos mínimos que explican su función. Las reglas sin programar dan error si se usan. `pytest.ini` y primera prueba (18 comprobaciones en verde) · PENDIENTE
