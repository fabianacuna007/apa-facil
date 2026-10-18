"""
Paso 5.2: la portada APA de estudiante, a partir de los datos del paso 5.1.

Según la web oficial de APA (Title page setup, Title case capitalization y
Order of pages):
- el título, 3 o 4 líneas por debajo del margen superior, centrado, en negrita
  y con mayúscula en las palabras principales;
- una línea en blanco a doble espacio;
- debajo, centrados y cada uno en su línea: autores, afiliación, asignatura,
  profesor y fecha (las líneas las compone datos_portada.parrafos_de_portada);
- todo a doble espacio, con el número de página (paso 1.6), y el texto empieza
  en una página nueva («start each section on a new page»).

La portada se genera solo si el documento no tiene ya una (decisión del
propietario: sustituirla o mantenerla es el paso 5.3). Sin título no se genera:
APA lo exige. El título del principio del texto (paso 4.4) sale entonces de los
datos, con las mismas mayúsculas que en la portada.

Mayúsculas del título (decisión del propietario: por defecto, con la opción
--titulo-tal-cual se deja como lo escribe el estudiante). APA llama palabras
menores, en minúscula, a todos los artículos y a las conjunciones y
preposiciones de tres letras o menos; las demás (sustantivos, verbos,
adjetivos, adverbios, pronombres y cualquier palabra de cuatro letras o más)
van con mayúscula, y también la primera del título, la del subtítulo y la que
sigue a dos puntos, raya o punto. Las palabras que pueden ser de las dos clases
según la frase («que», «lo», «vía») se dejan como están y se avisa.
"""

import re

from docx.oxml import OxmlElement

from apa_facil.datos_portada import LINEAS_EN_BLANCO_ANTES_DEL_TITULO, parrafos_de_portada
from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import clasificar
from apa_facil.reglas import titulos

# Palabras menores en español, con los criterios de APA.
ARTICULOS = {"el", "la", "los", "las", "un", "una", "unos", "unas"}
CONTRACCIONES = {"al", "del"}   # preposición y artículo
CONJUNCIONES = {"y", "e", "o", "u", "ni", "si", "mas"}
PREPOSICIONES = {"a", "de", "en", "con", "por", "sin", "so", "pro"}
MENORES = ARTICULOS | CONTRACCIONES | CONJUNCIONES | PREPOSICIONES
# Según la frase son menores (conjunción, artículo, preposición) o principales (pronombre, sustantivo).
DUDOSAS = {"que", "lo", "vía", "via"}

PALABRA = re.compile(r"[^\W\d_]+(?:['’][^\W\d_]+)*")   # «O'Donnell» es una sola palabra
# Lo que hace empezar un subtítulo o una frase: la palabra siguiente va con mayúscula.
ABRE_SUBTITULO = re.compile(r"[:.?!…—\n]|\s–|–\s")


def _con_mayuscula(palabra):
    return palabra[0].upper() + palabra[1:]


def _nombres_propios(clasificados):
    """
    Las parejas «Palabra Palabra» que el trabajo escribe a mitad de frase, para no
    pasar a minúscula el artículo de un nombre propio («La Rioja», «El Salvador»).
    """
    parejas = set()
    for c in clasificados:
        if c.tipo in ("titulo", "titulo_referencias", "anexo", "vacio", "imagen") or titulos.en_mayusculas(c.texto):
            continue
        for frase in re.split(r"[.!?¿¡:]\s+", c.texto):
            palabras = PALABRA.findall(frase)
            for anterior, siguiente in zip(palabras[1:], palabras[2:]):   # la primera de la frase, no
                if anterior.lower() in MENORES and anterior[0].isupper() and siguiente[0].isupper():
                    parejas.add((anterior, siguiente))
    return parejas


def mayusculas_apa(titulo, clasificados=()):
    """
    El título con las mayúsculas de APA («title case»). Devuelve (título, dudosas):
    las palabras que se han dejado como estaban porque su clase depende de la frase.

    - Una palabra principal en minúscula pasa a mayúscula inicial; una con mayúsculas
      dentro («iPhone») o una sigla («UE», «COVID») se dejan.
    - Una palabra menor con mayúscula inicial pasa a minúscula, salvo si va al
      principio, si es una sigla o si es parte de un nombre propio que el trabajo
      escribe así («La Rioja»).
    - Un título escrito entero en mayúsculas se pasa antes a minúsculas como en el
      paso 4.3, con los nombres propios y las siglas como los escribe el trabajo.
    """
    if titulos.en_mayusculas(titulo):
        titulo = titulos.a_minusculas(titulo, titulos._formas_del_texto(clasificados))
    mixto = any(c.islower() for c in titulo)
    propios = _nombres_propios(clasificados)
    piezas = list(PALABRA.finditer(titulo))
    resultado, dudosas, hasta, forzar = [], [], 0, True
    for numero, encontrada in enumerate(piezas):
        palabra = encontrada.group(0)
        resultado.append(titulo[hasta:encontrada.start()])
        minuscula = palabra.lower()
        sigla = mixto and len(palabra) >= 2 and palabra.isupper()
        siguiente = piezas[numero + 1].group(0) if numero + 1 < len(piezas) else ""
        if sigla:
            nueva = palabra
        elif forzar:
            nueva = _con_mayuscula(palabra) if palabra.islower() else palabra
        elif minuscula in DUDOSAS:
            nueva = palabra
            dudosas.append(palabra)
        elif minuscula in MENORES:
            nombre_propio = (_con_mayuscula(minuscula), _con_mayuscula(siguiente)) in propios
            nueva = _con_mayuscula(minuscula) if nombre_propio else (minuscula if palabra[1:].islower() or
                                                                    len(palabra) == 1 else palabra)
        else:
            nueva = _con_mayuscula(palabra) if palabra.islower() else palabra
        resultado.append(nueva)
        hasta = encontrada.end()
        entre = titulo[hasta:piezas[numero + 1].start()] if numero + 1 < len(piezas) else ""
        forzar = bool(ABRE_SUBTITULO.search(entre))
    resultado.append(titulo[hasta:])
    return "".join(resultado), dudosas


# --- Los párrafos ---------------------------------------------------------------------

def _parrafo(texto="", id_estilo=None, centrado=False, esquema=False):
    """Un párrafo nuevo con su texto (un «\\n» es un salto de línea). El resto del formato, del estilo."""
    p = OxmlElement("w:p")
    if id_estilo or centrado or esquema:
        ppr = p.get_or_add_pPr()
        if id_estilo:
            ppr.get_or_add_pStyle().set(w("val"), id_estilo)
        if esquema:
            ppr.get_or_add_outlineLvl().set(w("val"), "0")
        if centrado:
            ppr.get_or_add_jc().set(w("val"), "center")
    for numero, linea in enumerate(texto.split("\n") if texto else []):
        r = OxmlElement("w:r")
        if numero:
            r.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.text = linea
        if linea != linea.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        p.append(r)
    return p


def _id_del_estilo_del_titulo(doc):
    estilos = titulos.estilos_del_titulo_del_trabajo(doc)
    return estilos[0].get(w("styleId")) if estilos else None


def _empezar_pagina(bloque, ultimo_de_la_portada):
    """
    El texto empieza en una página nueva: «salto de página anterior» en su primer párrafo.
    Devuelve True si ha hecho falta un salto de página escrito (cuenta como un salto de línea en el texto).
    """
    if bloque is not None and bloque.tag == w("p"):
        ppr = bloque.get_or_add_pPr()
        salto = ppr.get_or_add_pageBreakBefore()
        salto.attrib.pop(w("val"), None)   # encendido (quizá estaba apagado a mano)
        return False
    # Una tabla al principio no admite ese ajuste: un salto de página al final de la portada.
    r = OxmlElement("w:r")
    salto = OxmlElement("w:br")
    salto.set(w("type"), "page")
    r.append(salto)
    ultimo_de_la_portada.append(r)
    return True


def _extracto(texto, largo=60):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _elementos(datos):
    """Qué lleva la portada, para el informe: «el título», «los autores»…"""
    return [nombre for nombre, dado in (
        ("el título", datos.titulo), ("los autores" if len(datos.autores) > 1 else "el autor", datos.autores),
        ("la afiliación", datos.departamento or datos.universidad),
        ("la asignatura", datos.codigo_asignatura or datos.asignatura),
        ("los profesores" if len(datos.profesores) > 1 else "el profesor", datos.profesores),
        ("la fecha de entrega", datos.fecha)) if dado]


def _lista(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


def aplicar(doc, opciones, informe, estructura=None):
    datos = opciones.get("portada")
    if not datos:
        return
    clasificados = estructura if estructura is not None else clasificar(doc)
    existente = [c for c in clasificados if c.zona == "portada" and c.texto]
    if existente:
        informe.aviso(f"Tu documento ya tiene una portada («{_extracto(existente[0].texto, 40)}»), así que de "
                      "momento se deja como está y no se ha añadido la portada APA con tus datos.")
        return
    if not datos.titulo:
        informe.aviso("No se ha añadido la portada APA porque falta el título del trabajo, que APA exige. "
                      "Añádelo en tu archivo de datos o con la opción --titulo.")
        return
    id_estilo = _id_del_estilo_del_titulo(doc)
    if id_estilo is None:   # lo crea el paso 4.4 cuando hay un título; aquí siempre debería estar
        informe.error("No se ha podido añadir la portada APA: falta el estilo «Título» de Word.")
        return

    if opciones.get("titulo_tal_cual"):
        titulo, dudosas = datos.titulo, []
    else:
        titulo, dudosas = mayusculas_apa(datos.titulo, clasificados)

    cuerpo = doc.element.body
    primero_del_documento = next((el for el in cuerpo if el.tag in (w("p"), w("tbl"), w("sdt"))), None)
    textos = parrafos_de_portada(datos, titulo)
    portada = [_parrafo(texto, id_estilo=id_estilo) if numero == LINEAS_EN_BLANCO_ANTES_DEL_TITULO
               else _parrafo(texto, centrado=bool(texto))
               for numero, texto in enumerate(textos)]
    for p in portada:
        if primero_del_documento is not None:
            primero_del_documento.addprevious(p)
        else:
            cuerpo.insert(len(cuerpo) - 1 if len(cuerpo) and cuerpo[-1].tag == w("sectPr") else len(cuerpo), p)

    # El título al principio del texto (paso 4.4), el mismo de la portada.
    mensaje_del_titulo = None
    primero = next((c for c in clasificados if c.zona == "cuerpo" and c.tipo != "vacio"
                    and c.elemento.getparent() is not None), None)
    if primero is not None and " ".join(texto_parrafo(primero.elemento).split()) != " ".join(titulo.split()):
        nuevo = _parrafo(titulo, id_estilo=id_estilo, esquema=True)
        if titulos._salto_de_pagina_antes(Estilos(doc), primero.elemento):
            nuevo.get_or_add_pPr().get_or_add_pageBreakBefore()
            primero.elemento.get_or_add_pPr().get_or_add_pageBreakBefore().set(w("val"), "0")
        primero.elemento.addprevious(nuevo)
        informe.cambio_previsto(primero.numero, "añadir_titulo", nuevo=titulo)
        mensaje_del_titulo = (f"Título del trabajo: se ha puesto «{_extracto(titulo)}» al principio del texto, "
                              "centrado y en negrita, como pide APA.")
    if _empezar_pagina(portada[-1].getnext(), portada[-1]):
        textos[-1] += "\n"
    informe.cambio_previsto(1, "añadir_portada", parrafos=tuple(textos), datos=datos)

    informe.cambio("Portada: se ha añadido la portada APA de estudiante con " + _lista(_elementos(datos)) +
                   ", centrados y a doble espacio, con el título en negrita tres líneas por debajo del margen "
                   "superior. El texto empieza en una página nueva.")
    if mensaje_del_titulo:
        informe.cambio(mensaje_del_titulo)
    if titulo != datos.titulo:
        informe.cambio(f"Portada: el título lleva ahora mayúscula inicial en las palabras principales, como pide "
                       f"APA: «{_extracto(datos.titulo)}» → «{_extracto(titulo)}». Revísalo, sobre todo los "
                       "nombres propios; si prefieres dejarlo como lo escribiste, usa la opción --titulo-tal-cual.")
    if dudosas:
        vistas = list(dict.fromkeys(dudosas))
        informe.aviso("En el título se ha" + ("n" if len(vistas) > 1 else "") + " dejado como " +
                      ("las" if len(vistas) > 1 else "la") + " escribiste " +
                      ", ".join(f"«{p}»" for p in vistas) + ": APA pone mayúscula a los pronombres y a los "
                      "sustantivos, pero no a los artículos, ni a las conjunciones y preposiciones cortas. "
                      "Revisa cuál es en tu título.")
