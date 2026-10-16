"""
Regla 1.6: número de página arriba a la derecha en todas las páginas,
también en la portada. En APA de estudiante el encabezado solo lleva eso.

- Un único encabezado para todo el documento: lo define la primera sección y
  las demás se enlazan a él. Lleva el campo PAGE de Word, alineado a la derecha.
- Se desactivan "primera página diferente" y "pares e impares diferentes",
  que harían que algunas páginas usaran otro encabezado sin número.
- La numeración es continua desde 1: se quitan los reinicios y los números
  romanos de las secciones (con aviso).
- El texto, las imágenes y las tablas que hubiera en los encabezados se
  eliminan y se avisa (lo prevé el ROADMAP).
- En los pies se quitan los números de página. Un párrafo que solo sirve para
  numerar ("- 3 -", "Página 3 de 9") se quita entero; si tiene otro texto, el
  texto se conserva y se avisa.
"""

import re

from docx.oxml import OxmlElement

from apa_facil.documento import poner_espaciado, texto_legible, w

CAMPOS_DE_PAGINA = ("PAGE", "NUMPAGES", "SECTIONPAGES")
MARCAS_DE_CAMPO = re.compile(r"\[(nº de página|total de páginas|páginas de la sección)\]")
SOLO_NUMERACION = re.compile(r"^[\s\-–—/|·.:,()\[\]]*((p[aá]g(ina)?s?|page|de|of|\d+)[\s\-–—/|·.:,()\[\]]*)*$",
                             re.IGNORECASE)


def _codigo(texto):
    partes = (texto or "").split()
    return partes[0].upper() if partes else ""


def tiene_numero_de_pagina(p):
    codigos = [_codigo(t.text) for t in p.iter(w("instrText"))]
    codigos += [_codigo(f.get(w("instr"))) for f in p.iter(w("fldSimple"))]
    return any(c in CAMPOS_DE_PAGINA for c in codigos)


def _quitar_campos_de_pagina(p):
    """Quita los campos de número de página de un párrafo, dejando el resto del texto."""
    for campo in list(p.iter(w("fldSimple"))):
        if _codigo(campo.get(w("instr"))) in CAMPOS_DE_PAGINA:
            campo.getparent().remove(campo)
    capturados, codigo, profundidad = [], "", 0
    for el in list(p.iter(w("fldChar"), w("instrText"), w("t"), w("tab"), w("br"))):
        etiqueta = el.tag
        tipo = el.get(w("fldCharType")) if etiqueta == w("fldChar") else None
        if tipo == "begin":
            profundidad += 1
            if profundidad == 1:
                capturados, codigo = [], ""
        if profundidad:
            capturados.append(el)
            if etiqueta == w("instrText"):
                codigo += el.text or ""
        if tipo == "end":
            profundidad -= 1
            if profundidad == 0 and _codigo(codigo) in CAMPOS_DE_PAGINA:
                for capturado in capturados:
                    capturado.getparent().remove(capturado)


def _texto_sin_campos(p):
    return MARCAS_DE_CAMPO.sub("", texto_legible(p)).strip()


def _partes(secciones, nombres):
    """Las partes de encabezado o pie distintas (varias secciones pueden compartir una)."""
    vistas, resultado = set(), []
    for seccion in secciones:
        for nombre in nombres:
            cabecera = getattr(seccion, nombre)
            if cabecera.is_linked_to_previous:
                continue
            parte = cabecera.part
            if id(parte) not in vistas:
                vistas.add(id(parte))
                resultado.append(cabecera)
    return resultado


def _resumir(cabecera):
    """Lo que contiene un encabezado, en palabras: textos (sin los números de página), imágenes y tablas."""
    elemento = cabecera._element
    textos = [_texto_sin_campos(p) for p in elemento.iter(w("p"))
              if next(p.iterancestors(w("tbl")), None) is None]
    textos = [t for t in textos if t]
    imagenes = sum(1 for _ in elemento.iter(w("drawing"), w("pict")))
    tablas = sum(1 for _ in elemento.iter(w("tbl")))
    return textos, imagenes, tablas


def _parrafo_numero(cabecera):
    """Deja el encabezado con un solo párrafo: el número de página a la derecha."""
    elemento = cabecera._element
    for hijo in list(elemento):
        elemento.remove(hijo)
    p = cabecera.add_paragraph()._p
    ppr = p.get_or_add_pPr()
    ppr.get_or_add_jc().set(w("val"), "right")
    poner_espaciado(ppr, 240)
    ind = ppr.get_or_add_ind()
    ind.set(w("left"), "0")
    ind.set(w("firstLine"), "0")
    for tipo, texto in (("begin", None), ("instr", " PAGE "), ("separate", None), ("texto", "1"), ("end", None)):
        run = OxmlElement("w:r")
        if tipo == "instr":
            hijo = OxmlElement("w:instrText")
            hijo.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            hijo.text = texto
        elif tipo == "texto":
            hijo = OxmlElement("w:t")
            hijo.text = texto
        else:
            hijo = OxmlElement("w:fldChar")
            hijo.set(w("fldCharType"), tipo)
        run.append(hijo)
        p.append(run)


def _numeracion_continua(doc):
    """Quita reinicios y formatos no arábigos de la numeración. Devuelve qué se ha cambiado."""
    reinicios = romanos = 0
    for seccion in doc.sections:
        tipo = seccion._sectPr.find(w("pgNumType"))
        if tipo is None:
            continue
        if tipo.get(w("start")) is not None:
            del tipo.attrib[w("start")]
            reinicios += 1
        if tipo.get(w("fmt")) not in (None, "decimal"):
            del tipo.attrib[w("fmt")]
            romanos += 1
    return reinicios, romanos


def aplicar(doc, opciones, informe):
    secciones = list(doc.sections)

    # 1. Qué había en los encabezados (para avisar) antes de borrarlo.
    eliminado = {"textos": [], "imagenes": 0, "tablas": 0}
    for cabecera in _partes(secciones, ("header", "first_page_header", "even_page_header")):
        textos, imagenes, tablas = _resumir(cabecera)
        eliminado["textos"] += [t for t in textos if t not in eliminado["textos"]]
        eliminado["imagenes"] += imagenes
        eliminado["tablas"] += tablas

    # 2. Pies: fuera los números de página.
    quitados, conservados = [], []
    for pie in _partes(secciones, ("footer", "first_page_footer", "even_page_footer")):
        for p in list(pie._element.iter(w("p"))):
            original = texto_legible(p).strip()
            if tiene_numero_de_pagina(p):
                if SOLO_NUMERACION.match(_texto_sin_campos(p)):
                    p.getparent().remove(p)
                else:
                    _quitar_campos_de_pagina(p)
                    conservados.append(_texto_sin_campos(p))
                if original not in quitados:
                    quitados.append(original)
            elif original and original not in conservados:
                conservados.append(original)
        if pie._element.find(w("p")) is None:   # un pie no puede quedar sin ningún párrafo
            pie.add_paragraph()

    # 3. Un único encabezado con el número, en todas las páginas.
    doc.settings.odd_and_even_pages_header_footer = False
    for i, seccion in enumerate(secciones):
        seccion.different_first_page_header_footer = False
        seccion.first_page_header.is_linked_to_previous = True
        seccion.even_page_header.is_linked_to_previous = True
        if i:
            seccion.header.is_linked_to_previous = True
    secciones[0].header.is_linked_to_previous = False
    _parrafo_numero(secciones[0].header)
    reinicios, romanos = _numeracion_continua(doc)

    _informar(informe, eliminado, quitados, conservados, reinicios, romanos)


def _informar(informe, eliminado, quitados, conservados, reinicios, romanos):
    informe.cambio("Número de página: arriba a la derecha en todas las páginas, también en la portada.")
    for texto in quitados:
        informe.cambio(f"Se ha quitado el número de página del pie («{texto}»).")
    if eliminado["textos"]:
        textos = "», «".join(eliminado["textos"])
        informe.aviso(f"El encabezado contenía «{textos}». Se ha eliminado porque APA de estudiante no lleva "
                      "encabezado de página (solo el número).")
    for cuantas, singular, plural in ((eliminado["imagenes"], "una imagen", "imágenes"),
                                      (eliminado["tablas"], "una tabla", "tablas")):
        if cuantas:
            que = singular if cuantas == 1 else f"{cuantas} {plural}"
            informe.aviso(f"El encabezado tenía {que}. Se ha eliminado porque APA de estudiante no lleva "
                          "encabezado de página (solo el número).")
    for texto in conservados:
        informe.aviso(f"El pie de página contiene «{texto}». Se ha dejado porque no es un número de página, "
                      "pero APA de estudiante no usa pie de página: revisa si lo necesitas.")
    if reinicios:
        informe.cambio("La numeración de páginas ya no se reinicia en ninguna sección: APA numera todo el "
                       "trabajo seguido, empezando por la portada como página 1.")
    if romanos:
        informe.cambio("Los números de página en romanos (por ejemplo, en el índice) pasan a números normales, "
                       "como pide APA.")
