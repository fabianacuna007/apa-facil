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

Sin título no se genera: APA lo exige. El título del principio del texto (paso
4.4) sale entonces de los datos, con las mismas mayúsculas que en la portada
(paso 4.6; --titulo-tal-cual lo deja como lo escribe el estudiante).

Paso 5.3: si el documento ya tiene una portada (la de la universidad), por
defecto se sustituye por la de APA (decisión del propietario): se quitan sus
párrafos, el logo incluido, y se anota en el informe lo que se ha quitado. Se
mantiene con --mantener-portada (APA admite otra portada si la universidad o el
profesor la exigen), sin datos de portada o si quitarla podría borrar algo que
no es de la portada (una tabla, una nota, un campo, un párrafo largo). Si se
mantiene, se deja con su diseño (decisión del propietario) y solo se hace que el
texto empiece en una página nueva.
"""

from docx.oxml import OxmlElement

from apa_facil.datos_portada import LINEAS_EN_BLANCO_ANTES_DEL_TITULO, parrafos_de_portada
from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import PALABRAS_DE_PARRAFO_LARGO, clasificar
from apa_facil.reglas import titulos
from apa_facil.reglas.titulos import mayusculas_apa

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


# --- 5.3 La portada que ya tenía el documento ------------------------------------------

# Lo que una portada no lleva: si aparece, el clasificador puede haberse equivocado y quitar la
# portada borraría algo del trabajo (o el texto de un cuadro de texto, que cuenta aparte).
NO_SE_QUITA_SI_LLEVA = ("footnoteReference", "endnoteReference", "fldChar", "fldSimple", "commentReference",
                        "txbxContent")
BLOQUES = (w("p"), w("tbl"), w("sdt"))


def portada_existente(clasificados):
    """
    Los párrafos de la portada que ya tenía el documento (vacíos e imágenes incluidos), o []. Cuenta
    también el texto de las tablas que haya entre ellos: hay portadas maquetadas con tablas, con
    todo su texto dentro (revisión del propietario, 2026-10-20).
    """
    portada = [c for c in clasificados if c.zona == "portada"]
    if any(c.texto for c in portada):
        return portada
    return portada if _textos(portada) else []


def _tramo(portada):
    """Los bloques del cuerpo desde el primer párrafo de la portada hasta el último (con las tablas de en medio)."""
    if not portada:
        return []
    primero, ultimo = portada[0].elemento, portada[-1].elemento
    if primero.getparent() is not ultimo.getparent():
        return [c.elemento for c in portada]
    hijos = list(primero.getparent())
    # La portada empieza al principio del documento: también si lo primero es una tabla.
    return hijos[:hijos.index(ultimo) + 1]


def _textos(portada):
    """Los textos de la portada, también los de sus tablas, en orden."""
    return [t for b in _tramo(portada) for p in b.iter(w("p")) if (t := texto_parrafo(p).strip())]


def _por_que_no_se_puede_quitar(portada):
    """None si la portada se puede quitar sin riesgo; si no, el motivo, para el aviso."""
    elementos = [c.elemento for c in portada]
    bloques = [el for el in elementos[0].getparent() if el.tag in BLOQUES]
    if bloques[:len(elementos)] != elementos:
        return "hay una tabla u otro elemento entre sus líneas"
    for numero, c in enumerate(portada):
        if any(next(c.elemento.iter(w(e)), None) is not None for e in NO_SE_QUITA_SI_LLEVA):
            return "lleva una nota, un campo, un comentario o un cuadro de texto"
        if len(c.texto.split()) >= PALABRAS_DE_PARRAFO_LARGO:
            return "tiene un párrafo largo, que podría ser parte del texto"
        if c.elemento.find(f"{w('pPr')}/{w('sectPr')}") is not None and numero != len(portada) - 1:
            return "tiene un salto de sección entre sus líneas"
    return None


def que_hacer(opciones, clasificados):
    """
    Qué se hace con la portada: ("nueva" | "sustituir" | "mantener" | None, motivo). La regla
    del título del trabajo (paso 4.4) lo usa para saber si el título sale de los datos.
    """
    datos, existente = opciones.get("portada"), portada_existente(clasificados)
    if not existente:
        return ("nueva", None) if datos and datos.titulo else (None, "sin_titulo" if datos else None)
    if opciones.get("mantener_portada"):
        return "mantener", "pedido"
    if not datos:
        return "mantener", "sin_datos"
    if not datos.titulo:
        return "mantener", "sin_titulo"
    motivo = _por_que_no_se_puede_quitar(existente)
    return ("mantener", motivo) if motivo else ("sustituir", None)


def intactos(opciones, clasificados):
    """
    Los párrafos (w:p) de la portada que se mantiene, también los de sus tablas y cuadros de texto:
    las reglas de formato no los tocan (revisión del propietario, 2026-10-20: una portada de
    universidad maquetada con tablas perdía sus tamaños, sus espacios y sus márgenes). Vacío si la
    portada se sustituye o no hay. El pipeline lo calcula una vez y lo guarda en la clasificación.
    """
    if que_hacer(opciones, clasificados)[0] != "mantener":
        return frozenset()
    return parrafos_de(clasificados)


def parrafos_de(clasificados):
    """Todos los párrafos (w:p) de la portada que tiene el documento, también los de sus tablas (o ninguno)."""
    return frozenset(p for bloque in _tramo(portada_existente(clasificados)) for p in bloque.iter(w("p")))


def congelar(doc, parrafos):
    """
    Escribe en cada párrafo, como formato propio, el que se ve ahora (espaciado, alineación,
    sangrías, fuente y tamaño): así no cambia cuando las reglas cambian los estilos.
    """
    from apa_facil.documento import poner_espaciado, poner_fuente, poner_tamano
    from apa_facil.reglas.fuente import _letras, _runs_propios
    from apa_facil.reglas.interlineado import _efectivo
    estilos = Estilos(doc)
    for p in parrafos:
        ppr = p.get_or_add_pPr()
        poner_espaciado(ppr, *_efectivo(estilos, p))
        ppr.get_or_add_jc().set(w("val"), estilos.parrafo(p, "jc", "val") or "left")
        ind = ppr.get_or_add_ind()
        for atributo in ("left", "right", "firstLine", "hanging"):
            valor = estilos.parrafo(p, "ind", atributo)
            if valor is not None:
                ind.set(w(atributo), valor)
        for r in _runs_propios(p):
            if _letras(r):
                rpr = r.get_or_add_rPr()
                poner_fuente(rpr, estilos.fuente(r, p))
                poner_tamano(rpr, estilos.tamano(r, p))


def _termina_en_salto(p):
    """¿Acaba el párrafo en un salto de página o de sección que empieza página? (no las marcas de Word)"""
    if any(b.get(w("type")) == "page" for b in p.iter(w("br"))):
        return True
    seccion = p.find(f"{w('pPr')}/{w('sectPr')}")
    tipo = seccion.find(w("type")) if seccion is not None else None
    return seccion is not None and (tipo is None or tipo.get(w("val")) != "continuous")


def _siguiente_bloque(el):
    return next((s for s in el.itersiblings() if s.tag in BLOQUES), None)


def _mantener(doc, informe, existente, motivo, datos):
    """La portada se deja como está; el texto empieza en una página nueva («start each section on a new page»)."""
    primera = _textos(existente)[0]
    inicio = f"Se ha mantenido tu portada («{_extracto(primera, 40)}»)"
    if motivo == "pedido":
        informe.aviso(f"{inicio}, como has pedido. APA pide la portada de estudiante salvo que tu universidad o tu "
                      "profesor exijan otra" + (": no se han usado tus datos de portada." if datos else "."))
    elif motivo == "sin_datos":
        informe.aviso(f"{inicio} porque no has dado los datos para la portada APA de estudiante (título, autores, "
                      "universidad, asignatura, profesor y fecha). Si quieres sustituirla, dalos en un archivo con "
                      "--datos o con las opciones de portada; si tu universidad exige la suya, déjala así.")
    elif motivo == "sin_titulo":
        informe.aviso(f"{inicio} porque falta el título del trabajo, que la portada APA exige. Añádelo en tu "
                      "archivo de datos o con la opción --titulo.")
    else:
        informe.aviso(f"{inicio} y no se ha sustituido por la de APA porque {motivo}: al quitarla podría borrarse "
                      "algo que no es de la portada. Revísala tú.")

    ultimo = existente[-1].elemento
    siguiente = _siguiente_bloque(ultimo)
    if _termina_en_salto(ultimo) or siguiente is None:
        return
    if siguiente.tag != w("p"):
        informe.aviso("Después de tu portada hay una tabla, así que no se ha podido hacer que el texto empiece en "
                      "una página nueva: pon tú un salto de página al final de la portada.")
        return
    if not titulos._salto_de_pagina_antes(Estilos(doc), siguiente):
        _empezar_pagina(siguiente, ultimo)
        informe.cambio("Portada: el texto empieza ahora en una página nueva, después de tu portada, como pide APA.")


def _quitar(existente, informe):
    """
    Quita la portada que tenía el documento. Sus marcadores pasan al párrafo siguiente.
    Devuelve el salto de sección con que acababa (o None), para ponerlo al final de la nueva.
    """
    ultimo = existente[-1].elemento
    seccion = ultimo.find(f"{w('pPr')}/{w('sectPr')}")
    siguiente = _siguiente_bloque(ultimo)
    marcadores = [m for c in existente for m in c.elemento.iter(w("bookmarkStart"), w("bookmarkEnd"))]
    if marcadores and siguiente is not None and siguiente.tag == w("p"):
        ppr = siguiente.find(w("pPr"))
        for marcador in reversed(marcadores):
            (ppr.addnext if ppr is not None else lambda m: siguiente.insert(0, m))(marcador)
    for c in existente:
        c.elemento.getparent().remove(c.elemento)
        informe.cambio_previsto(c.numero, "quitar_portada")
    return seccion


def aplicar(doc, opciones, informe, estructura=None):
    datos = opciones.get("portada")
    clasificados = estructura if estructura is not None else clasificar(doc)
    accion, motivo = que_hacer(opciones, clasificados)
    if accion is None:
        if motivo == "sin_titulo":
            informe.aviso("No se ha añadido la portada APA porque falta el título del trabajo, que APA exige. "
                          "Añádelo en tu archivo de datos o con la opción --titulo.")
        return
    existente = portada_existente(clasificados)
    if accion == "mantener":
        _mantener(doc, informe, existente, motivo, datos)
        return
    id_estilo = _id_del_estilo_del_titulo(doc)
    if id_estilo is None:   # lo crea el paso 4.4 cuando hay un título; aquí siempre debería estar
        informe.error("No se ha podido añadir la portada APA: falta el estilo «Título» de Word.")
        return

    if opciones.get("titulo_tal_cual"):
        titulo, dudosas = datos.titulo, []
    else:
        titulo, dudosas = mayusculas_apa(datos.titulo, clasificados)

    seccion, quitadas, imagenes = None, [], 0
    if accion == "sustituir":
        quitadas = [c.texto for c in existente if c.texto]
        imagenes = sum(1 for c in existente for _ in c.elemento.iter(w("drawing"), w("pict")))
        seccion = _quitar(existente, informe)

    cuerpo = doc.element.body
    primero_del_documento = next((el for el in cuerpo if el.tag in BLOQUES), None)
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
    if seccion is not None:
        # La portada era una sección propia (el 05): el salto de sección pasa a la nueva, que así
        # sigue acabando en él (el número de secciones no cambia) y el texto empieza página.
        portada[-1].get_or_add_pPr().append(seccion)
    elif _empezar_pagina(_siguiente_bloque(portada[-1]), portada[-1]):
        textos[-1] += "\n"
    informe.cambio_previsto(1, "añadir_portada", parrafos=tuple(textos), datos=datos)

    if accion == "sustituir":
        logo = f" y {'la imagen' if imagenes == 1 else f'las {imagenes} imágenes'} (el logo)" if imagenes else ""
        informe.cambio(f"Portada: se ha sustituido tu portada por la portada APA de estudiante. Se han quitado sus "
                       f"{len(quitadas)} líneas ({titulos._lista(quitadas)}){logo}. Revisa que tus datos de "
                       "portada no se dejen nada que necesites; si tu universidad exige su propia portada, usa la "
                       "opción --mantener-portada.")
    informe.cambio("Portada: se ha añadido la portada APA de estudiante con " + _lista(_elementos(datos)) +
                   ", centrados y a doble espacio, con el título en negrita tres líneas por debajo del margen "
                   "superior. El texto empieza en una página nueva.")
    if mensaje_del_titulo:
        informe.cambio(mensaje_del_titulo)
    if titulo != datos.titulo:
        informe.cambio(f"Portada: el título lleva ahora mayúscula inicial en las palabras principales, como pide "
                       f"APA: «{_extracto(datos.titulo)}» → «{_extracto(titulo)}». Revísalo, sobre todo los "
                       "nombres propios; si prefieres dejarlo como lo escribiste, usa la opción --titulo-tal-cual.")
    titulos.avisar_dudosas(informe, dudosas, "el título de la portada")
