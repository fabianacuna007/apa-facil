"""
Regla 4.1: estilos de título de Word (Título 1 a Título 5) con el formato APA,
y cada título detectado con el estilo de su nivel.

Usar los estilos de Word, y no formato directo, mantiene el panel de
navegación y el índice de Word, y es lo que recomienda APA por accesibilidad.

Qué hace:
  1. Busca los estilos de título por su nombre interno ("heading 1"…), que
     Word guarda en inglés aunque esté en español; el identificador puede ser
     otro ("Ttulo1"). Si un estilo está definido dos veces con el mismo
     identificador, deja solo la primera definición: si no, el resultado
     dependería de cuál lea Word. Crea, basados en Normal, los que hagan falta.
  2. Da a los estilos Título 1 a 5 el formato APA de su nivel (FORMATO): negrita,
     centrado el 1, cursiva el 3 y el 5, sangría de 1,27 cm el 4 y el 5. Y en
     todos: sin color, subrayado, mayúsculas forzadas, bordes ni sombreado,
     "Conservar con el siguiente" (el título no se queda solo al final de la
     página) y su nivel de esquema, que es lo que lee el panel de navegación.
     La fuente, el tamaño y el interlineado ya los han puesto las reglas
     anteriores en todos los estilos.
  3. Pone a cada título del clasificador el estilo de su nivel. Los de confianza
     baja no: un párrafo de texto convertido por error en título se notaría
     mucho. Se avisa para que el estudiante lo revise.
  4. Avisa de lo que APA desaconseja en la estructura: un apartado con un solo
     subapartado y un título que se salta un nivel (de 1 a 3).

No toca el texto de los títulos (mayúsculas y numeración: paso 4.3) ni el
formato directo de sus párrafos (paso 4.2). Los rótulos de referencias y de
anexos tienen sus propios pasos (6.4 y 6.7); si ya usaban Título 1, cambian con
el estilo, y quedan centrados y en negrita, como pide APA.
"""

from docx.oxml import OxmlElement

from apa_facil.documento import w
from apa_facil.estructura import BAJA, clasificar

NIVELES = range(1, 6)
SANGRIA = 720   # 1,27 cm en twips

# Formato APA de cada nivel: (alineación, cursiva, sangría de primera línea). Todos van en negrita.
FORMATO = {
    1: ("center", False, 0),
    2: ("left", False, 0),
    3: ("left", True, 0),
    4: ("left", False, SANGRIA),
    5: ("left", True, SANGRIA),
}

# Lo que APA no usa en los títulos y Word sí pone en sus estilos (azul, mayúsculas, bordes…).
DECORACION_DE_TEXTO = ("color", "u", "caps", "smallCaps", "highlight", "shd", "spacing", "w", "effect",
                       "outline", "shadow", "emboss", "imprint", "strike", "dstrike")
DECORACION_DE_PARRAFO = ("pBdr", "shd")


def _nombre(estilo):
    nombre = estilo.find(w("name"))
    return (nombre.get(w("val")) or "").lower() if nombre is not None else ""


def nombre_interno(nivel):
    return f"heading {nivel}"


def estilos_de_titulo(doc):
    """Los estilos de título de Word de cada nivel: {1: [w:style, …], …} (puede haber duplicados)."""
    resultado = {nivel: [] for nivel in NIVELES}
    for estilo in doc.styles.element.iter(w("style")):
        if estilo.get(w("type")) != "paragraph":
            continue
        for nivel in NIVELES:
            if _nombre(estilo) == nombre_interno(nivel):
                resultado[nivel].append(estilo)
    return resultado


def _id_de_normal(doc):
    for estilo in doc.styles.element.iter(w("style")):
        if estilo.get(w("type")) == "paragraph" and estilo.get(w("default")) in ("1", "true"):
            return estilo.get(w("styleId"))
    return None


def _quitar_duplicados(estilos):
    """Deja la primera definición de cada estilo: las demás con el mismo identificador se borran."""
    primero = estilos[0]
    for repetido in estilos[1:]:
        if repetido.get(w("styleId")) == primero.get(w("styleId")):
            repetido.getparent().remove(repetido)
    return primero


def _crear(doc, nivel):
    """Un estilo «Título N» nuevo, basado en Normal (hereda su fuente, tamaño e interlineado)."""
    raiz = doc.styles.element
    ids = {e.get(w("styleId")) for e in raiz.iter(w("style"))}
    estilo = OxmlElement("w:style")
    estilo.set(w("type"), "paragraph")
    estilo.set(w("styleId"), next(i for i in (f"Heading{nivel}", f"Titulo{nivel}", f"APA_Titulo{nivel}")
                                  if i not in ids))
    estilo.get_or_add_name().set(w("val"), nombre_interno(nivel))
    normal = _id_de_normal(doc)
    if normal:
        estilo.get_or_add_basedOn().set(w("val"), normal)
    estilo.get_or_add_uiPriority().set(w("val"), "9")
    if nivel > 1:
        estilo.get_or_add_unhideWhenUsed()
    estilo.get_or_add_qFormat()
    raiz.append(estilo)
    return estilo


def _activar(el, activado=True):
    """Pone un ajuste de sí o no (negrita, cursiva…) explícitamente, para que no se herede otro."""
    if activado:
        el.attrib.pop(w("val"), None)
    else:
        el.set(w("val"), "0")


def _quitar_decoracion(rpr, ppr=None):
    for etiqueta in DECORACION_DE_TEXTO:
        for el in rpr.findall(w(etiqueta)):
            rpr.remove(el)
    if ppr is not None:
        for etiqueta in DECORACION_DE_PARRAFO:
            for el in ppr.findall(w(etiqueta)):
                ppr.remove(el)


def _dar_formato(doc, estilo, nivel):
    alineacion, cursiva, sangria = FORMATO[nivel]
    ppr = estilo.get_or_add_pPr()
    ppr.get_or_add_keepNext()
    ppr.get_or_add_keepLines()
    ppr.get_or_add_jc().set(w("val"), alineacion)
    ind = ppr.get_or_add_ind()
    for atributo in list(ind.attrib):
        del ind.attrib[atributo]
    ind.set(w("left"), "0")
    ind.set(w("firstLine"), str(sangria))
    ppr.get_or_add_outlineLvl().set(w("val"), str(nivel - 1))

    rpr = estilo.get_or_add_rPr()
    _quitar_decoracion(rpr, ppr)
    _activar(rpr.get_or_add_b())
    _activar(rpr.get_or_add_bCs())
    _activar(rpr.get_or_add_i(), cursiva)
    _activar(rpr.get_or_add_iCs(), cursiva)

    normal = _id_de_normal(doc)
    if normal:
        estilo.get_or_add_next().set(w("val"), normal)

    # El estilo de carácter vinculado («Título 1 Car») se usa si el título se aplica a parte de un párrafo.
    vinculo = estilo.find(w("link"))
    if vinculo is not None:
        for otro in doc.styles.element.iter(w("style")):
            if otro.get(w("type")) == "character" and otro.get(w("styleId")) == vinculo.get(w("val")):
                if otro.find(w("rPr")) is not None:
                    _quitar_decoracion(otro.find(w("rPr")))


def _extracto(texto, largo=40):
    return texto if len(texto) <= largo else texto[:largo].rstrip() + "…"


def _lista(textos, maximo=3):
    citados = ", ".join(f"«{_extracto(t)}»" for t in textos[:maximo])
    return citados + ("…" if len(textos) > maximo else "")


def _avisos_de_estructura(clasificados, informe):
    """Un apartado con un solo subapartado, y títulos que se saltan un nivel."""
    # Los rótulos de referencias y de anexos cierran los apartados como un título de nivel 1.
    titulos = [(c.nivel if c.tipo == "titulo" else 1, c) for c in clasificados
               if (c.tipo == "titulo" and c.nivel) or c.tipo in ("titulo_referencias", "anexo")]
    anterior = 0
    for i, (nivel, c) in enumerate(titulos):
        if c.tipo == "titulo" and nivel > anterior + 1:
            if anterior:
                previo = titulos[i - 1][1]
                informe.aviso(f"El título «{_extracto(c.texto)}» es de nivel {nivel}, pero el anterior "
                              f"(«{_extracto(previo.texto)}») es de nivel {anterior}: se salta un nivel. APA "
                              f"pide no saltarse niveles: revisa si debería ser de nivel {anterior + 1}.")
            else:
                informe.aviso(f"El primer título del trabajo, «{_extracto(c.texto)}», es de nivel {nivel}. "
                              "APA pide empezar por el nivel 1: revisa su nivel.")
        anterior = nivel

        hijos = []
        for nivel_siguiente, siguiente in titulos[i + 1:]:
            if nivel_siguiente <= nivel:
                break
            if nivel_siguiente == nivel + 1:
                hijos.append(siguiente)
        if len(hijos) == 1:
            informe.aviso(f"El apartado «{_extracto(c.texto)}» tiene un solo subapartado "
                          f"(«{_extracto(hijos[0].texto)}»). APA recomienda dividir un apartado solo si "
                          "tiene al menos dos subapartados: añade otro o integra su contenido en el apartado.")


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    titulos = [c for c in clasificados if c.tipo == "titulo" and c.nivel]
    seguros = [c for c in titulos if c.confianza != BAJA]
    usados = {c.nivel for c in seguros}

    # 1 y 2. Un estilo por nivel (creándolo si un título lo necesita), con el formato APA.
    por_nivel = {}
    for nivel, estilos in estilos_de_titulo(doc).items():
        if estilos:
            por_nivel[nivel] = _quitar_duplicados(estilos)
        elif nivel in usados:
            por_nivel[nivel] = _crear(doc, nivel)
    for nivel, estilo in por_nivel.items():
        _dar_formato(doc, estilo, nivel)
        if estilo.find(f"{w('pPr')}/{w('numPr')}") is not None and nivel in usados:
            informe.aviso(f"El estilo «Título {nivel}» numera automáticamente sus títulos. APA no numera los "
                          "títulos, pero se ha dejado así: quita la numeración del estilo en Word si no la "
                          "exige tu universidad.")

    # 3. Cada título con el estilo de su nivel.
    sin_estilo = []
    for c in seguros:
        destino = por_nivel[c.nivel].get(w("styleId"))
        pstyle = c.elemento.get_or_add_pPr().get_or_add_pStyle()
        if pstyle.get(w("val")) != destino:
            sin_estilo.append(c.texto)
            pstyle.set(w("val"), destino)

    if seguros:   # sin títulos, arreglar los estilos no se nota: no se cuenta
        niveles = sorted(por_nivel)
        rango = (f"«Título {niveles[0]}»" if len(niveles) == 1
                 else f"«Título {niveles[0]}» a «Título {niveles[-1]}»")
        informe.cambio(f"Títulos: los estilos {rango} tienen ahora el formato APA de cada nivel (negrita; "
                       "centrado el nivel 1; cursiva el 3 y el 5; sangría el 4 y el 5), sin colores ni "
                       "subrayados, y un título ya no se queda solo al final de una página.")
    if sin_estilo:
        varios = len(sin_estilo) > 1
        informe.cambio(f"Títulos: {len(sin_estilo)} título{'s' if varios else ''} "
                       f"({_lista(sin_estilo)}) usa{'n' if varios else ''} ahora el estilo de título de Word "
                       "de su nivel, así que aparece" + ("n" if varios else "") +
                       " en el panel de navegación de Word y en el índice automático, si lo usas.")
    for c in titulos:
        if c.confianza == BAJA:
            informe.aviso(f"«{_extracto(c.texto)}» parece un título, pero no es seguro, así que se ha dejado "
                          f"como estaba. Si es un título, ponle en Word el estilo «Título {c.nivel}».")

    # 4. Estructura.
    _avisos_de_estructura(clasificados, informe)
