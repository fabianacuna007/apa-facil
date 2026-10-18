"""
Reglas 4.1 y 4.2: estilos de título de Word (Título 1 a Título 5) con el
formato APA, cada título detectado con el estilo de su nivel, y sin formato
puesto a mano que contradiga ese nivel.

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
  4. Paso 4.2: el formato puesto a mano en el párrafo manda sobre el estilo, así
     que se quita el que contradice el nivel del título: una alineación o una
     sangría distintas, bordes, sombreado, "Conservar con el siguiente"
     desactivado, la negrita desactivada, las mayúsculas forzadas y la
     cursiva. La cursiva solo si va en todo el título: en una parte (el nombre
     de una especie, el título de un libro) APA la usa, y se deja. El color, el
     subrayado y el resaltado ya los quita la limpieza (paso 3.5).
  5. Avisa de lo que APA desaconseja en la estructura: un apartado con un solo
     subapartado y un título que se salta un nivel (de 1 a 3). Y de lo que se
     deja para más adelante: la numeración automática de Word en un título
     (paso 4.3) y el punto final de los niveles 4 y 5, porque añadirlo modifica
     el texto (decisión del propietario: se añade en el paso 4.5, al unir el
     título con su párrafo).

No toca el texto de los títulos (mayúsculas y numeración: paso 4.3). Los rótulos de referencias y de
anexos tienen sus propios pasos (6.4 y 6.7); si ya usaban Título 1, cambian con
el estilo, y quedan centrados y en negrita, como pide APA.
"""

from docx.oxml import OxmlElement

from apa_facil.documento import Estilos, w
from apa_facil.estructura import ACABA_EN_PUNTO, BAJA, clasificar
from apa_facil.reglas.fuente import _letras, _runs_propios

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


# --- 4.2 Formato puesto a mano --------------------------------------------------------

def _apagado(el):
    return el is not None and el.get(w("val")) in ("0", "false", "off")


def _encendido(el):
    return el is not None and not _apagado(el)


def _entero(valor):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def _quitar(padre, *etiquetas):
    """Quita los hijos con esas etiquetas. ¿Había alguno?"""
    quitados = [el for etiqueta in etiquetas for el in padre.findall(w(etiqueta))]
    for el in quitados:
        padre.remove(el)
    return bool(quitados)


def _sangria_de_otro_nivel(ind, sangria):
    """¿Dice la sangría del párrafo algo distinto de la de su nivel?"""
    esperado = {"firstLine": sangria, "hanging": 0, "left": 0, "start": 0,
                "firstLineChars": 0, "hangingChars": 0, "leftChars": 0, "startChars": 0}
    return any(ind.get(w(atributo)) is not None and _entero(ind.get(w(atributo))) != valor
               for atributo, valor in esperado.items())


def cursiva_efectiva(estilos, r, p):
    """¿Se ve en cursiva este fragmento? (formato directo → estilo de carácter → estilo del párrafo…)"""
    for capa in estilos._capas_texto(r, p):
        el = capa.find(w("i")) if capa is not None else None
        if el is not None:
            return not _apagado(el)
    return False


def fragmentos_con_letras(p):
    return [r for r in _runs_propios(p) if _letras(r)]


def _formato_a_mano(estilos, c):
    """Quita del título lo puesto a mano que contradice su nivel. Devuelve qué se ha quitado."""
    alineacion, cursiva, sangria = FORMATO[c.nivel]
    p, quitado = c.elemento, []
    ppr = p.find(w("pPr"))
    if ppr is not None:
        jc = ppr.find(w("jc"))
        if jc is not None and {"start": "left"}.get(jc.get(w("val")), jc.get(w("val"))) != alineacion:
            ppr.remove(jc)
            quitado.append("la alineación")
        ind = ppr.find(w("ind"))
        if ind is not None and _sangria_de_otro_nivel(ind, sangria):
            ppr.remove(ind)
            quitado.append("la sangría")
        if _quitar(ppr, *DECORACION_DE_PARRAFO):
            quitado.append("los bordes o el sombreado")
        if _apagado(ppr.find(w("keepNext"))):
            _quitar(ppr, "keepNext")
            quitado.append("«Conservar con el siguiente» desactivado")

    fragmentos = fragmentos_con_letras(p)
    negrita = mayusculas = False
    for r in fragmentos:
        rpr = r.find(w("rPr"))
        if rpr is None:
            continue
        if _apagado(rpr.find(w("b"))):
            negrita |= _quitar(rpr, "b", "bCs")
        for etiqueta in ("caps", "smallCaps"):
            if _encendido(rpr.find(w(etiqueta))):
                mayusculas |= _quitar(rpr, etiqueta)
    if negrita:
        quitado.append("la negrita desactivada")
    if mayusculas:
        quitado.append("las mayúsculas forzadas")
    # La cursiva, solo si todo el título la tiene al revés que su nivel: en una parte, APA la usa.
    if fragmentos and all(cursiva_efectiva(estilos, r, p) != cursiva for r in fragmentos):
        if any(r.find(w("rPr")) is not None and _quitar(r.find(w("rPr")), "i", "iCs") for r in fragmentos):
            quitado.append("la cursiva desactivada" if cursiva else "la cursiva")
    return quitado


def _lista_de_palabras(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


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

    # 4. Sin formato puesto a mano que contradiga el nivel (paso 4.2).
    estilos = Estilos(doc)
    corregidos = [(c, quitado) for c in seguros for quitado in [_formato_a_mano(estilos, c)] if quitado]

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
    if corregidos:
        varios = len(corregidos) > 1
        ejemplos = "; ".join(f"«{_extracto(c.texto)}»: {_lista_de_palabras(q)}" for c, q in corregidos[:3])
        informe.cambio(f"Títulos: se ha quitado el formato puesto a mano que no corresponde a su nivel en "
                       f"{len(corregidos)} título{'s' if varios else ''} ({ejemplos}"
                       f"{'…' if len(corregidos) > 3 else ''}).")
    for c in titulos:
        if c.confianza == BAJA:
            informe.aviso(f"«{_extracto(c.texto)}» parece un título, pero no es seguro, así que se ha dejado "
                          f"como estaba. Si es un título, ponle en Word el estilo «Título {c.nivel}».")

    # 5. Avisos.
    _avisos_de_estructura(clasificados, informe)
    numerados = [c.texto for c in seguros if c.elemento.find(f"{w('pPr')}/{w('numPr')}") is not None]
    if numerados:
        varios = len(numerados) > 1
        informe.aviso(f"{len(numerados)} título{'s' if varios else ''} ({_lista(numerados)}) "
                      f"lleva{'n' if varios else ''} la numeración automática de Word. APA no numera los "
                      "títulos, pero se ha dejado así: quítala en Word si no la exige tu universidad.")
    sin_punto = [c.texto for c in seguros if c.nivel >= 4 and not ACABA_EN_PUNTO.search(c.texto)]
    if sin_punto:
        varios = len(sin_punto) > 1
        informe.aviso(f"{len(sin_punto)} título{'s' if varios else ''} de nivel 4 o 5 ({_lista(sin_punto)}) "
                      f"no termina{'n' if varios else ''} en punto. En APA, estos títulos terminan en punto y "
                      "el texto sigue en la misma línea; de momento se han dejado como estaban.")
