"""
Reglas de la fase 4 (pasos 4.1 a 4.5): estilos de título de Word (Título 1 a
Título 5) con el formato APA, cada título detectado con el estilo de su nivel y
sin formato puesto a mano que contradiga ese nivel, el título del trabajo al
principio del texto y los niveles 4 y 5 en la misma línea que su texto.

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
     subapartado y un título que se salta un nivel (de 1 a 3). Y de la
     numeración automática de Word en un título, que se deja.

  6. Paso 4.3: los títulos numerados ("1.1. Objetivos"). Quitar la numeración
     modifica el texto, así que por defecto solo se avisa (decisión del
     propietario); con --quitar-numeracion se quita, y se anota como cambio
     previsto para la comprobación de integridad. Paso 4.6: las mayúsculas de
     APA (mayúscula inicial en las palabras principales), por defecto (decisión
     del propietario; --titulos-tal-cual las deja), también en los títulos
     escritos en mayúsculas ("RESUMEN" → "Resumen").
  7. Paso 4.4: el título del trabajo, copiado de la portada, al principio del
     texto, y el rótulo «Introducción» (aviso, o --quitar-introduccion).
  8. Paso 4.5: los títulos de nivel 4 y 5, en la misma línea que su texto (con
     el separador de estilo de Word) y terminados en punto.

Fuera de eso, no toca el texto de los títulos. Los rótulos de referencias y de
anexos tienen sus propios pasos (6.4 y 6.7); si ya usaban Título 1, cambian con
el estilo, y quedan centrados y en negrita, como pide APA.
"""

import copy
import re
from collections import Counter

from docx.oxml import OxmlElement

from apa_facil.documento import (NUMERACION_DE_APARTADO, Estilos, es_rotulo_de_introduccion, lineas_de_texto,
                                 texto_parrafo, w)
from apa_facil.estructura import BAJA, clasificar
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
    return _estilo_nuevo(doc, nombre_interno(nivel), (f"Heading{nivel}", f"Titulo{nivel}", f"APA_Titulo{nivel}"),
                         prioridad=9, oculto_hasta_usarlo=nivel > 1)


def _estilo_nuevo(doc, nombre, identificadores, prioridad, oculto_hasta_usarlo=False):
    raiz = doc.styles.element
    ids = {e.get(w("styleId")) for e in raiz.iter(w("style"))}
    estilo = OxmlElement("w:style")
    estilo.set(w("type"), "paragraph")
    estilo.set(w("styleId"), next(i for i in identificadores if i not in ids))
    estilo.get_or_add_name().set(w("val"), nombre)
    normal = _id_de_normal(doc)
    if normal:
        estilo.get_or_add_basedOn().set(w("val"), normal)
    estilo.get_or_add_uiPriority().set(w("val"), str(prioridad))
    if oculto_hasta_usarlo:
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


def _dar_formato(doc, estilo, nivel, esquema=True):
    """
    El formato APA del nivel en el estilo. Con esquema=False (el estilo «Title» del
    título del trabajo), sin nivel de esquema: la portada del 02 usa ese estilo y no
    es un título; el título del trabajo lo lleva en su propio párrafo.
    """
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
    if esquema:
        ppr.get_or_add_outlineLvl().set(w("val"), str(nivel - 1))
    else:
        _quitar(ppr, "outlineLvl", "numPr")

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


# --- 4.3 Numeración y mayúsculas --------------------------------------------------------
#
# Quitar la numeración o las mayúsculas de un título modifica el texto. Por eso,
# por defecto solo se avisa (decisión del propietario), y se cambia únicamente
# con las opciones --quitar-numeracion y --quitar-mayusculas. Cada cambio se anota
# como cambio previsto, y la comprobación de integridad acepta solo eso.

PALABRA = re.compile(r"[^\W\d_]+")
# La numeración del principio, aunque no la siga un espacio (cuando la separa un tabulador del XML).
NUMERACION_PEGADA = re.compile(r"(?:\d{1,2}(?:\.\d{1,2})*[.)]|\d{1,2}(?:\.\d{1,2})+|[IVXLC]{1,6}[.)]|"
                               r"[A-Za-z][.)])\s*")
ROMANO = re.compile(r"M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})")
NO_SON_ROMANOS = {"mi", "di", "vi", "li"}   # palabras que también son números romanos válidos
REMISION = re.compile(r"\b(apartados?|secci[oó]n|secciones|cap[ií]tulos?|ep[ií]grafes?)\s+"
                      r"(\d{1,2}(?:\.\d{1,2})*)\b", re.IGNORECASE)


def numeracion(texto):
    """La numeración de apartado con que empieza un título ("1.1. "), o None."""
    encontrada = NUMERACION_DE_APARTADO.match(texto)
    return encontrada.group(0) if encontrada and texto[encontrada.end():].strip() else None


def en_mayusculas(texto):
    """¿Está escrito entero en mayúsculas (sin contar la numeración del principio)?"""
    letras = [c for c in texto[len(numeracion(texto) or ""):] if c.isalpha()]
    return len(letras) >= 2 and all(c.isupper() for c in letras)


def _piezas(p):
    """Los trozos de texto del párrafo (w:t y tabuladores escritos), en orden, sin los de cuadros de texto."""
    return [el for el in p.iter(w("t"), w("tab")) if next(el.iterancestors(w("p")), None) is p
            and (el.tag == w("t") or el.getparent().tag == w("r"))]


def _quitar_numeracion(p, quitado):
    """Quita `quitado` del principio del párrafo, aunque esté repartido en varios fragmentos."""
    falta = len(quitado)
    for pieza in _piezas(p):
        if not falta:
            break
        if pieza.tag == w("tab"):
            pieza.getparent().remove(pieza)
            falta -= 1
            continue
        texto = pieza.text or ""
        pieza.text, falta = texto[falta:], max(0, falta - len(texto))


def _formas_del_texto(clasificados):
    """
    Cómo se escribe cada palabra a mitad de frase en el resto del trabajo, si no es
    en minúscula: {"españa": "España", "ue": "UE"}. Así se conservan los nombres
    propios y las siglas al quitar las mayúsculas de un título.
    """
    cuentas = {}
    for c in clasificados:
        if c.tipo in ("titulo", "titulo_referencias", "anexo", "vacio", "imagen") or en_mayusculas(c.texto):
            continue
        for frase in re.split(r"[.!?¿¡]\s+", c.texto):
            for palabra in PALABRA.findall(frase)[1:]:   # la primera de una frase va en mayúscula siempre
                cuentas.setdefault(palabra.lower(), Counter())[palabra] += 1
    formas = {}
    for minuscula, cuenta in cuentas.items():
        forma, veces = cuenta.most_common(1)[0]
        if forma != minuscula and veces > cuenta[minuscula]:
            formas[minuscula] = forma
    return formas


def a_minusculas(texto, formas):
    """
    "RESUMEN" → "Resumen", "1. LA UE EN ESPAÑA" → "1. La UE en España" (si el trabajo
    escribe así "UE" y "España"). Mayúscula inicial, nombres propios y siglas como
    en el resto del trabajo, y números romanos ("SIGLO XIX") como están.
    """
    prefijo = NUMERACION_PEGADA.match(texto)
    inicio = prefijo.end() if prefijo else 0
    vistas = []

    def cambiar(encontrada):
        palabra = encontrada.group(0)
        minuscula = palabra.lower()
        if minuscula in formas:
            forma = formas[minuscula]
        elif len(palabra) >= 2 and ROMANO.fullmatch(palabra) and minuscula not in NO_SON_ROMANOS:
            forma = palabra
        else:
            forma = minuscula
        vistas.append(palabra)
        return forma[0].upper() + forma[1:] if len(vistas) == 1 else forma

    return texto[:inicio] + PALABRA.sub(cambiar, texto[inicio:])


def _lineas_de_piezas(p):
    """Los w:t del párrafo agrupados por líneas (separadas por saltos de línea), sin los de cuadros de texto."""
    lineas = [[]]
    for el in p.iter(w("t"), w("br"), w("cr")):
        if next(el.iterancestors(w("p")), None) is not p:
            continue
        if el.tag == w("t"):
            lineas[-1].append(el)
        elif el.get(w("type")) in (None, "textWrapping"):
            lineas.append([])
    return [linea for linea in lineas if linea]


def _poner_mayusculas(p, cambiar):
    """
    Cambia las mayúsculas del texto del párrafo, línea a línea (`cambiar` recibe el
    texto de una línea y devuelve el nuevo y sus palabras dudosas), dentro de cada
    fragmento: el formato no se mueve. Devuelve las palabras dudosas, o None si no
    se ha podido (una letra que cambia de longitud, como «ß»).
    """
    nuevos, dudosas = [], []
    for piezas in _lineas_de_piezas(p):
        completo = "".join(t.text or "" for t in piezas)
        nuevo, dudas = cambiar(completo)
        if len(nuevo) != len(completo):
            return None
        nuevos.append((piezas, nuevo))
        dudosas += dudas
    for piezas, nuevo in nuevos:
        desde = 0
        for t in piezas:
            largo = len(t.text or "")
            t.text = nuevo[desde:desde + largo]
            desde += largo
    return dudosas


# --- 4.6 Mayúsculas de APA en los títulos («title case») -------------------------------
#
# La web oficial de APA (Title case capitalization) pide mayúscula inicial en las
# palabras principales en el título del trabajo y en todos los títulos (niveles 1 a
# 5). Palabras menores, en minúscula: todos los artículos y las conjunciones y
# preposiciones de tres letras o menos; las demás (sustantivos, verbos, adjetivos,
# adverbios, pronombres y cualquier palabra de cuatro letras o más) van con
# mayúscula, y también la primera del título, la del subtítulo y la que sigue a dos
# puntos, raya o punto. Las palabras que pueden ser de las dos clases según la
# frase («que», «lo», «vía») se dejan como están y se avisa. Decisión del
# propietario: se aplica por defecto (el título de la portada y del texto se deja
# con --titulo-tal-cual; los de los apartados, con --titulos-tal-cual).

# Palabras menores en español, con los criterios de APA.
ARTICULOS = {"el", "la", "los", "las", "un", "una", "unos", "unas"}
CONTRACCIONES = {"al", "del"}   # preposición y artículo
CONJUNCIONES = {"y", "e", "o", "u", "ni", "si", "mas"}
PREPOSICIONES = {"a", "de", "en", "con", "por", "sin", "so", "pro"}
MENORES = ARTICULOS | CONTRACCIONES | CONJUNCIONES | PREPOSICIONES
# Según la frase son menores (conjunción, artículo, preposición) o principales (pronombre, sustantivo).
DUDOSAS = {"que", "lo", "vía", "via"}

PALABRA_DE_TITULO = re.compile(r"[^\W\d_]+(?:['’][^\W\d_]+)*")   # «O'Donnell» es una sola palabra
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
        if c.tipo in ("titulo", "titulo_referencias", "anexo", "vacio", "imagen") or en_mayusculas(c.texto):
            continue
        for frase in re.split(r"[.!?¿¡:]\s+", c.texto):
            palabras = PALABRA_DE_TITULO.findall(frase)
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
    - La numeración del principio se deja como está.
    """
    prefijo = numeracion(titulo) or ""   # la numeración del principio («a) », «IV. ») no se toca
    if prefijo:
        resto, dudosas = mayusculas_apa(titulo[len(prefijo):], clasificados)
        return prefijo + resto, dudosas
    if en_mayusculas(titulo):
        titulo = a_minusculas(titulo, _formas_del_texto(clasificados))
    mixto = any(c.islower() for c in titulo)
    propios = _nombres_propios(clasificados)
    piezas = list(PALABRA_DE_TITULO.finditer(titulo))
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
            nombre_propio = bool(siguiente) and (_con_mayuscula(minuscula), _con_mayuscula(siguiente)) in propios
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



def avisar_dudosas(informe, dudosas, donde):
    """Aviso de las palabras que se han dejado como estaban porque pueden ser de las dos clases."""
    vistas = list(dict.fromkeys(dudosas))
    if not vistas:
        return
    varias = len(vistas) > 1
    informe.aviso(f"En {donde} se ha{'n' if varias else ''} dejado como estaba{'n' if varias else ''} " +
                  ", ".join(f"«{p}»" for p in vistas) + ": APA pone mayúscula a los pronombres y a los "
                  "sustantivos, pero no a los artículos, ni a las conjunciones y preposiciones cortas. "
                  "Revisa cuál es en tu caso.")


def _texto_de_titulos(doc, opciones, informe, clasificados, seguros):
    con_numeracion = [c for c in seguros if numeracion(texto_parrafo(c.elemento))]
    quitar_numeracion, tal_cual = opciones.get("quitar_numeracion"), opciones.get("titulos_tal_cual")

    if con_numeracion and not quitar_numeracion:
        varios = len(con_numeracion) > 1
        informe.aviso(f"{len(con_numeracion)} título{'s' if varios else ''} "
                      f"empieza{'n' if varios else ''} con una numeración ({_lista([c.texto for c in con_numeracion])}). "
                      "APA no numera los títulos, pero muchas universidades lo exigen, así que se ha dejado como "
                      "estaba. Si la tuya no lo pide, la herramienta puede quitarla: usa la opción --quitar-numeracion.")
    if tal_cual:
        sin_apa = [c.texto for c in seguros if mayusculas_apa(c.texto, clasificados)[0] != c.texto]
        if sin_apa:
            varios = len(sin_apa) > 1
            informe.aviso(f"{len(sin_apa)} título{'s' if varios else ''} no lleva{'n' if varios else ''} las "
                          f"mayúsculas de APA ({_lista(sin_apa)}): APA pone mayúscula inicial en las palabras "
                          "principales. Se ha dejado como estaba porque has usado la opción --titulos-tal-cual.")

    cambiados, con_mayusculas, dudosas = [], [], []   # (antes, después) de cada título modificado
    for c in seguros:
        p = c.elemento
        antes = texto_parrafo(p)
        quitado = numeracion(antes) if quitar_numeracion else None
        if quitado:
            _quitar_numeracion(p, quitado)
            informe.cambio_previsto(c.numero, "quitar_numeracion", quitado=quitado)
        if not tal_cual:
            sin_cambiar = texto_parrafo(p)
            dudosas += _poner_mayusculas(p, lambda texto: mayusculas_apa(texto, clasificados)) or []
            if texto_parrafo(p) != sin_cambiar:
                informe.cambio_previsto(c.numero, "mayusculas_apa", nuevo=texto_parrafo(p))
                con_mayusculas.append((sin_cambiar, texto_parrafo(p)))
        if texto_parrafo(p) != antes:
            cambiados.append((antes, texto_parrafo(p)))

    if quitar_numeracion and con_numeracion:
        quitados = [(a, d) for a, d in cambiados if numeracion(a)]
        informe.cambio(f"Títulos: se ha quitado la numeración de {len(quitados)} "
                       f"título{'s' if len(quitados) > 1 else ''} ({_antes_y_despues(quitados)}), como pide APA.")
        remisiones = [encontrada.group(0) for c in clasificados if c.tipo != "titulo"
                      for encontrada in REMISION.finditer(c.texto)]
        if remisiones:
            informe.aviso(f"El texto remite a apartados por su número ({_lista(remisiones)}). Sin la numeración "
                          "de los títulos, esas remisiones ya no se entienden: cámbialas por el nombre del apartado.")
    if con_mayusculas:
        varios = len(con_mayusculas) > 1
        informe.cambio(f"Títulos: {len(con_mayusculas)} título{'s' if varios else ''} lleva{'n' if varios else ''} "
                       "ahora mayúscula inicial en las palabras principales, como pide APA "
                       f"({_antes_y_despues(con_mayusculas)}). Se han respetado los nombres propios y las siglas que "
                       "el trabajo escribe así en el texto; revisa que hayan quedado bien. Si prefieres dejar los "
                       "títulos como estaban, usa la opción --titulos-tal-cual.")
    avisar_dudosas(informe, dudosas, "los títulos")
    indice = any(c.tipo == "indice" for c in clasificados) and not opciones.get("quitar_indice")
    if cambiados and indice and doc.settings.element.find(w("updateFields")) is None:
        informe.aviso("El índice todavía muestra los títulos como estaban. Actualízalo en Word: clic derecho sobre "
                      "el índice, «Actualizar campos» y «Actualizar toda la tabla» (si lo escribiste a mano, "
                      "cámbialo tú).")


# --- 4.4 Título del trabajo y rótulo «Introducción» -------------------------------------
#
# APA repite el título del trabajo al principio del texto, centrado y en negrita
# como un nivel 1, y la introducción no lleva rótulo. El título se copia tal cual de
# la portada (lo encuentra el clasificador, sobre el original): nada se inventa, y la
# comprobación de integridad exige que sea una copia exacta. Va con el estilo «Title»
# de Word («Título»), con el formato del nivel 1, y con su nivel de esquema en el
# propio párrafo: así sale en el panel de navegación, pero no en el índice
# automático, que recoge los estilos Título 1 a 3.

NOMBRE_DEL_ESTILO_DEL_TITULO = "title"
PIEZAS_QUE_SE_COPIAN = ("rPr", "t", "tab", "br", "cr", "noBreakHyphen", "sym")
TEXTO_QUE_SE_CONSERVA = ("i", "iCs", "vertAlign")   # cursiva (el título de una obra) y superíndices
NO_SE_BORRA_SI_LLEVA = ("sectPr", "footnoteReference", "endnoteReference", "drawing", "pict", "fldChar",
                        "fldSimple", "commentReference")


def estilos_del_titulo_del_trabajo(doc):
    return [e for e in doc.styles.element.iter(w("style"))
            if e.get(w("type")) == "paragraph" and _nombre(e) == NOMBRE_DEL_ESTILO_DEL_TITULO]


def _copia_del_titulo(id_estilo, partes):
    """Un párrafo nuevo con el texto de los párrafos del título (separados por un salto de línea)."""
    nuevo = OxmlElement("w:p")
    ppr = nuevo.get_or_add_pPr()
    ppr.get_or_add_pStyle().set(w("val"), id_estilo)
    ppr.get_or_add_outlineLvl().set(w("val"), "0")
    for numero, parte in enumerate(partes):
        if numero:
            salto = OxmlElement("w:r")
            salto.append(OxmlElement("w:br"))
            nuevo.append(salto)
        copias = []
        for r in _runs_propios(parte.elemento):
            copia = OxmlElement("w:r")
            for pieza in r:
                if etiqueta_de(pieza) == "br" and pieza.get(w("type")) in ("page", "column"):
                    continue   # el salto de página de la portada no se copia
                if etiqueta_de(pieza) in PIEZAS_QUE_SE_COPIAN:
                    copia.append(copy.deepcopy(pieza))
            rpr = copia.find(w("rPr"))
            if rpr is not None:   # el resto del formato lo pone el estilo
                for propiedad in list(rpr):
                    if etiqueta_de(propiedad) not in TEXTO_QUE_SE_CONSERVA:
                        rpr.remove(propiedad)
            if len(copia):
                copias.append(copia)
        _sin_espacio_en_los_extremos(copias)
        for copia in copias:
            nuevo.append(copia)
    return nuevo


def _sin_espacio_en_los_extremos(runs):
    """Quita los espacios, tabuladores y saltos del principio y del final del texto copiado."""
    piezas = [el for r in runs for el in r if etiqueta_de(el) in ("t", "tab", "br", "cr")]
    for orden in (piezas, list(reversed(piezas))):
        for pieza in orden:
            if pieza.getparent() is None:
                continue   # ya quitada desde el otro extremo
            if etiqueta_de(pieza) != "t":
                pieza.getparent().remove(pieza)
                continue
            texto = pieza.text or ""
            limpio = texto.lstrip() if orden is piezas else texto.rstrip()
            pieza.text = limpio
            if limpio:
                break


def etiqueta_de(el):
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


def _salto_de_pagina_antes(estilos, p):
    return _encendido(next((capa.find(w("pageBreakBefore")) for capa in estilos._capas_parrafo(p)
                            if capa is not None and capa.find(w("pageBreakBefore")) is not None), None))


def _subapartados(clasificados, apartado):
    """Los títulos de nivel más bajo que van dentro de un apartado (hasta el siguiente de su nivel)."""
    dentro, empezado = [], False
    for c in clasificados:
        if c is apartado:
            empezado = True
            continue
        if not empezado or c.tipo not in ("titulo", "titulo_referencias", "anexo"):
            continue
        if c.tipo != "titulo" or (c.nivel or 1) <= apartado.nivel:
            break
        dentro.append(c)
    return dentro


def _quitar_introduccion(doc, opciones, informe, clasificados):
    """El rótulo «Introducción» del primer apartado. Devuelve el párrafo si se ha borrado."""
    primero = next((c for c in clasificados if c.zona == "cuerpo" and c.tipo == "titulo"), None)
    if primero is None or primero.nivel != 1 or not es_rotulo_de_introduccion(texto_parrafo(primero.elemento)):
        return None
    rotulo = texto_parrafo(primero.elemento)
    if not opciones.get("quitar_introduccion"):
        informe.aviso(f"El primer apartado se titula «{rotulo}». En APA la introducción no lleva título: el título "
                      "del trabajo, al principio del texto, hace sus veces. Se ha dejado así; quítalo tú o usa la "
                      "opción --quitar-introduccion.")
        return None
    subapartados = _subapartados(clasificados, primero)
    if subapartados:
        informe.aviso(f"No se ha quitado el rótulo «{rotulo}» porque el apartado tiene subapartados "
                      f"({_lista([c.texto for c in subapartados])}): sin él, quedarían colgando del título del "
                      "trabajo. Si quieres seguir APA, deja la introducción sin subapartados.")
        return None
    p = primero.elemento
    if any(next(p.iter(w(e)), None) is not None for e in NO_SE_BORRA_SI_LLEVA):
        informe.aviso(f"No se ha quitado el rótulo «{rotulo}» porque lleva algo más que texto (una nota, una "
                      "imagen, un campo o un salto de sección). Quítalo tú en Word.")
        return None
    # Los marcadores (los usan el índice y las referencias cruzadas) pasan al párrafo siguiente.
    siguiente = p.getnext()
    marcadores = [m for m in p.iter(w("bookmarkStart"), w("bookmarkEnd"))]
    if marcadores and (siguiente is None or siguiente.tag != w("p")):
        informe.aviso(f"No se ha quitado el rótulo «{rotulo}» porque lleva un marcador de Word. Quítalo tú.")
        return None
    for marcador in reversed(marcadores):
        ppr = siguiente.find(w("pPr"))
        (ppr.addnext if ppr is not None else lambda m: siguiente.insert(0, m))(marcador)
    p.getparent().remove(p)
    informe.cambio_previsto(primero.numero, "quitar_introduccion")
    informe.cambio(f"Títulos: se ha quitado el rótulo «{rotulo}» del primer apartado: en APA la introducción no "
                   "lleva título.")
    return primero


def _titulo_del_trabajo(doc, opciones, informe, clasificados, quitado):
    # El estilo «Title», con el formato APA del nivel 1 (la portada del 02 también lo usa,
    # y la portada APA del paso 5.2).
    from apa_facil.reglas import portada   # aquí: la regla de la portada usa este módulo
    partes = [c for c in clasificados if c.titulo_del_trabajo]
    # Con portada APA (nueva o en lugar de la que había, pasos 5.2 y 5.3), el título sale de los datos.
    portada_apa = portada.que_hacer(opciones, clasificados)[0] in ("nueva", "sustituir")
    estilos = estilos_del_titulo_del_trabajo(doc)
    estilo = (_quitar_duplicados(estilos) if estilos else
              _estilo_nuevo(doc, NOMBRE_DEL_ESTILO_DEL_TITULO, ("Title", "Ttulo", "APA_Titulo"), prioridad=10)
              if partes or portada_apa else None)
    if estilo is not None:
        _dar_formato(doc, estilo, 1, esquema=False)

    if portada_apa:
        return   # el título lo pone la regla de la portada (paso 5.2), con el de los datos
    if not partes:
        informe.aviso("No se ha encontrado el título del trabajo en la portada, así que no se ha repetido al "
                      "principio del texto. APA lo pide ahí, centrado y en negrita: escríbelo tú, con el estilo "
                      "«Título» de Word.")
        return
    primero = next((c for c in clasificados if c.zona == "cuerpo" and c.tipo != "vacio" and c is not quitado), None)
    if primero is None:
        return
    texto = "\n".join(linea for c in partes for linea in lineas_de_texto(texto_parrafo(c.elemento)))
    if " ".join(texto_parrafo(primero.elemento).split()).lower() == " ".join(texto.split()).lower():
        return   # ya está (quizá con otras mayúsculas: las pone el paso 4.6, porque es un título)
    nuevo = _copia_del_titulo(estilo.get(w("styleId")), partes)
    if texto_parrafo(nuevo) != texto:
        informe.aviso("No se ha podido repetir el título del trabajo al principio del texto: escríbelo tú, "
                      "centrado y en negrita, con el estilo «Título» de Word.")
        return
    dudosas = []
    if not opciones.get("titulo_tal_cual"):   # paso 4.6: las mayúsculas de APA
        dudosas = _poner_mayusculas(nuevo, lambda linea: mayusculas_apa(linea, clasificados)) or []
        texto = texto_parrafo(nuevo)
    estilos_efectivos = Estilos(doc)
    if _salto_de_pagina_antes(estilos_efectivos, primero.elemento):
        # Si el primer apartado empezaba página (su estilo lo pide), la empieza el título.
        nuevo.find(w("pPr")).get_or_add_pageBreakBefore()
        primero.elemento.get_or_add_pPr().get_or_add_pageBreakBefore().set(w("val"), "0")
    primero.elemento.addprevious(nuevo)
    informe.cambio_previsto(primero.numero, "añadir_titulo", nuevo=texto, desde=[c.numero for c in partes])
    informe.cambio(f"Título del trabajo: se ha repetido «{_extracto(' '.join(texto.split()), 60)}» al principio "
                   "del texto, centrado y en negrita, como pide APA.")
    original = "\n".join(lineas_de_texto("\n".join(texto_parrafo(c.elemento) for c in partes)))
    if texto != original:
        informe.cambio(f"Título del trabajo: al principio del texto lleva mayúscula inicial en las palabras "
                       f"principales, como pide APA («{_extracto(' '.join(original.split()), 50)}» → "
                       f"«{_extracto(' '.join(texto.split()), 50)}»). Si prefieres dejarlo como en la portada, usa "
                       "la opción --titulo-tal-cual.")
    avisar_dudosas(informe, dudosas, "el título del trabajo")
    if opciones.get("titulo_tal_cual") and any(en_mayusculas(c.texto) for c in partes):
        informe.aviso(f"El título del trabajo está escrito en mayúsculas ({_lista([c.texto for c in partes])}). "
                      "APA lo escribe con mayúscula inicial en las palabras principales: se ha dejado así porque "
                      "has usado la opción --titulo-tal-cual.")


# --- 4.5 Títulos de nivel 4 y 5 en la misma línea que su texto --------------------------
#
# APA pone los títulos de nivel 4 y 5 en la misma línea que el párrafo que los
# sigue, terminados en punto. No se unen los dos párrafos: se usa el «separador de
# estilo» de Word (lo que hace Ctrl+Alt+Intro), que oculta la marca de párrafo del
# título. Así Word muestra título y texto en una línea, pero en el documento siguen
# siendo dos párrafos: el título conserva su estilo (panel de navegación, índice) y
# el texto no se toca. Lo único que cambia es el final del título: el punto y un
# espacio (decisión del propietario: por defecto). LibreOffice no entiende el
# separador de estilo y muestra el título en su propia línea.

ACABA_CON_SIGNO = re.compile(r"[.…?!][\"'»”)]*$")
ACABA_CON_DOS_PUNTOS = re.compile(r"[:;][\"'»”)]*$")
ETIQUETAS_DE_BLOQUE = (w("p"), w("tbl"), w("sdt"), w("sectPr"))


def siguiente_bloque(p):
    """Lo que va justo después del párrafo en el documento (otro párrafo, una tabla…), sin contar marcadores."""
    siguiente = p.getnext()
    while siguiente is not None and siguiente.tag not in ETIQUETAS_DE_BLOQUE:
        siguiente = siguiente.getnext()
    return siguiente


def en_linea(p):
    """¿Tiene el párrafo la marca de párrafo oculta (el separador de estilo de Word)?"""
    rpr = p.find(f"{w('pPr')}/{w('rPr')}")
    return rpr is not None and all(_encendido(rpr.find(w(e))) for e in ("vanish", "specVanish"))


def _ocultar_marca_de_parrafo(p):
    ppr = p.get_or_add_pPr()
    rpr = ppr.find(w("rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        despues = next((ppr.find(w(e)) for e in ("sectPr", "pPrChange") if ppr.find(w(e)) is not None), None)
        if despues is not None:
            despues.addprevious(rpr)
        else:
            ppr.append(rpr)
    for etiqueta in ("vanish", "specVanish"):
        _activar(rpr.find(w(etiqueta)) if rpr.find(w(etiqueta)) is not None
                 else getattr(rpr, f"get_or_add_{etiqueta}")())


def _anadir_al_final(p, anadido):
    """Un fragmento nuevo al final del párrafo, con el formato del último que tiene texto."""
    ultimo = next((r for r in reversed(fragmentos_con_letras(p))), None)
    r = OxmlElement("w:r")
    if ultimo is not None and ultimo.find(w("rPr")) is not None:
        r.append(copy.deepcopy(ultimo.find(w("rPr"))))
    t = OxmlElement("w:t")
    t.text = anadido
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p.append(r)


def _titulos_en_linea(informe, clasificados, seguros):
    por_elemento = {c.elemento: c for c in clasificados}
    unidos, sin_texto, con_dos_puntos = [], [], []
    for c in seguros:
        if c.nivel < 4 or c.elemento.getparent() is None:
            continue
        siguiente = siguiente_bloque(c.elemento)
        detras = por_elemento.get(siguiente)
        if detras is None or detras.tipo != "cuerpo":
            sin_texto.append(c.texto)
            continue
        texto = texto_parrafo(c.elemento)
        if not texto.strip() or texto[-1].isspace():
            sin_texto.append(c.texto)   # acaba en espacio: no se le puede poner el punto detrás
            continue
        anadido = " " if ACABA_CON_SIGNO.search(texto) or ACABA_CON_DOS_PUNTOS.search(texto) else ". "
        if ACABA_CON_DOS_PUNTOS.search(texto):
            con_dos_puntos.append(c.texto)
        _anadir_al_final(c.elemento, anadido)
        _ocultar_marca_de_parrafo(c.elemento)
        informe.cambio_previsto(c.numero, "añadir_al_final", nuevo=anadido)
        unidos.append(c.texto)

    if unidos:
        varios = len(unidos) > 1
        informe.cambio(f"Títulos: {len(unidos)} título{'s' if varios else ''} de nivel 4 o 5 ({_lista(unidos)}) "
                       f"va{'n' if varios else ''} ahora en la misma línea que su texto y termina"
                       f"{'n' if varios else ''} en punto, como pide APA. En Word siguen siendo títulos (salen en "
                       "el panel de navegación); en LibreOffice se ven en su propia línea.")
    if con_dos_puntos:
        informe.aviso(f"{_lista(con_dos_puntos)} termina{'n' if len(con_dos_puntos) > 1 else ''} en dos puntos o "
                      "punto y coma. En APA, los títulos de nivel 4 y 5 terminan en punto: cámbialo tú.")
    if sin_texto:
        varios = len(sin_texto) > 1
        informe.aviso(f"{len(sin_texto)} título{'s' if varios else ''} de nivel 4 o 5 ({_lista(sin_texto)}) no "
                      f"va{'n' if varios else ''} seguido{'s' if varios else ''} de un párrafo de texto, así que no "
                      "se ha podido poner en la misma línea que el texto, como pide APA, ni terminar en punto. "
                      "Revisa si debería ser de otro nivel.")


def _antes_y_despues(pares, maximo=3):
    texto = ", ".join(f"«{_extracto(a)}» → «{_extracto(d)}»" for a, d in pares[:maximo])
    return texto + ("…" if len(pares) > maximo else "")


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

    # 6. El rótulo «Introducción» (paso 4.4): se avisa, y se quita solo si se pide.
    quitado = _quitar_introduccion(doc, opciones, informe, clasificados)

    # 7. Numeración y mayúsculas (paso 4.3): se avisa, y se cambia solo si se pide.
    _texto_de_titulos(doc, opciones, informe, clasificados, [c for c in seguros if c is not quitado])

    # 8. Los títulos de nivel 4 y 5, en la misma línea que su texto y terminados en punto (paso 4.5).
    _titulos_en_linea(informe, clasificados, [c for c in seguros if c is not quitado])

    # 9. El título del trabajo al principio del texto (paso 4.4).
    _titulo_del_trabajo(doc, opciones, informe, clasificados, quitado)
