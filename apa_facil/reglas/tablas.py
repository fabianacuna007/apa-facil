"""
Paso 6.1: las tablas.

Según la web oficial de APA (Table setup, Sample tables y Paragraph alignment and
indentation):
- encima de la tabla, su número («Tabla 1») en negrita y, en la línea siguiente,
  su título en cursiva, con mayúscula inicial en las palabras principales; los
  dos a la izquierda, sin sangría y a doble espacio (el doble espacio ya lo pone
  el paso 1.3);
- debajo, si la hay, la nota: «Nota.» en cursiva y el texto en letra normal, a la
  izquierda y sin sangría;
- sin bordes verticales ni alrededor de cada celda: una línea arriba y otra abajo
  de la tabla, otra debajo de los encabezados y otra debajo de los encabezados
  que abarcan varias columnas;
- si ocupa más de una página, la fila de encabezados se repite en cada una;
- si es demasiado ancha, va en una página horizontal (lo avisa el paso 1.1);
- arriba o abajo de la página, no en medio, y con una línea en blanco entre la
  tabla y el texto: solo se avisa.

Lo que cambia el texto, anotado en el informe y comprobado por la integridad
(decisiones del propietario): la leyenda que está debajo sube encima; el número
y el título que van en la misma línea se separan, y desaparece lo que había
entre ellos («: », «. »); el título lleva las mayúsculas de APA, como los
títulos del paso 4.6 (--titulos-tal-cual las deja); si ninguna tabla tiene
número, se añade «Tabla 1», «Tabla 2»… a las que tienen título (si unas lo
tienen y otras no, solo se avisa, para no descuadrar la numeración); y a la nota
que empieza por «Fuente:» se le pone «Nota. » delante. El estilo de tabla APA se
aplica por defecto: --tablas-tal-cual deja los bordes y los colores.

No se tocan las tablas de la portada o del índice, las de una sola celda (un
recuadro) ni las que llevan imágenes (son figuras: paso 6.2). En un apéndice no
se añade el número: APA las numera con la letra del apéndice (paso 6.7).
"""

import copy
import re
from dataclasses import dataclass, field
from typing import Optional

from docx.oxml import OxmlElement
from lxml import etree

from apa_facil.documento import (LEYENDA_CON_NUMERO, NOTA_SIN_ROTULO, ROTULO_DE_NOTA, Estilos, texto_parrafo,
                                 w)
from apa_facil.estructura import ESTILOS_DE_LEYENDA, NOTA, _bloques, clasificar
from apa_facil.reglas import titulos
from apa_facil.reglas.fuente import _runs_propios
from apa_facil.reglas.resumen import _activado, _partir, _poner

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
ZONAS_SIN_TABLAS_APA = ("portada", "indice")
# Lo que impide separar el número del título sin dudas: cambios marcados, saltos,
# imágenes, notas, comentarios, símbolos.
NO_SE_DIVIDE_SI_LLEVA = ("ins", "del", "delText", "moveFrom", "moveTo", "br", "cr", "drawing", "pict",
                         "footnoteReference", "endnoteReference", "commentReference", "sectPr", "sym",
                         "noBreakHyphen", "txbxContent")
# Orden de los hijos de w:tblPr, w:tcPr, w:tcBorders y w:trPr (el esquema de Word lo exige).
DESPUES_DE_TBLBORDERS = ("shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
                         "tblPrChange")
DESPUES_DE_TCBORDERS = ("shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
                        "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange")
DESPUES_DE_SHD = DESPUES_DE_TCBORDERS[1:]
DESPUES_DE_TBLHEADER = ("tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange")
LADOS = ("top", "left", "bottom", "right", "insideH", "insideV")
NUMERO_O_LETRA = re.compile(r"Tabla [A-Z]?[1-9]\d{0,2}")   # «Tabla 1» o, en un apéndice, «Tabla A1»


@dataclass
class _Tabla:
    elemento: object
    zona: str
    primer_parrafo: Optional[int]   # número de su primer párrafo en el original (para la integridad)
    encima: list = field(default_factory=list)   # Parrafo de las leyendas de encima, en orden
    debajo: list = field(default_factory=list)
    numero: object = None           # el Parrafo con «Tabla N»
    titulo: object = None           # el w:p del título (puede ser nuevo)
    notas: list = field(default_factory=list)


def _extracto(texto, largo=40):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _plural(n, singular, plural=None):
    return f"{n} {singular if n == 1 else plural or singular + 's'}"


def _lista(textos, maximo=3):
    return ", ".join(f"«{_extracto(t)}»" for t in textos[:maximo]) + ("…" if len(textos) > maximo else "")


def _texto(c):
    return " ".join(texto_parrafo(c.elemento).split())


# --- Qué tablas hay y cuáles son sus leyendas ---------------------------------------------

def _es_tabla_de_datos(tbl):
    """Una tabla de verdad: no un recuadro de una sola celda ni una tabla para colocar imágenes."""
    celdas = sum(1 for _ in tbl.iter(w("tc")))
    imagenes = next(tbl.iter(w("drawing"), w("pict")), None)
    return celdas > 1 and imagenes is None


def _sigue_una_tabla(bloques, j, de):
    """¿Va el bloque j justo antes de una tabla (solo con párrafos vacíos en medio)?"""
    for b in bloques[j + 1:]:
        if b.tag == w("tbl"):
            return True
        c = de.get(b)
        if c is None or c.tipo != "vacio":
            return False
    return False


def _leyendas(bloques, k, paso, de, usados, maximo):
    """Las leyendas de tabla seguidas encima (paso -1) o debajo (+1) del bloque k, sin contar los vacíos."""
    encontradas, j = [], k + paso
    while 0 <= j < len(bloques) and len(encontradas) < maximo:
        c = de.get(bloques[j])
        if c is None or c.elemento in usados or (c.tipo != "vacio" and c.tipo != "leyenda_tabla"):
            break
        if c.tipo == "leyenda_tabla":
            # Una leyenda con número que va justo antes de otra tabla es de esa otra tabla.
            if paso > 0 and LEYENDA_CON_NUMERO.match(_texto(c)) and _sigue_una_tabla(bloques, j, de):
                break
            encontradas.append(c)
        j += paso
    return encontradas if paso > 0 else encontradas[::-1]


def _tablas(doc, clasificados):
    de = {c.elemento: c for c in clasificados if c.elemento.getparent() is not None}
    primeros = getattr(clasificados, "tablas", {})
    bloques = _bloques(doc)
    resultado, usados = [], set()
    for k, tbl in enumerate(bloques):
        if tbl.tag != w("tbl") or not _es_tabla_de_datos(tbl):
            continue
        vecinos = [de[b] for b in reversed(bloques[:k]) if b in de] + [de[b] for b in bloques[k + 1:] if b in de]
        zona = vecinos[0].zona if vecinos else "cuerpo"
        if zona in ZONAS_SIN_TABLAS_APA:
            continue
        t = _Tabla(tbl, zona, primeros.get(tbl))
        t.encima = _leyendas(bloques, k, -1, de, usados, 2)
        usados.update(c.elemento for c in t.encima)
        t.debajo = _leyendas(bloques, k, +1, de, usados, 3)
        usados.update(c.elemento for c in t.debajo)
        resultado.append(t)
    return resultado


# --- Separar, mover y añadir ---------------------------------------------------------------

def _largo(r):
    return sum(len(el.text or "") if el.tag == w("t") else 1 for el in r if el.tag in (w("t"), w("tab")))


def _posiciones(p):
    """Cada fragmento del párrafo con la posición (en su texto) en que empieza."""
    resultado, posicion = [], 0
    for r in _runs_propios(p):
        resultado.append((r, posicion))
        posicion += _largo(r)
    return resultado


def _se_puede_dividir(p, corte):
    """¿Se puede partir el párrafo en `corte` sin partir un campo de Word (p. ej. la numeración automática)?"""
    if any(next(p.iter(w(e)), None) is not None for e in NO_SE_DIVIDE_SI_LLEVA):
        return False
    abiertos = []
    for r, posicion in _posiciones(p):
        for marca in r.iter(w("fldChar")):
            if marca.get(w("fldCharType")) == "begin":
                abiertos.append(posicion)
            elif marca.get(w("fldCharType")) == "end" and abiertos and (abiertos.pop() < corte) != (posicion < corte):
                return False
    for simple in p.iter(w("fldSimple")):
        dentro = [posicion for r, posicion in _posiciones(p) if next(r.iterancestors(w("fldSimple")), None) is simple]
        if dentro and (min(dentro) < corte) != (max(dentro) < corte):
            return False
    return not abiertos


def _quedarse_con(p, desde, hasta, corte):
    """
    Deja en el párrafo solo el texto entre las posiciones `desde` y `hasta`. Los
    fragmentos sin texto (marcas de campo) se quedan en la parte en la que están
    respecto a `corte`. Se quitan los fragmentos que se quedan vacíos.
    """
    primera = desde == 0
    for r, inicio in _posiciones(p):
        largo = _largo(r)
        if not largo:
            if (inicio < corte) != primera:
                r.getparent().remove(r)
            continue
        posicion = inicio
        for el in [el for el in r if el.tag in (w("t"), w("tab"))]:
            if el.tag == w("tab"):
                if not desde <= posicion < hasta:
                    r.remove(el)
                posicion += 1
                continue
            texto = el.text or ""
            el.text = texto[max(0, desde - posicion):max(0, hasta - posicion)]
            el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            posicion += len(texto)
        if not _largo(r):
            r.getparent().remove(r)
    for contenedor in list(p.iter(w("hyperlink"), w("fldSimple"), w("smartTag"))):
        if next(contenedor.iter(w("r")), None) is None:
            contenedor.getparent().remove(contenedor)


def _dividir(p, fin, desde):
    """Parte la leyenda en dos párrafos: el número (hasta `fin`) y el título (desde `desde`). Devuelve el del título."""
    titulo = copy.deepcopy(p)
    for marca in list(titulo.iter(w("bookmarkStart"), w("bookmarkEnd"))):   # no se repiten
        marca.getparent().remove(marca)
    for atributo in [a for a in titulo.attrib if a.startswith(f"{{{W14}}}")]:   # identificadores de párrafo
        del titulo.attrib[atributo]
    _quedarse_con(p, 0, fin, desde)
    _quedarse_con(titulo, desde, len(texto_parrafo(titulo)), desde)
    p.addnext(titulo)
    return titulo


def _parrafo_nuevo(modelo, texto):
    """Un párrafo con `texto`, con el formato de párrafo de `modelo` y el de su primer fragmento."""
    nuevo = OxmlElement("w:p")
    ppr = modelo.find(w("pPr"))
    if ppr is not None:
        nuevo.append(copy.deepcopy(ppr))
    nuevo.append(_fragmento_como(modelo, texto))
    return nuevo


def _fragmento_como(p, texto):
    """Un fragmento nuevo con `texto` y el formato del primer fragmento con texto de `p`."""
    r = OxmlElement("w:r")
    modelo = next((x for x in _runs_propios(p) if _largo(x)), None)
    if modelo is not None and modelo.find(w("rPr")) is not None:
        r.append(copy.deepcopy(modelo.find(w("rPr"))))
    t = OxmlElement("w:t")
    t.text = texto
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def _anadir_rotulo_de_nota(p):
    """Pone «Nota. » al principio del párrafo, en un fragmento nuevo."""
    r = _fragmento_como(p, ROTULO_DE_NOTA)
    primero = next(iter(_runs_propios(p)), None)
    if primero is not None:
        primero.addprevious(r)
    else:
        p.append(r)


# --- Formato --------------------------------------------------------------------------------

def _formato_de_parrafo(p, conservar_con_el_siguiente):
    """A la izquierda, sin sangría y, el número y el título, pegados a lo que sigue. Devuelve si cambia algo."""
    antes = copy.deepcopy(p.find(w("pPr")))
    ppr = p.get_or_add_pPr()
    ppr.get_or_add_jc().set(w("val"), "left")
    ind = ppr.get_or_add_ind()
    for atributo in ("hanging", "hangingChars", "firstLineChars", "leftChars", "startChars", "start"):
        ind.attrib.pop(w(atributo), None)
    ind.set(w("left"), "0")
    ind.set(w("firstLine"), "0")
    if conservar_con_el_siguiente:
        ppr.get_or_add_keepNext().attrib.pop(w("val"), None)
    return antes is None or _xml(antes) != _xml(ppr)


def _xml(el):
    return etree.tostring(el)


def _formato_de_texto(doc, p, cursiva_hasta, negrita):
    """
    Negrita (`negrita`) o no en todo el párrafo, y cursiva solo hasta la posición
    `cursiva_hasta` (0: nada; None: todo). Devuelve si cambia algo.
    """
    estilos = Estilos(doc)
    if cursiva_hasta:
        for r, inicio in _posiciones(p):
            if inicio < cursiva_hasta < inicio + _largo(r):
                _partir(r, cursiva_hasta - inicio)
                break
    cambiado = False
    for r, inicio in _posiciones(p):
        if not "".join(t.text or "" for t in r.iter(w("t"))).strip():
            continue
        cursiva = cursiva_hasta is None or inicio < cursiva_hasta
        for etiqueta, quiere in (("i", cursiva), ("b", negrita)):
            if _activado(estilos, r, p, etiqueta) != quiere:
                _poner(r, etiqueta, quiere)
                cambiado = True
    return cambiado


def _sin_color_de_estilo(doc, p):
    """Quita el color del estilo de leyenda de Word («Descripción»), que por defecto es gris azulado."""
    estilos = Estilos(doc)
    for estilo in estilos.cadena(estilos.estilo_de(p)):
        nombre = estilos.nombre(estilo.get(w("styleId"))).lower()
        if any(c in nombre for c in ESTILOS_DE_LEYENDA):
            for color in list(estilo.iter(w("color"))):
                color.getparent().remove(color)


def _insertar_en_orden(padre, nuevo, siguientes):
    """Pone `nuevo` en `padre` delante del primero de los hijos `siguientes` que haya (o al final)."""
    for hijo in padre:
        if hijo.tag in [w(s) for s in siguientes]:
            hijo.addprevious(nuevo)
            return
    padre.append(nuevo)


def _borde(lado, linea):
    el = OxmlElement(f"w:{lado}")
    if linea:
        for atributo, valor in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", "auto")):
            el.set(w(atributo), valor)
    else:
        el.set(w("val"), "nil")
    return el


def _filas(tbl):
    return [tr for tr in tbl.iter(w("tr")) if next(tr.iterancestors(w("tbl")), None) is tbl]


def _celdas(tr):
    return [tc for tc in tr.iter(w("tc")) if next(tc.iterancestors(w("tr")), None) is tr]


def _continua_de_arriba(tc):
    vmerge = tc.find(f"{w('tcPr')}/{w('vMerge')}")
    return vmerge is not None and vmerge.get(w("val")) in (None, "continue")


def _filas_de_encabezado(tbl):
    """La primera fila y las que la continúan (celdas combinadas en vertical con las de arriba)."""
    filas = _filas(tbl)
    n = 1
    while n < len(filas) - 1 and any(_continua_de_arriba(tc) for tc in _celdas(filas[n])):
        n += 1
    return filas[:n]


def _repetir_encabezados(tbl):
    """Marca las filas de encabezado para que se repitan en cada página. Devuelve si cambia algo."""
    cambiado = False
    for tr in _filas_de_encabezado(tbl):
        trpr = tr.get_or_add_trPr()
        marca = trpr.find(w("tblHeader"))
        if marca is None:
            marca = OxmlElement("w:tblHeader")
            _insertar_en_orden(trpr, marca, DESPUES_DE_TBLHEADER)
            cambiado = True
        elif marca.get(w("val")) in ("0", "false", "off"):
            marca.attrib.pop(w("val"))
            cambiado = True
    return cambiado


def _colores_del_estilo(doc, tbl):
    """¿Pone color al texto el estilo de tabla de Word que usa la tabla?"""
    estilo = tbl.find(f"{w('tblPr')}/{w('tblStyle')}")
    if estilo is None:
        return False
    estilos = Estilos(doc)
    return any(c.get(w("val")) not in (None, "auto", "000000")
               for e in estilos.cadena(estilo.get(w("val"))) for c in e.iter(w("color")))


def _estilo_apa(doc, tbl):
    """Solo líneas horizontales (arriba, debajo de los encabezados y abajo), sin fondos y con el texto sin color."""
    tblpr = tbl.find(w("tblPr"))
    antes = _xml(tbl)
    for viejo in tblpr.findall(w("tblBorders")) + tblpr.findall(w("shd")):
        tblpr.remove(viejo)
    bordes = OxmlElement("w:tblBorders")
    for lado in LADOS:
        bordes.append(_borde(lado, lado in ("top", "bottom")))
    _insertar_en_orden(tblpr, bordes, DESPUES_DE_TBLBORDERS)

    filas = _filas(tbl)
    encabezado = _filas_de_encabezado(tbl)
    for numero, tr in enumerate(filas):
        for tc in _celdas(tr):
            tcpr = tc.get_or_add_tcPr()
            for viejo in tcpr.findall(w("tcBorders")) + tcpr.findall(w("shd")):
                tcpr.remove(viejo)
            abarca = tcpr.find(w("gridSpan")) is not None and int(tcpr.find(w("gridSpan")).get(w("val"), "1")) > 1
            debajo = (tr is encabezado[-1] or numero == len(filas) - 1
                      or (tr in encabezado and abarca))   # encabezado que abarca varias columnas
            celda = OxmlElement("w:tcBorders")
            for lado in ("top", "left", "bottom", "right"):
                celda.append(_borde(lado, (lado == "top" and numero == 0) or (lado == "bottom" and debajo)))
            _insertar_en_orden(tcpr, celda, DESPUES_DE_TCBORDERS)
            fondo = OxmlElement("w:shd")
            for atributo, valor in (("val", "clear"), ("color", "auto"), ("fill", "auto")):
                fondo.set(w(atributo), valor)
            _insertar_en_orden(tcpr, fondo, DESPUES_DE_SHD)

    poner_negro = _colores_del_estilo(doc, tbl)
    for p in tbl.iter(w("p")):
        ppr = p.find(w("pPr"))
        for fondo in ppr.findall(w("shd")) if ppr is not None else []:
            ppr.remove(fondo)
        for r in _runs_propios(p):
            if next(r.iterancestors(w("hyperlink")), None) is not None:
                continue
            rpr = r.find(w("rPr"))
            for marca in rpr.findall(w("color")) + rpr.findall(w("shd")) + rpr.findall(w("highlight")) \
                    if rpr is not None else []:
                rpr.remove(marca)
            if poner_negro and _largo(r):
                r.get_or_add_rPr().get_or_add_color().set(w("val"), "auto")
    return _xml(tbl) != antes


# --- Avisos ---------------------------------------------------------------------------------

def _sin_linea_en_blanco(doc, tablas):
    """Las tablas con texto pegado encima o debajo (sin un párrafo vacío en medio)."""
    bloques = _bloques(doc)
    posicion = {b: i for i, b in enumerate(bloques)}

    def con_texto(i):
        return 0 <= i < len(bloques) and bloques[i].tag == w("p") and texto_parrafo(bloques[i]).strip() != ""

    pegadas = []
    for n, t in enumerate(tablas, start=1):
        arriba = [x for x in (t.numero.elemento if t.numero else None, t.titulo) if x is not None]
        primero = arriba[0] if arriba else t.elemento
        ultimo = t.notas[-1].elemento if t.notas else t.elemento
        if con_texto(posicion[primero] - 1) or con_texto(posicion[ultimo] + 1):
            pegadas.append(_nombre(t, n))
    return pegadas


def _nombre(t, n):
    return _texto(t.numero).split(":")[0].split(".")[0] if t.numero else f"la tabla {n}"


# --- La regla -------------------------------------------------------------------------------

def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    tablas = _tablas(doc, clasificados)
    if not tablas:
        return
    tal_cual, titulos_tal_cual = opciones.get("tablas_tal_cual"), opciones.get("titulos_tal_cual")
    subidas, separadas, numeradas, con_nota, mayusculas, dudosas = [], [], [], [], [], []
    sin_mayusculas, sin_titulo, sin_mover, sin_dividir, notas_encima, cuadros = [], [], [], [], [], []
    con_formato = encabezados = con_estilo = 0

    for t in tablas:
        for c in t.encima + t.debajo:
            if NOTA.match(_texto(c)):
                (t.notas if c in t.debajo else notas_encima).append(c)
        t.numero = next((c for c in t.encima + t.debajo if LEYENDA_CON_NUMERO.match(_texto(c))), None)
        lista = t.encima if t.numero in t.encima or t.numero is None else t.debajo
        resto = [c for c in lista if c is not t.numero and c not in t.notas and c not in notas_encima]
        if t.numero is None:
            titulo = resto[-1] if resto and lista is t.encima else None
        else:
            despues = lista[lista.index(t.numero) + 1:]
            titulo = next((c for c in despues if c in resto), None)
        t.titulo = titulo.elemento if titulo is not None else None

        # 1. La leyenda que está debajo sube encima (con su título, si va en otro párrafo).
        if t.numero is not None and t.numero in t.debajo:
            if t.primer_parrafo is None:
                sin_mover.append(_texto(t.numero))
            else:
                for c in [t.numero] + ([titulo] if titulo is not None else []):
                    t.elemento.addprevious(c.elemento)
                    informe.cambio_previsto(c.numero, "mover_leyenda", desde=(t.primer_parrafo,))
                subidas.append(_texto(t.numero))

        # 2. El número y el título en la misma línea se separan.
        if t.numero is not None:
            p = t.numero.elemento
            texto = texto_parrafo(p)
            leyenda = LEYENDA_CON_NUMERO.match(texto)
            if not leyenda.group(1).startswith("Tabla "):
                cuadros.append(leyenda.group(1))
            if leyenda.group(3):
                if _se_puede_dividir(p, leyenda.end(2)) and t.titulo is None:
                    t.titulo = _dividir(p, leyenda.end(1), leyenda.end(2))
                    if not titulos_tal_cual:
                        dudosas += titulos._poner_mayusculas(
                            t.titulo, lambda x: titulos.mayusculas_apa(x, clasificados)) or []
                    informe.cambio_previsto(t.numero.numero, "dividir_leyenda", quitado=leyenda.group(2),
                                            nuevo=texto_parrafo(t.titulo))
                    separadas.append((texto, leyenda.group(1), leyenda.group(3), texto_parrafo(t.titulo)))
                else:
                    sin_dividir.append(texto)
            elif titulo is None:
                sin_titulo.append(_texto(t.numero))

        # 3. Las mayúsculas de APA en el título que ya iba en su propio párrafo.
        if titulo is not None:
            antes = texto_parrafo(titulo.elemento)
            if titulos_tal_cual:
                if titulos.mayusculas_apa(antes, clasificados)[0] != antes:
                    sin_mayusculas.append(antes)
            else:
                dudosas += titulos._poner_mayusculas(
                    titulo.elemento, lambda x: titulos.mayusculas_apa(x, clasificados)) or []
                if texto_parrafo(titulo.elemento) != antes:
                    informe.cambio_previsto(titulo.numero, "mayusculas_apa", nuevo=texto_parrafo(titulo.elemento))
                    mayusculas.append((antes, texto_parrafo(titulo.elemento)))

        # 4. «Nota.» delante de la nota que empieza por «Fuente:».
        for c in t.notas:
            if NOTA_SIN_ROTULO.match(texto_parrafo(c.elemento)):
                _anadir_rotulo_de_nota(c.elemento)
                informe.cambio_previsto(c.numero, "añadir_nota", nuevo=ROTULO_DE_NOTA)
                con_nota.append(texto_parrafo(c.elemento))

    # 5. El número de las tablas que no lo tienen, si ninguna lo tiene (en un apéndice, no: paso 6.7).
    sin_numero = [t for t in tablas if t.numero is None and t.titulo is not None]
    sin_nada = [n for n, t in enumerate(tablas, start=1) if t.numero is None and t.titulo is None]
    if sin_numero and not any(t.numero for t in tablas):
        siguiente = 1
        for t in tablas:
            if t.numero is None and t.titulo is not None and t.zona != "anexos":
                titulo = next(c for c in t.encima if c.elemento is t.titulo)
                nuevo = _parrafo_nuevo(t.titulo, f"Tabla {siguiente}")
                t.titulo.addprevious(nuevo)
                informe.cambio_previsto(titulo.numero, "añadir_numero_de_tabla", nuevo=f"Tabla {siguiente}")
                numeradas.append((f"Tabla {siguiente}", texto_parrafo(t.titulo)))
                t.numero = _NumeroNuevo(nuevo)
                siguiente += 1

    # 6. El formato de las leyendas y de las notas, y la tabla.
    for t in tablas:
        cambiado = False
        if t.numero is not None:
            p = t.numero.elemento
            _sin_color_de_estilo(doc, p)
            cambiado |= _formato_de_parrafo(p, True)
            cambiado |= _formato_de_texto(doc, p, 0, True)
        if t.titulo is not None:
            _sin_color_de_estilo(doc, t.titulo)
            cambiado |= _formato_de_parrafo(t.titulo, True)
            cambiado |= _formato_de_texto(doc, t.titulo, None, False)
        for c in t.notas:
            rotulo = NOTA.match(texto_parrafo(c.elemento))
            _sin_color_de_estilo(doc, c.elemento)
            cambiado |= _formato_de_parrafo(c.elemento, False)
            cambiado |= _formato_de_texto(doc, c.elemento, rotulo.end(), False)
        con_formato += cambiado
        encabezados += _repetir_encabezados(t.elemento)
        if not tal_cual:
            con_estilo += _estilo_apa(doc, t.elemento)

    _informar(informe, tablas, subidas, separadas, numeradas, con_nota, mayusculas, dudosas, con_formato,
              encabezados, con_estilo)
    _avisar(doc, informe, opciones, tablas, sin_mayusculas, sin_titulo, sin_mover, sin_dividir, notas_encima,
            cuadros, sin_numero, sin_nada)


@dataclass
class _NumeroNuevo:
    """El párrafo «Tabla N» añadido (hace las veces de un Parrafo clasificado)."""
    elemento: object


def _informar(informe, tablas, subidas, separadas, numeradas, con_nota, mayusculas, dudosas, con_formato,
              encabezados, con_estilo):
    if subidas:
        informe.cambio(f"Tablas: {'la leyenda' if len(subidas) == 1 else 'las leyendas'} {_lista(subidas)} "
                       f"estaba{'n' if len(subidas) > 1 else ''} debajo de la tabla y ahora va{'n' if len(subidas) > 1 else ''} "
                       "encima, como pide APA.")
    if separadas:
        texto, numero, _, titulo = separadas[0]
        con_mayusculas = any(antes != despues for _, _, antes, despues in separadas)
        informe.cambio(f"Tablas: en {_plural(len(separadas), 'leyenda')} se ha separado el número del título, que "
                       f"ahora va en la línea siguiente («{_extracto(texto)}» → «{numero}» y, debajo, "
                       f"«{_extracto(titulo)}»), como pide APA."
                       + (" El título lleva además mayúscula inicial en las palabras principales (si prefieres "
                          "dejarlo como estaba, usa la opción --titulos-tal-cual)." if con_mayusculas else ""))
    if numeradas:
        informe.cambio(f"Tablas: se ha añadido el número a {_plural(len(numeradas), 'tabla')} que no lo "
                       f"tenía{'n' if len(numeradas) > 1 else ''} («{numeradas[0][0]}» encima de "
                       f"«{_extracto(numeradas[0][1])}»). Comprueba que el texto se refiere a "
                       f"{'ella' if len(numeradas) == 1 else 'ellas'} con ese número: APA numera las tablas en el "
                       "orden en que se mencionan.")
    if con_nota:
        informe.cambio(f"Tablas: {'la nota' if len(con_nota) == 1 else 'las notas'} de debajo de la tabla "
                       f"empieza{'n' if len(con_nota) > 1 else ''} ahora por «Nota.», como pide APA "
                       f"({_lista(con_nota)}).")
    if mayusculas:
        informe.cambio(f"Tablas: {_plural(len(mayusculas), 'título')} de tabla lleva{'n' if len(mayusculas) > 1 else ''} "
                       "ahora mayúscula inicial en las palabras principales, como pide APA ("
                       + ", ".join(f"«{_extracto(a)}» → «{_extracto(d)}»" for a, d in mayusculas[:3])
                       + "). Si prefieres dejarlos como estaban, usa la opción --titulos-tal-cual.")
    titulos.avisar_dudosas(informe, dudosas, "los títulos de las tablas")
    if con_formato:
        informe.cambio(f"Tablas: el número y el título de {_plural(con_formato, 'tabla')} van ahora a la izquierda y "
                       "sin sangría, el número en negrita y el título en cursiva; en las notas, solo «Nota.» va en "
                       "cursiva.")
    if encabezados:
        informe.cambio(f"Tablas: la fila de encabezados de {_plural(encabezados, 'tabla')} se repetirá en cada "
                       "página si la tabla ocupa más de una, como pide APA.")
    if con_estilo:
        informe.cambio(f"Tablas: {_plural(con_estilo, 'tabla')} lleva{'n' if con_estilo > 1 else ''} ahora el "
                       "estilo de APA: solo líneas horizontales (arriba, debajo de los encabezados y abajo), sin "
                       "bordes verticales, sin colores de fondo y con el texto en negro. Si prefieres dejarlas como "
                       "estaban, usa la opción --tablas-tal-cual.")


def _avisar(doc, informe, opciones, tablas, sin_mayusculas, sin_titulo, sin_mover, sin_dividir, notas_encima,
            cuadros, sin_numero, sin_nada):
    if sin_mayusculas:
        informe.aviso(f"{_plural(len(sin_mayusculas), 'título')} de tabla no lleva{'n' if len(sin_mayusculas) > 1 else ''} "
                      f"las mayúsculas de APA ({_lista(sin_mayusculas)}). Se ha dejado como estaba porque has usado "
                      "la opción --titulos-tal-cual.")
    if sin_numero and any(t.numero for t in tablas if not isinstance(t.numero, _NumeroNuevo)):
        informe.aviso(f"{_plural(len(sin_numero), 'tabla')} no lleva{'n' if len(sin_numero) > 1 else ''} número "
                      f"({_lista([texto_parrafo(t.titulo) for t in sin_numero])}). Como otras sí lo llevan, no se ha "
                      "añadido, para no descuadrar la numeración: ponlo tú encima del título («Tabla 2»…).")
    anexos = [t for t in sin_numero if t.zona == "anexos" and t.numero is None]
    if anexos and not any(t.numero for t in tablas if not isinstance(t.numero, _NumeroNuevo)):
        informe.aviso(f"{_plural(len(anexos), 'tabla')} de un apéndice no lleva{'n' if len(anexos) > 1 else ''} "
                      "número: APA las numera con la letra del apéndice («Tabla A1»). Ponlo tú.")
    if sin_nada:
        informe.aviso(f"{'La tabla ' + str(sin_nada[0]) if len(sin_nada) == 1 else 'Las tablas ' + ', '.join(map(str, sin_nada))} "
                      f"no lleva{'n' if len(sin_nada) > 1 else ''} número ni título. APA pone encima de cada tabla "
                      "«Tabla N» en negrita y, debajo, su título en cursiva: escríbelos tú.")
    if sin_titulo:
        informe.aviso(f"{_lista(sin_titulo)} no tiene{'n' if len(sin_titulo) > 1 else ''} título: APA lo pone en la "
                      "línea siguiente al número, en cursiva. Escríbelo tú.")
    if sin_dividir:
        informe.aviso(f"En {_lista(sin_dividir)} el número y el título van en la misma línea, pero la leyenda lleva "
                      "un campo de Word o cambios marcados y no se ha podido separar: pon tú el título en la línea "
                      "siguiente.")
    if sin_mover:
        informe.aviso(f"La leyenda {_lista(sin_mover)} está debajo de la tabla y no se ha podido subir: APA la pone "
                      "encima. Muévela tú.")
    if notas_encima:
        informe.aviso(f"La nota {_lista([_texto(c) for c in notas_encima])} está encima de la tabla: APA pone las "
                      "notas debajo. Muévela tú.")
    if cuadros:
        informe.aviso(f"APA llama «Tabla» a las tablas, y no «{cuadros[0].split()[0]}» ({_lista(cuadros)}): "
                      "cámbialo tú si tu universidad no pide otra cosa.")
    numeros = [int(m.group(1)) for t in tablas if t.numero is not None and t.zona != "anexos"
               for m in [re.search(r"\s(\d+)$", LEYENDA_CON_NUMERO.match(texto_parrafo(t.numero.elemento)).group(1))]
               if m]
    if numeros and numeros != list(range(1, len(numeros) + 1)):
        informe.aviso(f"La numeración de las tablas no va seguida ({', '.join(map(str, numeros))}): APA las numera "
                      "1, 2, 3… en el orden en que se mencionan en el texto. Revísala tú.")
    pegadas = _sin_linea_en_blanco(doc, tablas)
    if pegadas:
        informe.aviso(f"{_lista(pegadas)} no tiene{'n' if len(pegadas) > 1 else ''} una línea en blanco entre la tabla "
                      "y el texto: APA la recomienda para separarlas. Añádela tú con un párrafo vacío encima del "
                      "número y debajo de la tabla (o de su nota).")
    informe.aviso("Comprueba en Word que cada tabla queda arriba o abajo de la página, y no en medio del texto, como "
                  "recomienda APA (también puede ir en una página para ella sola). La herramienta no puede saberlo: "
                  "depende de cómo reparte Word las páginas.")
