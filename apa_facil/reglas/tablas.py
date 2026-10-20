"""
Paso 6.1: las tablas.

Según la web oficial de APA (Table setup, Sample tables y Paragraph alignment and
indentation):
- encima, «Tabla 1» en negrita y, en la línea siguiente, el título en cursiva;
  debajo, la nota con «Nota.» en cursiva; lo comparten con las figuras y lo hace
  leyendas.py (con lo que cambia el texto y las decisiones del propietario);
- sin bordes verticales ni alrededor de cada celda: una línea arriba y otra abajo
  de la tabla, otra debajo de los encabezados y otra debajo de los encabezados
  que abarcan varias columnas;
- si ocupa más de una página, la fila de encabezados se repite en cada una;
- una tabla corta, que cabe en una página, no se parte entre dos (decisión del
  propietario, 2026-10-20): todas sus filas menos la última llevan «Conservar con
  el siguiente» y ninguna fila se divide, así que si no cabe al final de una
  página pasa entera a la siguiente;
- si es demasiado ancha, va en una página horizontal (lo avisa el paso 1.1).

El estilo de tabla APA se aplica por defecto: --tablas-tal-cual deja los bordes
y los colores. No se tocan las tablas de la portada o del índice, las de una
sola celda (un recuadro) ni las que llevan imágenes (las figuras son del 6.2).
"""

from docx.oxml import OxmlElement

from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import clasificar
from apa_facil.reglas import leyendas
from apa_facil.reglas.fuente import _runs_propios
from apa_facil.reglas.leyendas import TABLA, _largo, _plural, _xml

ZONAS_SIN_TABLAS_APA = leyendas.ZONAS_SIN_LEYENDAS_APA
NUMERO_O_LETRA = TABLA.numero_o_letra
# Orden de los hijos de w:tblPr, w:tcPr, w:tcBorders y w:trPr (el esquema de Word lo exige).
DESPUES_DE_TBLBORDERS = ("shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
                         "tblPrChange")
DESPUES_DE_TCBORDERS = ("shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
                        "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange")
DESPUES_DE_SHD = DESPUES_DE_TCBORDERS[1:]
DESPUES_DE_TBLHEADER = ("tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange")
DESPUES_DE_CANTSPLIT = ("trHeight", "tblHeader") + DESPUES_DE_TBLHEADER
# Una tabla es corta si cabe holgadamente en una página: unas 25 líneas (el interior va a espacio
# sencillo) contando, en cada fila, la celda con más líneas; una línea de la página tiene unos 85
# caracteres, que se reparten entre las columnas.
LINEAS_DE_TABLA_CORTA = 25
CARACTERES_POR_LINEA = 85
LADOS = ("top", "left", "bottom", "right", "insideH", "insideV")


def _es_tabla_de_datos(tbl):
    """Una tabla de verdad: no un recuadro de una sola celda ni una tabla para colocar imágenes."""
    if tbl.tag != w("tbl"):
        return False
    celdas = sum(1 for _ in tbl.iter(w("tc")))
    imagenes = next(tbl.iter(w("drawing"), w("pict")), None)
    return celdas > 1 and imagenes is None


# --- Encabezados y estilo APA ---------------------------------------------------------------

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


def _lineas(tbl):
    """Cuántas líneas ocupa la tabla, más o menos: en cada fila, las de la celda que más tiene."""
    total = 0
    for tr in _filas(tbl):
        celdas = _celdas(tr)
        ancho = max(CARACTERES_POR_LINEA // max(len(celdas), 1), 1)
        total += max([sum(max(-(-len(texto_parrafo(p)) // ancho), 1) for p in tc.iter(w("p"))) for tc in celdas]
                     or [1])
    return total


def _mantener_entera(tbl):
    """Si la tabla es corta, que Word no la parta entre dos páginas. Devuelve si cambia algo."""
    filas = _filas(tbl)
    if len(filas) < 2 or _lineas(tbl) > LINEAS_DE_TABLA_CORTA:
        return False
    cambiado = False
    for i, tr in enumerate(filas):
        trpr = tr.get_or_add_trPr()
        marca = trpr.find(w("cantSplit"))
        if marca is None:
            _insertar_en_orden(trpr, OxmlElement("w:cantSplit"), DESPUES_DE_CANTSPLIT)
            cambiado = True
        elif marca.get(w("val")) in ("0", "false", "off"):
            marca.attrib.pop(w("val"))
            cambiado = True
        if i == len(filas) - 1:
            continue
        for p in tr.iter(w("p")):
            ppr = p.get_or_add_pPr()
            seguir = ppr.find(w("keepNext"))
            if seguir is None or seguir.get(w("val")) in ("0", "false", "off"):
                ppr.get_or_add_keepNext().attrib.pop(w("val"), None)
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


# --- La regla -------------------------------------------------------------------------------

def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    primeros = getattr(clasificados, "tablas", {})
    tablas = leyendas.buscar(doc, clasificados, TABLA, _es_tabla_de_datos, primeros.get)
    if not tablas:
        return
    leyendas.tratar(doc, opciones, informe, clasificados, tablas, TABLA)
    encabezados = sum(_repetir_encabezados(t.elemento) for t in tablas)
    enteras = sum(_mantener_entera(t.elemento) for t in tablas)
    con_estilo = 0 if opciones.get("tablas_tal_cual") else sum(_estilo_apa(doc, t.elemento) for t in tablas)
    if encabezados:
        informe.cambio(f"Tablas: la fila de encabezados de {_plural(encabezados, 'tabla')} se repetirá en cada "
                       "página si la tabla ocupa más de una, como pide APA.")
    if enteras:
        informe.cambio(f"Tablas: {_plural(enteras, 'tabla corta', 'tablas cortas')} no se partirá{'n' if enteras > 1 else ''} entre "
                       "dos páginas: si no cabe al final de una, pasa entera a la siguiente.")
    if con_estilo:
        informe.cambio(f"Tablas: {_plural(con_estilo, 'tabla')} lleva{'n' if con_estilo > 1 else ''} ahora el "
                       "estilo de APA: solo líneas horizontales (arriba, debajo de los encabezados y abajo), sin "
                       "bordes verticales ni colores de fondo, y con el texto en negro. Si prefieres dejarlas como "
                       "estaban, usa la opción --tablas-tal-cual.")
