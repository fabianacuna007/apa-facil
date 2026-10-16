"""
APA Fácil — fase 1.

Aplica el formato APA 7 de estudiante a una COPIA de un documento Word (.docx):
márgenes, fuente, interlineado, sangría, alineación y número de página.
Nunca cambia el texto: al final compara el texto de antes y el de después
y, si no coinciden, no guarda nada.

Uso:
    python apa.py pruebas/mi_trabajo.docx
    python apa.py pruebas/*.docx --fuente Arial

El resultado se guarda en la carpeta salida/ como mi_trabajo_APA.docx.
"""

import argparse
import sys
from pathlib import Path

import docx
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.part import XmlPart
from docx.shared import Cm
from lxml import etree

# --- Normas APA 7 de la fase 1 ---------------------------------------------

# Fuentes admitidas y su tamaño en puntos.
FUENTES = {
    "Times New Roman": 12,
    "Calibri": 11,
    "Arial": 11,
    "Georgia": 11,
    "Lucida Sans Unicode": 10,
}
MARGEN = Cm(2.54)
SANGRIA = "720"     # 1,27 cm medido en "twips" (veinteavos de punto): 1 cm = 567
DOBLE = "480"       # interlineado en Word: 240 = sencillo, 480 = doble
SENCILLO = "240"

# Fuentes de símbolos (viñetas, iconos). Si se cambian, cambia el carácter que
# se ve en pantalla, así que no se tocan.
FUENTES_SIMBOLO = {"symbol", "wingdings", "wingdings 2", "wingdings 3",
                   "webdings", "marlett", "mt extra"}

# Campos de Word que muestran números de página.
CAMPOS_PAGINA = ("PAGE", "NUMPAGES", "SECTIONPAGES")

# Títulos tras los que empieza la lista de referencias.
TITULOS_REFERENCIAS = {"referencias", "referencias bibliográficas", "bibliografía",
                       "bibliografia", "webgrafía", "webgrafia", "fuentes",
                       "fuentes consultadas", "references"}

CARPETA_SALIDA =Path(__file__).resolve().parent / "salida"

# Orden obligatorio de las propiedades dentro del XML de Word. Si se añade
# una propiedad en mal sitio, Word puede decir que el archivo está dañado.
ORDEN_RPR = [
    "ins", "del", "moveFrom", "moveTo", "rStyle", "rFonts", "b", "bCs", "i",
    "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
    "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
    "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em",
    "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]
ORDEN_PPR = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]


# --- Utilidades para el XML de Word -----------------------------------------

def nombre_local(el):
    return etree.QName(el).localname


def hijo(padre, tag, orden):
    """Devuelve la propiedad `tag` de `padre`; si no existe, la crea en su sitio."""
    el = padre.find(qn("w:" + tag))
    if el is not None:
        return el
    el = OxmlElement("w:" + tag)
    pos = orden.index(tag)
    for i, otro in enumerate(padre):
        n = nombre_local(otro)
        if n in orden and orden.index(n) > pos:
            padre.insert(i, el)
            return el
    padre.append(el)
    return el


def ppr_de(p):
    """Propiedades de párrafo (<w:pPr>), creándolas si faltan."""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    return ppr


def texto_parrafo(p, legible=False):
    """
    El texto de un párrafo, con tabuladores y saltos de línea. Con legible=True
    los campos de Word se muestran con palabras (para el informe).
    """
    partes = []
    for el in p.iter():
        n = nombre_local(el)
        if n == "instrText" and legible:
            campo = (el.text or "").split()
            nombres = {"PAGE": "[nº de página]", "NUMPAGES": "[total de páginas]",
                       "SECTIONPAGES": "[páginas de la sección]"}
            partes.append(nombres.get(campo[0].upper(), "") if campo else "")
        elif n in ("t", "delText", "instrText"):
            partes.append(el.text or "")
        elif n == "tab" and nombre_local(el.getparent()) == "r":
            partes.append("\t")
        elif n in ("br", "cr"):
            partes.append("\n")
        elif n == "noBreakHyphen":
            partes.append("-")
        elif n == "sym":
            partes.append("[" + (el.get(qn("w:char")) or "") + "]")
    return "".join(partes)


def tiene_campo_pagina(p):
    instrucciones = [el.text or "" for el in p.iter(qn("w:instrText"))]
    instrucciones += [el.get(qn("w:instr")) or "" for el in p.iter(qn("w:fldSimple"))]
    palabras = " ".join(instrucciones).upper().split()
    return any(c in palabras for c in CAMPOS_PAGINA)


def tiene_ancestro(el, tags):
    tags = {qn("w:" + t) for t in tags}
    padre = el.getparent()
    while padre is not None:
        if padre.tag in tags:
            return True
        padre = padre.getparent()
    return False


# --- Lectura de estilos (para saber qué formato "hereda" cada párrafo) ------

class Estilos:
    """Recorre la cadena de estilos: párrafo → su estilo → estilo base → ... → valores por defecto."""

    def __init__(self, documento):
        self.raiz = documento.styles.element
        self.por_id = {s.get(qn("w:styleId")): s for s in self.raiz.findall(qn("w:style"))}
        self.parrafo_defecto = next(
            (sid for sid, s in self.por_id.items()
             if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true")),
            None)
        defectos = self.raiz.find(qn("w:docDefaults"))
        self.ppr_defecto = defectos.find(qn("w:pPrDefault") + "/" + qn("w:pPr")) if defectos is not None else None
        self.rpr_defecto = defectos.find(qn("w:rPrDefault") + "/" + qn("w:rPr")) if defectos is not None else None

    def cadena(self, sid):
        vistos = set()
        while sid and sid in self.por_id and sid not in vistos:
            vistos.add(sid)
            estilo = self.por_id[sid]
            yield estilo
            base = estilo.find(qn("w:basedOn"))
            sid = base.get(qn("w:val")) if base is not None else None

    def nombre(self, sid):
        estilo = self.por_id.get(sid)
        if estilo is None:
            return sid or ""
        n = estilo.find(qn("w:name"))
        return n.get(qn("w:val")) if n is not None else sid

    def estilo_de(self, p):
        ps = p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        return ps.get(qn("w:val")) if ps is not None else self.parrafo_defecto

    def ppr(self, p, tag, atributo=None):
        """Valor efectivo de una propiedad de párrafo: directa, del estilo o por defecto."""
        niveles = [p.find(qn("w:pPr"))]
        niveles += [s.find(qn("w:pPr")) for s in self.cadena(self.estilo_de(p))]
        niveles.append(self.ppr_defecto)
        for nivel in niveles:
            if nivel is None:
                continue
            el = nivel.find(qn("w:" + tag))
            if el is None:
                continue
            if atributo is None:
                return el
            valor = el.get(qn("w:" + atributo))
            if valor is not None:
                return valor
        return None

    def rpr(self, r, p, tag):
        """Propiedad efectiva de un fragmento de texto (negrita, tamaño...)."""
        rpr = r.find(qn("w:rPr"))
        niveles = [rpr]
        if rpr is not None and rpr.find(qn("w:rStyle")) is not None:
            sid = rpr.find(qn("w:rStyle")).get(qn("w:val"))
            niveles += [s.find(qn("w:rPr")) for s in self.cadena(sid)]
        niveles += [s.find(qn("w:rPr")) for s in self.cadena(self.estilo_de(p))]
        niveles.append(self.rpr_defecto)
        for nivel in niveles:
            if nivel is not None and nivel.find(qn("w:" + tag)) is not None:
                return nivel.find(qn("w:" + tag))
        return None

    def tamano_base(self):
        """Tamaño (en medios puntos) del texto normal del documento."""
        for nivel in [s.find(qn("w:rPr")) for s in self.cadena(self.parrafo_defecto)] + [self.rpr_defecto]:
            if nivel is not None and nivel.find(qn("w:sz")) is not None:
                return int(nivel.find(qn("w:sz")).get(qn("w:val")))
        return 20  # valor de Word cuando no se indica nada (10 pt)


def activo(el):
    """Para propiedades sí/no como la negrita: <w:b/> o <w:b w:val="true"/>."""
    return el is not None and el.get(qn("w:val")) not in ("0", "false", "off")


# --- Cambios de formato ------------------------------------------------------

def poner_fuente(rpr, fuente, tamano=None):
    """Cambia la fuente (y el tamaño, si se indica) de un bloque <w:rPr>."""
    fuentes = rpr.find(qn("w:rFonts"))
    if fuentes is not None:
        actual = (fuentes.get(qn("w:ascii")) or fuentes.get(qn("w:hAnsi")) or "").lower()
        if actual in FUENTES_SIMBOLO:
            return False
    fuentes = hijo(rpr, "rFonts", ORDEN_RPR)
    for attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
        fuentes.attrib.pop(qn("w:" + attr), None)  # las fuentes del "tema" mandan sobre las normales
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        fuentes.set(qn("w:" + attr), fuente)
    if tamano is not None:
        hijo(rpr, "sz", ORDEN_RPR).set(qn("w:val"), str(tamano * 2))
        hijo(rpr, "szCs", ORDEN_RPR).set(qn("w:val"), str(tamano * 2))
    return True


def poner_interlineado(ppr, linea):
    espaciado = hijo(ppr, "spacing", ORDEN_PPR)
    for attr in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
        espaciado.attrib.pop(qn("w:" + attr), None)
    espaciado.set(qn("w:before"), "0")
    espaciado.set(qn("w:after"), "0")
    espaciado.set(qn("w:line"), linea)
    espaciado.set(qn("w:lineRule"), "auto")


def quitar_justificado(ppr):
    jc = ppr.find(qn("w:jc"))
    if jc is not None and jc.get(qn("w:val")) in ("both", "distribute", "lowKashida",
                                                     "mediumKashida", "highKashida", "thaiDistribute"):
        jc.set(qn("w:val"), "left")
        return True
    return False


def es_estilo_aparte(nombre):
    """Estilos que no son texto del cuerpo: encabezado, pie y notas al pie."""
    n = nombre.lower()
    return any(c in n for c in ("header", "footer", "footnote", "endnote",
                                "annotation", "comment", "balloon"))


def formatear_estilos(estilos, fuente, tamano, informe):
    """Capa 1 y 2: valores por defecto del documento y estilos."""
    raiz = estilos.raiz
    defectos = raiz.find(qn("w:docDefaults"))
    if defectos is None:
        defectos = OxmlElement("w:docDefaults")
        raiz.insert(0, defectos)
    rpr_def = hijo(hijo(defectos, "rPrDefault", ["rPrDefault", "pPrDefault"]), "rPr", ["rPr"])
    poner_fuente(rpr_def, fuente, tamano)
    estilos.rpr_defecto = rpr_def

    for estilo in raiz.findall(qn("w:style")):
        nombre = estilos.nombre(estilo.get(qn("w:styleId")))
        aparte = es_estilo_aparte(nombre)
        rpr = estilo.find(qn("w:rPr"))
        if rpr is not None:
            # Las notas al pie conservan su tamaño (APA lo permite); solo cambia la fuente.
            con_tamano = tamano if (rpr.find(qn("w:sz")) is not None and not aparte) else None
            poner_fuente(rpr, fuente, con_tamano)
        if estilo.get(qn("w:type")) == "paragraph" and not aparte:
            ppr = estilo.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                # el pPr va después de name, basedOn, next, link, ui..., y antes de rPr
                antes_de = estilo.find(qn("w:rPr"))
                if antes_de is not None:
                    antes_de.addprevious(ppr)
                else:
                    estilo.append(ppr)
            poner_interlineado(ppr, DOBLE)
            quitar_justificado(ppr)
    informe["estilos"] = len(raiz.findall(qn("w:style")))


def clasificar(p, estilos, tamano_base):
    """
    Decide si un párrafo es "texto normal" (lleva sangría de primera línea)
    o algo distinto. Devuelve (lleva_sangria, motivo, dudoso).
    """
    texto = texto_parrafo(p).strip()
    if not texto:
        return False, "vacío", False
    if tiene_ancestro(p, ["tc"]):
        return False, "celda de tabla", False
    if tiene_ancestro(p, ["sdtContent"]):
        return False, "índice o bloque especial", False
    if tiene_ancestro(p, ["txbxContent"]):
        return False, "cuadro de texto", False
    if any(True for _ in p.iter(qn("w:drawing"), qn("w:pict"), qn("w:object"))):
        return False, "imagen", False

    nombre = estilos.nombre(estilos.estilo_de(p)).lower()
    for clave in ("heading", "title", "subtitle", "toc", "caption", "título",
                  "titulo", "list", "quote", "cita", "bibliograph"):
        if clave in nombre:
            return False, f"estilo «{estilos.nombre(estilos.estilo_de(p))}»", False
    nivel = estilos.ppr(p, "outlineLvl", "val")
    if nivel is not None and nivel != "9":
        return False, "nivel de título", False
    num = estilos.ppr(p, "numPr")
    if num is not None:
        num_id = num.find(qn("w:numId"))
        if num_id is None or num_id.get(qn("w:val")) != "0":
            return False, "lista", False
    alineacion = estilos.ppr(p, "jc", "val")
    if alineacion in ("center", "right", "end"):
        return False, "centrado" if alineacion == "center" else "alineado a la derecha", False
    izquierda = estilos.ppr(p, "ind", "left") or estilos.ppr(p, "ind", "start") or "0"
    francesa = estilos.ppr(p, "ind", "hanging") or "0"
    if int(francesa) > 0:
        return False, "sangría francesa (¿referencia?)", False
    if int(izquierda) > 0:
        return False, "sangría izquierda (¿cita en bloque?)", False

    if texto.lower().startswith(("nota.", "nota:", "fuente:", "fuentes:")):
        return False, "nota de tabla o figura", False
    if texto[0] in "•●○■□▪◦·*–—-" and len(texto) > 1 and texto[1] in " \t":
        return False, "lista escrita a mano", False

    # Primer elemento con contenido: ¿es un tabulador escrito a mano?
    for el in p.iter(qn("w:t"), qn("w:tab")):
        if nombre_local(el.getparent()) != "r":
            continue
        if el.tag == qn("w:tab"):
            return False, "empieza con tabulador escrito a mano", True
        if (el.text or "").strip():
            break

    palabras = len(texto.split())
    fragmentos = [r for r in p.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()]
    if fragmentos and palabras <= 20:
        if all(activo(estilos.rpr(r, p, "b")) for r in fragmentos):
            return False, "título (línea corta en negrita)", False
        tamanos = [estilos.rpr(r, p, "sz") for r in fragmentos]
        if all(t is not None and int(t.get(qn("w:val"))) > tamano_base for t in tamanos):
            return False, "título (letra más grande)", False
    if palabras <= 12 and texto[-1] not in ".!?:;…»”\")":
        return False, "línea corta sin punto final (¿título o dato?)", True
    return True, "texto", False


def formatear_cuerpo(documento, estilos, fuente, tamano, informe):
    """Capa 3: el formato aplicado a mano en cada párrafo y fragmento del cuerpo."""
    cuerpo = documento.element.body
    tamano_base = estilos.tamano_base()
    parrafos = list(cuerpo.iter(qn("w:p")))

    # Primero se decide todo (con el formato original) y después se cambia.
    decisiones = []
    en_referencias = False
    for p in parrafos:
        sangria, motivo, dudoso = clasificar(p, estilos, tamano_base)
        es_titulo = motivo.startswith(("estilo «Heading", "estilo «Title", "nivel de título", "título"))
        if es_titulo and not tiene_ancestro(p, ["tc", "sdtContent"]):
            # Tras un título "Referencias" o "Bibliografía" empieza la lista de
            # referencias: APA usa ahí sangría francesa (fase 3), no de primera línea.
            en_referencias = texto_parrafo(p).strip().rstrip(":").lower() in TITULOS_REFERENCIAS
        elif en_referencias and (sangria or dudoso):
            sangria, motivo, dudoso = False, "lista de referencias (se formateará en la fase 3)", False
        decisiones.append((p, sangria, motivo, dudoso))

    for p, sangria, motivo, dudoso in decisiones:
        ppr = ppr_de(p)
        poner_interlineado(ppr, DOBLE)
        if quitar_justificado(ppr):
            informe["justificados"] += 1
        if sangria:
            ind = hijo(ppr, "ind", ORDEN_PPR)
            for attr in ("hanging", "hangingChars", "firstLineChars"):
                ind.attrib.pop(qn("w:" + attr), None)
            ind.set(qn("w:firstLine"), SANGRIA)
            informe["con_sangria"] += 1
        elif motivo != "vacío":
            informe["sin_sangria"][motivo] = informe["sin_sangria"].get(motivo, 0) + 1
        elif not (tiene_ancestro(p, ["tc"]) or any(True for _ in p.iter(qn("w:br"), qn("w:sectPr")))):
            informe["vacios"] += 1  # párrafo en blanco usado como espacio
        if dudoso:
            informe["dudas"].append((motivo, texto_parrafo(p).strip()))

    for rpr in cuerpo.iter(qn("w:rPr")):
        if nombre_local(rpr.getparent()) in ("rPrChange", "pPrChange"):
            continue
        con_tamano = tamano if rpr.find(qn("w:sz")) is not None else None
        if not poner_fuente(rpr, fuente, con_tamano):
            informe["simbolos"] += 1
    informe["parrafos"] = len(parrafos)


def poner_margenes(documento, informe):
    for seccion in documento.sections:
        seccion.top_margin = seccion.bottom_margin = MARGEN
        seccion.left_margin = seccion.right_margin = MARGEN
    informe["secciones"] = len(documento.sections)
    informe["apaisadas"] = sum(1 for s in documento.sections if s.page_width > s.page_height)


def parrafo_numero(fuente, tamano):
    """Un párrafo alineado a la derecha con el campo PAGE (el número se calcula solo)."""
    rpr = (f'<w:rPr><w:rFonts w:ascii="{fuente}" w:hAnsi="{fuente}" w:cs="{fuente}" '
           f'w:eastAsia="{fuente}"/><w:sz w:val="{tamano * 2}"/><w:szCs w:val="{tamano * 2}"/></w:rPr>')
    xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="{SENCILLO}" w:lineRule="auto"/>'
        '<w:jc w:val="right"/></w:pPr>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t>1</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p>'
    )
    return parse_xml(xml)


def numerar_encabezado(encabezado, fuente, tamano):
    """Pone el número arriba a la derecha sin borrar lo que ya hubiera en el encabezado."""
    raiz = encabezado._element
    for p in raiz.iter(qn("w:p")):
        if tiene_campo_pagina(p):
            hijo(ppr_de(p), "jc", ORDEN_PPR).set(qn("w:val"), "right")
            return "ya tenía número (alineado a la derecha)"
    parrafos = raiz.findall(qn("w:p"))
    nuevo = parrafo_numero(fuente, tamano)
    if len(parrafos) == 1 and not texto_parrafo(parrafos[0]).strip() and len(parrafos[0]) <= 1:
        parrafos[0].addprevious(nuevo)
        raiz.remove(parrafos[0])
        return "número añadido"
    # Va el primero, para que el número quede en la línea de arriba.
    primero = raiz[0] if len(raiz) else None
    if primero is not None:
        primero.addprevious(nuevo)
    else:
        raiz.append(nuevo)
    return "número añadido encima del texto que ya había"


def poner_numeros_pagina(documento, fuente, tamano, informe):
    """
    Word puede tener hasta tres encabezados por sección: el normal, el de la
    primera página y el de páginas pares. Si una sección no define el suyo,
    usa el de la sección anterior. Se numeran todos los que se llegan a ver.
    """
    pares = documento.settings.odd_and_even_pages_header_footer
    definidos = set()
    for seccion in documento.sections:
        tipos = [("normal", seccion.header, True),
                 ("primera", seccion.first_page_header, seccion.different_first_page_header_footer),
                 ("pares", seccion.even_page_header, pares)]
        for tipo, encabezado, se_usa in tipos:
            if encabezado.is_linked_to_previous:
                if not se_usa or tipo in definidos:
                    continue
                encabezado.is_linked_to_previous = False  # crea un encabezado propio
            definidos.add(tipo)
            informe["encabezados"].append(f"{tipo}: {numerar_encabezado(encabezado, fuente, tamano)}")


def partes(documento, clase_nombre):
    return [p for p in documento.part.package.iter_parts() if type(p).__name__ == clase_nombre]


def formatear_encabezados_y_pies(documento, fuente, tamano, informe):
    """Fuente en encabezados y pies, interlineado sencillo y fuera los números del pie."""
    for parte in partes(documento, "HeaderPart") + partes(documento, "FooterPart"):
        raiz = parte.element
        es_pie = type(parte).__name__ == "FooterPart"
        for p in raiz.iter(qn("w:p")):
            if es_pie and tiene_campo_pagina(p):
                quitado = " ".join(texto_parrafo(p, legible=True).split())
                for el in list(p):
                    if el.tag != qn("w:pPr"):
                        p.remove(el)
                informe["pies_quitados"].append(quitado)
            poner_interlineado(ppr_de(p), SENCILLO)
        for rpr in raiz.iter(qn("w:rPr")):
            poner_fuente(rpr, fuente, tamano if rpr.find(qn("w:sz")) is not None else None)


def formatear_notas(documento, fuente, informe):
    """Notas al pie y al final: misma fuente, tamaño original, interlineado sencillo."""
    for parte in documento.part.package.iter_parts():
        if parte.partname not in ("/word/footnotes.xml", "/word/endnotes.xml"):
            continue
        raiz = parte.element if isinstance(parte, XmlPart) else etree.fromstring(parte.blob)
        notas = [n for n in raiz if n.get(qn("w:type")) in (None, "normal")]
        for p in raiz.iter(qn("w:p")):
            poner_interlineado(ppr_de(p), SENCILLO)
            quitar_justificado(ppr_de(p))
        for rpr in raiz.iter(qn("w:rPr")):
            poner_fuente(rpr, fuente)
        if not isinstance(parte, XmlPart):
            parte._blob = etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)
        informe["notas"] += len(notas)


# --- Comprobación de que el contenido no cambia -----------------------------

def foto(documento):
    """
    Una "foto" del contenido: el texto de cada parte del documento y cuántas
    tablas, imágenes y notas hay. Los párrafos con número de página de
    encabezados y pies se excluyen, porque esos sí cambian a propósito.
    """
    resultado = {}
    for parte in documento.part.package.iter_parts():
        nombre = str(parte.partname)
        es_texto = nombre == "/word/document.xml" or nombre.startswith(
            ("/word/header", "/word/footer", "/word/footnotes", "/word/endnotes"))
        if not es_texto:
            continue
        raiz = parte.element if isinstance(parte, XmlPart) else etree.fromstring(parte.blob)
        excluir_numeros = nombre.startswith(("/word/header", "/word/footer"))
        textos = [texto_parrafo(p) for p in raiz.iter(qn("w:p"))
                  if not (excluir_numeros and tiene_campo_pagina(p))]
        textos = [t for t in textos if t.strip()] if excluir_numeros else textos
        cuentas = {n: sum(1 for _ in raiz.iter(qn("w:" + n)))
                   for n in ("tbl", "drawing", "pict", "footnoteReference", "endnoteReference")}
        if textos or any(cuentas.values()):
            resultado[nombre] = (textos, cuentas)
    resultado["imágenes"] = sum(1 for p in documento.part.package.iter_parts()
                                if type(p).__name__ == "ImagePart")
    return resultado


def comparar(antes, despues):
    problemas = []
    for clave in sorted(set(antes) | set(despues)):
        a, d = antes.get(clave), despues.get(clave)
        if a == d:
            continue
        if clave == "imágenes":
            problemas.append(f"imágenes: {a} antes, {d} después")
            continue
        (ta, ca), (td, cd) = a or ([], {}), d or ([], {})
        if ca != cd:
            problemas.append(f"{clave}: cambió el número de tablas/imágenes/notas ({ca} → {cd})")
        if ta != td:
            problemas.append(f"{clave}: el texto no coincide")
    return problemas


# --- Programa principal -----------------------------------------------------

def aplicar_apa(entrada, fuente="Times New Roman"):
    entrada = Path(entrada).resolve()
    tamano = FUENTES[fuente]
    salida = CARPETA_SALIDA / f"{entrada.stem}_APA.docx"
    if salida == entrada:
        raise ValueError("El archivo de salida sería el mismo que el original.")

    documento = docx.Document(str(entrada))
    antes = foto(documento)
    informe = {"archivo": entrada.name, "salida": salida, "justificados": 0,
               "con_sangria": 0, "sin_sangria": {}, "dudas": [], "simbolos": 0,
               "encabezados": [], "pies_quitados": [], "notas": 0, "vacios": 0}

    estilos = Estilos(documento)
    poner_margenes(documento, informe)
    formatear_estilos(estilos, fuente, tamano, informe)
    formatear_cuerpo(documento, estilos, fuente, tamano, informe)
    formatear_encabezados_y_pies(documento, fuente, tamano, informe)
    poner_numeros_pagina(documento, fuente, tamano, informe)
    formatear_notas(documento, fuente, informe)
    informe["indice"] = any(
        "TOC" in (el.text or "") for el in documento.element.body.iter(qn("w:instrText")))

    problemas = comparar(antes, foto(documento))
    if problemas:
        informe["problemas"] = problemas
        return informe  # no se guarda nada

    CARPETA_SALIDA.mkdir(exist_ok=True)
    documento.save(str(salida))
    # Segunda comprobación, sobre el archivo ya guardado.
    informe["problemas"] = comparar(antes, foto(docx.Document(str(salida))))
    return informe


def imprimir_informe(inf, fuente):
    print("=" * 70)
    print(f"📄 {inf['archivo']}")
    if inf["problemas"]:
        print("  ❌ NO SE HA GUARDADO: el contenido habría cambiado.")
        for p in inf["problemas"]:
            print("     -", p)
        return
    print(f"  ✔ Guardado en salida/{inf['salida'].name}")
    print(f"  ✔ Texto, tablas, imágenes y notas idénticos al original")
    apaisadas = f" ({inf['apaisadas']} apaisada, se respeta)" if inf["apaisadas"] else ""
    print(f"  ✔ Márgenes de 2,54 cm en {inf['secciones']} sección(es){apaisadas}")
    print(f"  ✔ Fuente {fuente} {FUENTES[fuente]} pt e interlineado doble en {inf['parrafos']} párrafos")
    if inf["justificados"]:
        print(f"  ✔ {inf['justificados']} párrafos justificados pasan a alineación izquierda")
    print(f"  ✔ Sangría de 1,27 cm en {inf['con_sangria']} párrafos de texto")
    if inf["sin_sangria"]:
        resumen = ", ".join(f"{n} {m}" for m, n in sorted(inf["sin_sangria"].items(), key=lambda x: -x[1]))
        print(f"    Sin sangría: {resumen}")
    for e in inf["encabezados"]:
        print(f"  ✔ Encabezado {e}")
    for q in inf["pies_quitados"]:
        print(f"  ✔ Quitado del pie de página: «{q}»")
    if inf["notas"]:
        print(f"  ✔ {inf['notas']} nota(s) al pie: fuente cambiada, tamaño original")
    if inf["simbolos"]:
        print(f"  ℹ {inf['simbolos']} fragmento(s) con fuente de símbolos sin tocar")
    if inf["vacios"]:
        print(f"  ⚠ {inf['vacios']} párrafo(s) en blanco usados como espacio: con interlineado doble "
              "dejan huecos grandes. No los borro; quítalos en Word si sobran.")
    if inf["indice"]:
        print("  ⚠ Tiene índice automático: al abrirlo en Word, actualízalo "
              "(clic derecho en el índice → Actualizar campos), porque las páginas cambian.")
    if inf["dudas"]:
        print("  ⚠ Párrafos dudosos (los he dejado SIN sangría; revísalos):")
        for motivo, texto in inf["dudas"]:
            corto = texto.replace("\t", "→").replace("\n", " ⏎ ")
            print(f"     - [{motivo}] «{corto[:70]}{'…' if len(corto) > 70 else ''}»")


def main():
    parser = argparse.ArgumentParser(description="Aplica formato APA 7 de estudiante a una copia de un .docx")
    parser.add_argument("archivos", nargs="+", help="uno o varios archivos .docx")
    parser.add_argument("--fuente", default="Times New Roman",
                        help="Times New Roman (12), Calibri (11), Arial (11), Georgia (11) o Lucida Sans Unicode (10)")
    args = parser.parse_args()

    fuente = next((f for f in FUENTES if f.lower() == args.fuente.lower()), None)
    if fuente is None:
        sys.exit(f"Fuente no admitida en APA 7: {args.fuente}. Usa una de: {', '.join(FUENTES)}")

    errores = 0
    for archivo in args.archivos:
        ruta = Path(archivo)
        if ruta.suffix.lower() != ".docx" or ruta.name.startswith("~$"):
            print(f"Se omite {ruta.name}: solo se admiten archivos .docx")
            continue
        try:
            informe = aplicar_apa(ruta, fuente)
        except Exception as e:  # un archivo dañado no debe parar los demás
            print("=" * 70)
            print(f"📄 {ruta.name}\n  ❌ No se ha podido procesar: {e}")
            errores += 1
            continue
        imprimir_informe(informe, fuente)
        errores += bool(informe["problemas"])
    sys.exit(1 if errores else 0)


if __name__ == "__main__":
    main()
