"""
Abrir y guardar documentos Word, y utilidades para trabajar con su XML.

Todo lo que python-docx no cubre (campos de número de página, notas al pie,
algunos ajustes de sección) se hace aquí con lxml, para que el resto del
programa no tenga que tocar el XML directamente.
"""

import difflib
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml import OxmlElement, parse_xml
from lxml import etree

CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "salida"
SUFIJO = "_APA"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(etiqueta):
    """Nombre completo de una etiqueta del XML de Word: w("p") → {…/main}p."""
    return f"{{{W}}}{etiqueta}"


class ErrorDocumento(Exception):
    """Un problema con el archivo, explicado en español para el estudiante."""


def abrir(ruta):
    """Abre un documento Word. Si no se puede, explica por qué con un ErrorDocumento."""
    ruta = Path(ruta)
    if not ruta.exists():
        raise ErrorDocumento(f"No se encuentra el archivo «{ruta.name}». Comprueba el nombre y la carpeta.")
    if ruta.is_dir():
        raise ErrorDocumento(f"«{ruta.name}» es una carpeta, no un documento.")
    if ruta.name.startswith("~$"):
        raise ErrorDocumento(f"«{ruta.name}» es un archivo temporal que Word crea mientras "
                             "un documento está abierto. No es tu trabajo: elige el otro archivo.")
    extension = ruta.suffix.lower()
    if extension == ".doc":
        raise ErrorDocumento(f"«{ruta.name}» está en el formato antiguo de Word (.doc). Ábrelo en Word "
                             "y usa «Guardar como» → «Documento de Word (.docx)».")
    if extension != ".docx":
        raise ErrorDocumento(f"«{ruta.name}» no es un documento de Word. Solo se admiten archivos .docx.")

    # Un .docx es en realidad un archivo ZIP con varios archivos XML dentro.
    try:
        if not zipfile.is_zipfile(ruta):
            raise ErrorDocumento(f"«{ruta.name}» está dañado o no es un documento de Word de verdad "
                                 "(aunque termine en .docx).")
        return docx.Document(str(ruta))
    except PermissionError as e:
        raise ErrorDocumento(f"No se puede leer «{ruta.name}». Puede que otro programa lo tenga "
                             "bloqueado: ciérralo y vuelve a intentarlo.") from e
    except (PackageNotFoundError, KeyError, ValueError, zipfile.BadZipFile, etree.XMLSyntaxError) as e:
        raise ErrorDocumento(f"«{ruta.name}» está dañado o no es un documento de Word de verdad "
                             "(aunque termine en .docx).") from e


def ruta_salida(entrada, carpeta=None):
    """Dónde se guarda la copia: salida/nombre_APA.docx. Nunca puede ser el original."""
    entrada = Path(entrada).resolve()
    carpeta = Path(carpeta) if carpeta is not None else CARPETA_SALIDA
    salida = (carpeta / f"{entrada.stem}{SUFIJO}.docx").resolve()
    if salida == entrada:
        raise ErrorDocumento(f"La copia de «{entrada.name}» sustituiría al original. No se guarda nada.")
    return salida


def guardar(documento, ruta):
    """Guarda el documento en `ruta`, creando la carpeta si hace falta."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    try:
        documento.save(str(ruta))
    except PermissionError as e:
        raise ErrorDocumento(f"No se puede guardar «{ruta.name}». Seguramente lo tienes abierto "
                             "en Word: ciérralo y vuelve a intentarlo.") from e


# --- Prueba de integridad ----------------------------------------------------
#
# Una "radiografía" recoge el texto del documento, en orden, y cuántos
# elementos tiene. Comparando la del original con la de la copia se comprueba
# que solo ha cambiado el formato.
#
# Los encabezados y pies de página quedan fuera a propósito: la regla 1.6 los
# cambia (quita números de página y texto) y lo registra en el informe.

@dataclass
class Radiografia:
    textos: list = field(default_factory=list)   # (ubicación, texto) de cada párrafo
    tablas: int = 0
    imagenes: int = 0
    notas_al_pie: int = 0
    notas_al_final: int = 0
    secciones: int = 0


ELEMENTOS = [  # (atributo, cómo se nombra en los mensajes)
    ("tablas", "tablas"),
    ("imagenes", "imágenes"),
    ("notas_al_pie", "notas al pie"),
    ("notas_al_final", "notas al final"),
    ("secciones", "secciones"),
]

_PIEZAS_DE_TEXTO = [w(e) for e in ("t", "delText", "tab", "br", "cr", "noBreakHyphen", "sym")]


def texto_parrafo(p):
    """
    El texto de un párrafo tal como lo escribió el estudiante, con tabuladores
    (→ \t) y saltos de línea (→ \n). No incluye el de los cuadros de texto
    que pueda contener: esos párrafos se leen por separado.
    """
    partes = []
    for el in p.iter(*_PIEZAS_DE_TEXTO):
        if next(el.iterancestors(w("p")), None) is not p:
            continue  # pertenece a un párrafo de dentro (cuadro de texto)
        nombre = etree.QName(el).localname
        if nombre in ("t", "delText"):
            partes.append(el.text or "")
        elif nombre == "tab":
            if etree.QName(el.getparent()).localname == "r":  # no las tabulaciones del formato
                partes.append("\t")
        elif nombre in ("br", "cr"):
            partes.append("\n")
        elif nombre == "noBreakHyphen":
            partes.append("-")
        elif nombre == "sym":
            partes.append(f"[símbolo {el.get(w('char'))}]")
    return "".join(partes)


def _raiz_de_parte(parte):
    """El XML de una parte del documento (python-docx no carga todas como XML)."""
    return parte.element if hasattr(parte, "element") else etree.fromstring(parte.blob)


def _notas(documento, nombre_parte, etiqueta):
    """Las notas al pie o al final de verdad (sin los separadores que Word añade)."""
    for parte in documento.part.package.iter_parts():
        if parte.partname == nombre_parte:
            return [n for n in _raiz_de_parte(parte).iter(w(etiqueta))
                    if n.get(w("type")) in (None, "normal")]
    return []


def radiografia(documento):
    """Texto y recuento de elementos de un documento abierto con python-docx."""
    cuerpo = documento.element.body
    r = Radiografia()
    for i, p in enumerate(cuerpo.iter(w("p")), start=1):
        dentro = " (dentro de una tabla)" if next(p.iterancestors(w("tc")), None) is not None else ""
        r.textos.append((f"Párrafo {i}{dentro}", texto_parrafo(p)))
    for tipo, parte, etiqueta in (("Nota al pie", "/word/footnotes.xml", "footnote"),
                                  ("Nota al final", "/word/endnotes.xml", "endnote")):
        for n, nota in enumerate(_notas(documento, parte, etiqueta), start=1):
            texto = "\n".join(texto_parrafo(p) for p in nota.iter(w("p")))
            r.textos.append((f"{tipo} {n}", texto))

    r.tablas = sum(1 for _ in cuerpo.iter(w("tbl")))
    r.imagenes = sum(1 for _ in cuerpo.iter(w("drawing"), w("pict")))
    r.notas_al_pie = sum(1 for _ in cuerpo.iter(w("footnoteReference")))
    r.notas_al_final = sum(1 for _ in cuerpo.iter(w("endnoteReference")))
    r.secciones = len(documento.sections)
    return r


def _extracto(texto, desde=0, largo=60):
    """Un trozo legible del texto, empezando cerca de `desde`."""
    inicio = max(0, desde - 15)
    trozo = texto[inicio:inicio + largo].replace("\t", "→").replace("\n", " ⏎ ")
    return ("…" if inicio else "") + trozo + ("…" if inicio + largo < len(texto) else "")


# Numeración de apartado al principio de un título, con el espacio que la sigue:
# "1. ", "1.1 ", "2.3.1. ", "1) ", "IV. ", "a) ". Un número sin punto ni paréntesis
# ("20 años de…") no lo es: es parte del título.
NUMERACION_DE_APARTADO = re.compile(r"(?:\d{1,2}(?:\.\d{1,2})*[.)]|\d{1,2}(?:\.\d{1,2})+|[IVXLC]{1,6}[.)]|"
                                    r"[A-Za-z][.)])[ \t ]+")


@dataclass
class CambioPrevisto:
    """
    Un cambio de contenido que una regla hace a propósito (está previsto en
    ROADMAP.md y la regla lo anota en el informe). Solo hay estos:
      - "borrar": se borra un párrafo sin texto (paso 3.5, párrafos vacíos);
      - "recortar": se quita `quitado` del principio del párrafo, que solo
        puede ser espacio en blanco (paso 3.5, tabuladores usados como sangría);
      - "quitar_numeracion": se quita `quitado` del principio del párrafo, que
        solo puede ser una numeración de apartado ("1.1. ", "a) "…) y deja
        texto detrás (paso 4.3, con la opción --quitar-numeracion);
      - "cambiar_mayusculas": el texto pasa a ser `nuevo`, que solo puede
        diferir en mayúsculas y minúsculas de un texto escrito entero en
        mayúsculas (paso 4.3, con la opción --quitar-mayusculas);
      - "quitar_introduccion": se borra el párrafo, que solo puede ser el rótulo
        «Introducción» (con o sin numeración y punto final) (paso 4.4, con la
        opción --quitar-introduccion);
      - "añadir_titulo": se añade antes del párrafo un párrafo nuevo con el
        texto `nuevo`, que tiene que ser una copia exacta de los párrafos
        `desde` del original, línea a línea y sin el espacio de los extremos ni
        los saltos de página (el título del trabajo de la portada, y su
        subtítulo, separados por un salto de línea) (paso 4.4);
      - "añadir_al_final": se añade `nuevo` al final del párrafo, que solo puede
        ser ". " (punto y espacio) o " " si ya acababa en punto o con otro signo,
        y nunca detrás de un espacio (paso 4.5: títulos de nivel 4 y 5 en la misma
        línea que su texto).
    Si un párrafo tiene varios, se comprueban en el orden en que se anotaron.
    """
    parrafo: int     # número del párrafo en el original: desde 1, contando los de las tablas
    accion: str      # "borrar", "recortar", "quitar_numeracion", "cambiar_mayusculas"…
    quitado: str = ""
    nuevo: str = ""
    desde: tuple = ()   # en "añadir_titulo": los números de los párrafos copiados


def lineas_de_texto(texto):
    """Las líneas con texto de un párrafo, sin el espacio de los extremos (un salto de página cuenta como línea)."""
    return [linea.strip() for linea in texto.split("\n") if linea.strip()]


ROTULO_DE_INTRODUCCION = ("introducción", "introduccion")


def es_rotulo_de_introduccion(texto):
    """«Introducción», «1. Introducción», «INTRODUCCIÓN.»…"""
    numeracion = NUMERACION_DE_APARTADO.match(texto)
    resto = texto[numeracion.end():] if numeracion else texto
    return resto.strip().rstrip(".:").strip().lower() in ROTULO_DE_INTRODUCCION


def _solo_cambian_mayusculas(antes, despues):
    """¿Es `despues` el mismo texto que `antes` (escrito entero en mayúsculas), con otras minúsculas?"""
    letras = [c for c in antes if c.isalpha()]
    return (despues != antes and len(despues) == len(antes) and bool(letras) and all(c.isupper() for c in letras)
            and all(a == d or (a.isalpha() and a.lower() == d.lower()) for a, d in zip(antes, despues)))


def _aplicar_previstos(antes, previstos):
    """
    Los textos que debería tener la copia si solo se han hecho los cambios
    previstos, y los problemas si un cambio previsto no es de los permitidos.
    """
    textos = list(antes.textos)
    problemas, borrar, insertar = [], set(), {}

    def copia_de_la_portada(cambio):
        """¿Es `nuevo` una copia exacta de párrafos seguidos del original, anteriores al lugar donde se añade?"""
        desde = list(cambio.desde)
        seguidos = desde and desde == list(range(desde[0], desde[0] + len(desde))) and desde[-1] < cambio.parrafo
        if not seguidos or not all(antes.textos[n - 1][0].startswith(f"Párrafo {n}") for n in desde):
            return False
        copiadas = [linea for n in desde for linea in lineas_de_texto(antes.textos[n - 1][1])]
        return bool(copiadas) and cambio.nuevo == "\n".join(copiadas)

    for cambio in previstos:
        i = cambio.parrafo - 1   # en la radiografía, los párrafos van primero y en orden
        if not 0 <= i < len(textos) or not textos[i][0].startswith(f"Párrafo {cambio.parrafo}"):
            problemas.append(f"Párrafo {cambio.parrafo}: una regla anotó un cambio, pero el párrafo no existe.")
            continue
        donde, texto = textos[i]
        if cambio.accion == "borrar" and not texto.strip():
            borrar.add(i)
        elif cambio.accion == "recortar" and cambio.quitado and not cambio.quitado.strip() \
                and texto.startswith(cambio.quitado):
            textos[i] = (donde, texto[len(cambio.quitado):])
        elif cambio.accion == "quitar_numeracion" and NUMERACION_DE_APARTADO.fullmatch(cambio.quitado) \
                and texto.startswith(cambio.quitado) and texto[len(cambio.quitado):].strip():
            textos[i] = (donde, texto[len(cambio.quitado):])
        elif cambio.accion == "cambiar_mayusculas" and _solo_cambian_mayusculas(texto, cambio.nuevo):
            textos[i] = (donde, cambio.nuevo)
        elif cambio.accion == "quitar_introduccion" and es_rotulo_de_introduccion(texto):
            borrar.add(i)
        elif cambio.accion == "añadir_titulo" and copia_de_la_portada(cambio):
            insertar.setdefault(i, []).append(("Título del trabajo, añadido", cambio.nuevo))
        elif cambio.accion == "añadir_al_final" and texto.strip() and not texto[-1].isspace() and \
                (cambio.nuevo == " " or (cambio.nuevo == ". " and not texto.rstrip("\"'»”)").endswith((".", "…")))):
            textos[i] = (donde, texto + cambio.nuevo)
        else:
            problemas.append(f"{donde}: una regla anotó «{cambio.accion}», pero eso cambiaría el texto "
                             f"(«{_extracto(texto)}»). Solo se pueden borrar párrafos vacíos y el rótulo "
                             "«Introducción», quitar espacios o tabuladores del principio, quitar la numeración "
                             "de un título, pasar a minúsculas un título escrito en mayúsculas, repetir el "
                             "título de la portada al principio del texto y poner el punto final a los títulos "
                             "de nivel 4 y 5.")
    resultado = []
    for i, t in enumerate(textos):
        resultado += insertar.get(i, [])
        if i not in borrar:
            resultado.append(t)
    return resultado, problemas


def comparar(antes, despues, previstos=(), maximo=10):
    """
    Diferencias de contenido entre dos radiografías, explicadas en español.
    Lista vacía = el contenido es idéntico, salvo los cambios previstos
    (CambioPrevisto) que se indiquen, que tienen que haberse hecho tal cual.
    """
    diferencias = []
    for atributo, nombre in ELEMENTOS:
        a, d = getattr(antes, atributo), getattr(despues, atributo)
        if a != d:
            diferencias.append(f"Había {a} {nombre} y ahora hay {d}.")

    esperados, problemas = _aplicar_previstos(antes, previstos)
    diferencias += problemas
    antes = Radiografia(textos=esperados)
    textos_a = [t for _, t in antes.textos]
    textos_d = [t for _, t in despues.textos]
    cambios = difflib.SequenceMatcher(None, textos_a, textos_d, autojunk=False).get_opcodes()
    for operacion, a1, a2, d1, d2 in cambios:
        if operacion == "equal":
            continue
        if operacion == "replace" and a2 - a1 == d2 - d1:
            for i, j in zip(range(a1, a2), range(d1, d2)):
                a, d = textos_a[i], textos_d[j]
                desde = next((k for k, (x, y) in enumerate(zip(a, d)) if x != y), min(len(a), len(d)))
                diferencias.append(f"{antes.textos[i][0]}: el texto ha cambiado: "
                                   f"«{_extracto(a, desde)}» → «{_extracto(d, desde)}»")
            continue
        for i in range(a1, a2):
            diferencias.append(f"{antes.textos[i][0]}: ha desaparecido («{_extracto(textos_a[i])}»).")
        for j in range(d1, d2):
            diferencias.append(f"{despues.textos[j][0]} de la copia: no estaba en el original "
                               f"(«{_extracto(textos_d[j])}»).")

    if len(diferencias) > maximo:
        resto = len(diferencias) - maximo
        diferencias = diferencias[:maximo] + [f"… y {resto} diferencia(s) más."]
    return diferencias


# --- Formato efectivo ---------------------------------------------------------
#
# En Word, el aspecto de un texto sale de tres capas, y manda la de más arriba:
#   1. los valores por defecto del documento;
#   2. los estilos (cada estilo puede heredar de otro: "Título 1" de "Normal"…);
#   3. el formato puesto a mano en el párrafo o en el fragmento de texto.
# La clase Estilos recorre esas capas para saber qué se ve de verdad.

FUENTE_POR_DEFECTO = "Times New Roman"   # lo que usa Word si nada indica otra cosa
TAMANO_POR_DEFECTO = 20                  # en medios puntos: 10 pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _fuentes_del_tema(documento):
    """Las fuentes del tema de Word (p. ej. Calibri para el texto y Calibri Light para títulos)."""
    for parte in documento.part.package.iter_parts():
        if str(parte.partname).startswith("/word/theme/"):
            raiz = _raiz_de_parte(parte)
            fuentes = {}
            for clase in ("major", "minor"):
                latina = raiz.find(f".//{{{A}}}{clase}Font/{{{A}}}latin")
                if latina is not None and latina.get("typeface"):
                    fuentes[clase] = latina.get("typeface")
            return fuentes
    return {}


class Estilos:
    """Calcula el formato efectivo de párrafos y fragmentos de texto."""

    def __init__(self, documento):
        self.raiz = documento.styles.element
        self.por_id = {s.get(w("styleId")): s for s in self.raiz.iter(w("style"))}
        self.parrafo_defecto = next(
            (sid for sid, s in self.por_id.items()
             if s.get(w("type")) == "paragraph" and s.get(w("default")) in ("1", "true")), None)
        defectos = self.raiz.find(w("docDefaults"))
        self.ppr_defecto = defectos.find(f"{w('pPrDefault')}/{w('pPr')}") if defectos is not None else None
        self.rpr_defecto = defectos.find(f"{w('rPrDefault')}/{w('rPr')}") if defectos is not None else None
        self.tema = _fuentes_del_tema(documento)

    def cadena(self, sid):
        """El estilo y todos aquellos de los que hereda, del más cercano al más lejano."""
        vistos = set()
        while sid and sid in self.por_id and sid not in vistos:
            vistos.add(sid)
            estilo = self.por_id[sid]
            yield estilo
            base = estilo.find(w("basedOn"))
            sid = base.get(w("val")) if base is not None else None

    def nombre(self, sid):
        estilo = self.por_id.get(sid)
        nombre = estilo.find(w("name")) if estilo is not None else None
        return nombre.get(w("val")) if nombre is not None else (sid or "")

    def estilo_de(self, p):
        ps = p.find(f"{w('pPr')}/{w('pStyle')}")
        return ps.get(w("val")) if ps is not None else self.parrafo_defecto

    def _capas_parrafo(self, p):
        yield p.find(w("pPr"))
        for estilo in self.cadena(self.estilo_de(p)):
            yield estilo.find(w("pPr"))
        yield self.ppr_defecto

    def _capas_texto(self, r, p):
        rpr = r.find(w("rPr"))
        yield rpr
        estilo_r = rpr.find(w("rStyle")) if rpr is not None else None
        if estilo_r is not None:
            for estilo in self.cadena(estilo_r.get(w("val"))):
                yield estilo.find(w("rPr"))
        for estilo in self.cadena(self.estilo_de(p)):
            yield estilo.find(w("rPr"))
        yield self.rpr_defecto

    def parrafo(self, p, etiqueta, atributo):
        """Valor efectivo de un atributo de párrafo, p. ej. ("spacing", "line") o ("jc", "val")."""
        for capa in self._capas_parrafo(p):
            el = capa.find(w(etiqueta)) if capa is not None else None
            if el is not None and el.get(w(atributo)) is not None:
                return el.get(w(atributo))
        return None

    def fuente(self, r, p):
        """Nombre de la fuente con que se ve un fragmento de texto."""
        for capa in self._capas_texto(r, p):
            fuentes = capa.find(w("rFonts")) if capa is not None else None
            if fuentes is None:
                continue
            tema = fuentes.get(w("asciiTheme")) or fuentes.get(w("hAnsiTheme"))
            if tema:  # el tema manda sobre el nombre escrito
                clase = "major" if tema.startswith("major") else "minor"
                return self.tema.get(clase, "(fuente del tema)")
            nombre = fuentes.get(w("ascii")) or fuentes.get(w("hAnsi"))
            if nombre:
                return nombre
        return FUENTE_POR_DEFECTO

    def tamano(self, r, p):
        """Tamaño en puntos con que se ve un fragmento de texto."""
        for capa in self._capas_texto(r, p):
            sz = capa.find(w("sz")) if capa is not None else None
            if sz is not None and sz.get(w("val")):
                return int(sz.get(w("val"))) / 2
        return TAMANO_POR_DEFECTO / 2


CAMPOS = {"PAGE": "[nº de página]", "NUMPAGES": "[total de páginas]",
          "SECTIONPAGES": "[páginas de la sección]", "DATE": "[fecha]", "TOC": "[índice]"}


def texto_legible(p):
    """
    El texto de un párrafo para mostrarlo a una persona: los campos de Word
    se nombran con palabras ("[nº de página]") en vez de mostrar su código.
    """
    partes, en_resultado, esperando_codigo = [], False, False
    for el in p.iter(w("t"), w("instrText"), w("fldChar"), w("fldSimple"), w("tab"), w("br")):
        nombre = etree.QName(el).localname
        if nombre == "fldChar":
            tipo = el.get(w("fldCharType"))
            esperando_codigo = tipo == "begin"
            en_resultado = tipo == "separate"   # entre "separate" y "end" va el valor calculado
        elif nombre == "instrText":
            codigo = (el.text or "").split()
            if esperando_codigo and codigo:     # el código puede venir partido en trozos
                partes.append(CAMPOS.get(codigo[0].upper(), f"[campo {codigo[0]}]"))
                esperando_codigo = False
        elif nombre == "fldSimple":
            codigo = (el.get(w("instr")) or "").split()
            partes.append(CAMPOS.get(codigo[0].upper(), "[campo]") if codigo else "[campo]")
        elif next(el.iterancestors(w("fldSimple")), None) is not None or en_resultado:
            continue
        elif nombre == "t":
            partes.append(el.text or "")
        elif nombre == "tab" and etree.QName(el.getparent()).localname == "r":
            partes.append("\t")
        elif nombre == "br":
            partes.append("\n")
    return "".join(partes)


# --- Secciones y anchos -------------------------------------------------------
#
# Word mide casi todo en "twips" (veinteavos de punto: 1 cm ≈ 567 twips) y las
# imágenes en EMU (1 twip = 635 EMU). Cada sección acaba en un párrafo que
# lleva sus ajustes (w:sectPr); la última sección los lleva al final del cuerpo.

TWIPS_POR_CM = 567
EMU_POR_TWIP = 635

WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
TIPOS_DE_DIBUJO = {"picture": "imagen", "chart": "gráfico"}


def cm(twips):
    """1440 → "2,54"."""
    return f"{int(twips) / TWIPS_POR_CM:.2f}".replace(".", ",")


def bloques_por_seccion(documento):
    """
    Los bloques del cuerpo (párrafos, tablas…) de cada sección, en el mismo
    orden que documento.sections.
    """
    secciones, actual = [], []
    for bloque in documento.element.body:
        if bloque.tag == w("sectPr"):
            continue
        actual.append(bloque)
        if bloque.tag == w("p") and bloque.find(f"{w('pPr')}/{w('sectPr')}") is not None:
            secciones.append(actual)
            actual = []
    secciones.append(actual)
    return secciones


def _entero(valor):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def _dentro_de(el, *etiquetas):
    return next(el.iterancestors(*etiquetas), None) is not None


def tablas_de(bloques):
    """Las tablas de primer nivel (no las que van dentro de otra tabla)."""
    return [t for b in bloques for t in b.iter(w("tbl")) if not _dentro_de(t, w("tbl"))]


def ancho_tabla(tbl):
    """Ancho de una tabla en twips, o None si es relativo (en %) o no se sabe."""
    ancho = tbl.find(f"{w('tblPr')}/{w('tblW')}")
    if ancho is not None and ancho.get(w("type")) == "pct":
        return None   # se adapta sola al espacio disponible
    columnas = sum(_entero(c.get(w("w"))) for c in tbl.iterfind(f"{w('tblGrid')}/{w('gridCol')}"))
    if columnas:
        return columnas
    if ancho is not None and ancho.get(w("type")) == "dxa":
        return _entero(ancho.get(w("w"))) or None
    return None


def dibujos_de(bloques):
    """
    Imágenes y gráficos que ocupan sitio en la página: (tipo, ancho en twips).
    Se omiten los de dentro de tablas (ya cuenta la tabla) y de cuadros de
    texto, y las copias de compatibilidad que Word guarda para versiones antiguas.
    """
    dibujos = []
    for b in bloques:
        for el in b.iter(f"{{{WP}}}inline", f"{{{WP}}}anchor"):
            if _dentro_de(el, w("tc"), w("txbxContent"), f"{{{MC}}}Fallback"):
                continue
            datos = el.find(f"{{{A}}}graphic/{{{A}}}graphicData")
            uri = datos.get("uri", "") if datos is not None else ""
            tipo = next((t for clave, t in TIPOS_DE_DIBUJO.items() if uri.endswith(clave)), None)
            extension = el.find(f"{{{WP}}}extent")
            if tipo and extension is not None:
                dibujos.append((tipo, _entero(extension.get("cx")) // EMU_POR_TWIP))
    return dibujos


# --- Fuentes ------------------------------------------------------------------
#
# La fuente de un fragmento se indica en w:rFonts con cuatro atributos (letras
# latinas, otras latinas, asiáticas y alfabetos complejos). Si además lleva
# atributos de tema (asciiTheme…), Word usa la fuente del tema e ignora el
# nombre escrito: por eso hay que quitarlos al poner una fuente.

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_ATRIBUTOS_FUENTE = [w(a) for a in ("ascii", "hAnsi", "eastAsia", "cs")]
_ATRIBUTOS_TEMA = [w(a) for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme")]
_PARTES_DE_TEXTO = ("/word/document.xml", "/word/header", "/word/footer",
                    "/word/footnotes.xml", "/word/endnotes.xml")


def poner_fuente(rpr, nombre):
    """Pone la fuente `nombre` en unas propiedades de texto (w:rPr)."""
    fuentes = rpr.get_or_add_rFonts()
    for atributo in _ATRIBUTOS_TEMA:
        fuentes.attrib.pop(atributo, None)
    for atributo in _ATRIBUTOS_FUENTE:
        fuentes.set(atributo, nombre)


def poner_tamano(rpr, puntos):
    """Pone el tamaño en puntos (Word lo guarda en medios puntos, en sz y szCs)."""
    medios = str(round(puntos * 2))
    sz = rpr.get_or_add_sz()
    sz.set(w("val"), medios)
    szcs = rpr.find(w("szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        sz.addnext(szcs)
    szcs.set(w("val"), medios)


def propiedades_por_defecto(documento):
    """Las propiedades de texto por defecto del documento (w:docDefaults), creándolas si faltan."""
    raiz = documento.styles.element
    defectos = raiz.find(w("docDefaults"))
    if defectos is None:
        defectos = OxmlElement("w:docDefaults")
        raiz.insert(0, defectos)
    rpr_defecto = defectos.find(w("rPrDefault"))
    if rpr_defecto is None:
        rpr_defecto = OxmlElement("w:rPrDefault")
        defectos.insert(0, rpr_defecto)
    rpr = rpr_defecto.find(w("rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rpr_defecto.append(rpr)
    return rpr


def en_ecuacion(el):
    return next(el.iterancestors(f"{{{M}}}oMath", f"{{{M}}}oMathPara"), None) is not None


def editar_partes_de_texto(documento, funcion):
    """
    Llama a funcion(nombre_de_parte, raiz_xml) con cada parte que tiene texto:
    cuerpo, encabezados, pies y notas. Las notas no las carga python-docx como
    XML, así que se leen, se modifican y se vuelven a guardar en la parte.
    """
    for parte in documento.part.package.iter_parts():
        nombre = str(parte.partname)
        if not nombre.startswith(_PARTES_DE_TEXTO):
            continue
        if hasattr(parte, "element"):
            funcion(nombre, parte.element)
        else:
            raiz = parse_xml(parte.blob)
            funcion(nombre, raiz)
            parte._blob = etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


# --- Espaciado de párrafo -------------------------------------------------------
#
# w:spacing guarda el interlineado (line, en 240 avos de línea si lineRule es
# "auto": 240 = sencillo, 360 = 1,5, 480 = doble; en twips si es "exact" o
# "atLeast") y el espacio antes y después del párrafo (before/after, en twips).
# Hay otros atributos que mandan sobre before/after: las medidas en líneas
# (beforeLines…) y el espaciado automático de los textos pegados de la web.

_ESPACIADO_QUE_ANULA = [w(a) for a in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing")]


def poner_espaciado(ppr, linea, regla="auto", antes=0, despues=0):
    """Pone interlineado y espacio antes y después en unas propiedades de párrafo (w:pPr)."""
    espaciado = ppr.get_or_add_spacing()
    for atributo in _ESPACIADO_QUE_ANULA:
        espaciado.attrib.pop(atributo, None)
    espaciado.set(w("line"), str(linea))
    espaciado.set(w("lineRule"), regla)
    espaciado.set(w("before"), str(antes))
    espaciado.set(w("after"), str(despues))


def parrafo_por_defecto(documento):
    """Las propiedades de párrafo por defecto del documento, creándolas si faltan."""
    defectos = propiedades_por_defecto(documento).getparent().getparent()   # w:docDefaults
    ppr_defecto = defectos.find(w("pPrDefault"))
    if ppr_defecto is None:
        ppr_defecto = OxmlElement("w:pPrDefault")
        defectos.append(ppr_defecto)
    ppr = ppr_defecto.find(w("pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        ppr_defecto.append(ppr)
    return ppr
