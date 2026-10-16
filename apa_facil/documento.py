"""
Abrir y guardar documentos Word, y utilidades para trabajar con su XML.

Todo lo que python-docx no cubre (campos de número de página, notas al pie,
algunos ajustes de sección) se hace aquí con lxml, para que el resto del
programa no tenga que tocar el XML directamente.
"""

import difflib
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import docx
from docx.opc.exceptions import PackageNotFoundError
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


def comparar(antes, despues, maximo=10):
    """
    Diferencias de contenido entre dos radiografías, explicadas en español.
    Lista vacía = el contenido es idéntico.
    """
    diferencias = []
    for atributo, nombre in ELEMENTOS:
        a, d = getattr(antes, atributo), getattr(despues, atributo)
        if a != d:
            diferencias.append(f"Había {a} {nombre} y ahora hay {d}.")

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
