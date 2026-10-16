"""
Abrir y guardar documentos Word, y utilidades para trabajar con su XML.

Todo lo que python-docx no cubre (campos de número de página, notas al pie,
algunos ajustes de sección) se hace aquí con lxml, para que el resto del
programa no tenga que tocar el XML directamente.
"""

import zipfile
from pathlib import Path

import docx
from docx.opc.exceptions import PackageNotFoundError
from lxml import etree

CARPETA_SALIDA = Path(__file__).resolve().parent.parent / "salida"
SUFIJO = "_APA"


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
