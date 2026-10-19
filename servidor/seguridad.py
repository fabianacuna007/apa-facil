"""
Paso 8.2: comprobaciones del archivo subido antes de dárselo al núcleo.

Un .docx es un ZIP. Se comprueba que lo sea de verdad (aunque termine en .docx)
y que no sea una «bomba ZIP»: un archivo pequeño que, al descomprimirlo, ocupa
gigas y deja el servidor sin memoria. Los tamaños descomprimidos que se miran
son los que declara el ZIP; zipfile nunca descomprime más de lo declarado, así
que el límite se cumple aunque el archivo mienta.

(Las «bombas XML», entidades que se multiplican al leerlas, no hacen falta
comprobarlas: python-docx lee el XML sin resolver entidades.)
"""

import os
import zipfile

from apa_facil import documento
from apa_facil.documento import ErrorDocumento

MB = 1024 * 1024
TAMANO_MAXIMO = int(os.environ.get("APA_TAMANO_MAXIMO_MB", 10)) * MB
DESCOMPRIMIDO_MAXIMO = 200 * MB     # todo el documento, descomprimido
PROPORCION_MAXIMA = 200             # cuánto puede crecer una pieza grande al descomprimirla
PIEZAS_MAXIMAS = 2000               # un trabajo normal tiene menos de 100
OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # formato de los .doc antiguos y de los Word con contraseña


class ErrorEntrada(ErrorDocumento):
    """Un archivo que no se acepta; `estado` es el código HTTP de la respuesta."""

    def __init__(self, mensaje, estado=422):
        super().__init__(mensaje)
        self.estado = estado


def demasiado_grande(nombre):
    return ErrorEntrada(f"«{nombre}» ocupa más de {TAMANO_MAXIMO // MB} MB, el máximo que se admite. "
                        "Si tiene muchas imágenes, comprímelas en Word (Formato de imagen › Comprimir "
                        "imágenes) y vuelve a subirlo.", estado=413)


def comprobar(ruta, nombre):
    """Lanza ErrorEntrada si el archivo `ruta` (que el estudiante subió como `nombre`) no se acepta."""
    documento.comprobar_nombre(nombre)
    tamano = os.path.getsize(ruta)
    if tamano == 0:
        raise ErrorEntrada(f"«{nombre}» está vacío (ocupa 0 bytes). Comprueba que has elegido el archivo bueno.")
    if tamano > TAMANO_MAXIMO:
        raise demasiado_grande(nombre)

    with open(ruta, "rb") as f:
        inicio = f.read(8)
    if inicio == OLE:
        raise ErrorEntrada(f"«{nombre}» tiene contraseña o está guardado en el formato antiguo de Word. "
                           "Ábrelo en Word, quita la contraseña (Archivo › Información › Proteger documento) "
                           "y usa «Guardar como» → «Documento de Word (.docx)».")
    danado = ErrorEntrada(f"«{nombre}» está dañado o no es un documento de Word de verdad (aunque termine en .docx).")
    if not inicio.startswith(b"PK"):
        raise danado
    try:
        with zipfile.ZipFile(ruta) as zip_:
            piezas = zip_.infolist()
    except (zipfile.BadZipFile, OSError, ValueError):
        raise danado from None

    nombres = {pieza.filename for pieza in piezas}
    if "[Content_Types].xml" not in nombres or not any(n.startswith("word/") for n in nombres):
        raise danado
    bomba = ErrorEntrada(f"«{nombre}» tiene un contenido comprimido anómalo y no se puede procesar con "
                         "seguridad. Ábrelo en Word y guárdalo de nuevo como .docx.")
    if len(piezas) > PIEZAS_MAXIMAS or sum(p.file_size for p in piezas) > DESCOMPRIMIDO_MAXIMO:
        raise bomba
    for pieza in piezas:
        if pieza.file_size > MB and pieza.file_size > PROPORCION_MAXIMA * max(pieza.compress_size, 1):
            raise bomba
