"""
Decide en qué orden se aplican las reglas APA a un documento:
abrir, diagnosticar, detectar la estructura, aplicar las reglas y verificar.

De momento solo abre, guarda y verifica: las reglas llegan en la fase 1.
"""

from apa_facil import documento
from apa_facil.documento import ErrorDocumento


def procesar(entrada, carpeta_salida=None):
    """Abre `entrada`, la procesa y guarda la copia. Devuelve la ruta de la copia."""
    doc = documento.abrir(entrada)
    antes = documento.radiografia(doc)  # siempre antes de tocar nada

    salida = documento.ruta_salida(entrada, carpeta_salida)
    documento.guardar(doc, salida)

    # Verificación: se vuelve a abrir la copia ya guardada y se compara con el
    # original. Si el contenido no coincide, la copia no se entrega.
    diferencias = documento.comparar(antes, documento.radiografia(documento.abrir(salida)))
    if diferencias:
        salida.unlink(missing_ok=True)
        detalle = "\n".join(f"     - {d}" for d in diferencias)
        raise ErrorDocumento("No se ha guardado la copia porque su contenido no coincide con el "
                             f"original. Por seguridad, no se entrega:\n{detalle}")
    return salida
