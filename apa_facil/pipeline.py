"""
Decide en qué orden se aplican las reglas APA a un documento:
abrir, diagnosticar, detectar la estructura, aplicar las reglas y verificar.

De momento (paso 0.4) solo abre y guarda: las reglas llegan en la fase 1.
"""

from apa_facil import documento


def procesar(entrada, carpeta_salida=None):
    """Abre `entrada`, la procesa y guarda la copia. Devuelve la ruta de la copia."""
    doc = documento.abrir(entrada)
    salida = documento.ruta_salida(entrada, carpeta_salida)
    documento.guardar(doc, salida)
    return salida
