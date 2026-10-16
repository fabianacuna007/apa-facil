"""
Decide en qué orden se aplican las reglas APA a un documento:
abrir, diagnosticar, detectar la estructura, aplicar las reglas y verificar.

Las reglas se van añadiendo a REGLAS a lo largo de la fase 1.
"""

from apa_facil import documento
from apa_facil.documento import ErrorDocumento
from apa_facil.informe import Informe
from apa_facil.reglas import fuente, margenes

REGLAS = [margenes, fuente]   # en este orden


def procesar(entrada, carpeta_salida=None, opciones=None):
    """
    Abre `entrada`, le aplica las reglas y guarda la copia.
    Devuelve la ruta de la copia y el informe de lo que se hizo.
    """
    doc = documento.abrir(entrada)
    antes = documento.radiografia(doc)  # siempre antes de tocar nada

    informe = Informe()
    for regla in REGLAS:
        regla.aplicar(doc, opciones or {}, informe)

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
    return salida, informe
