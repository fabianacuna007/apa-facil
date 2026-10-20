"""
Decide en qué orden se aplican las reglas APA a un documento:
abrir, diagnosticar, detectar la estructura, aplicar las reglas y verificar.

Las reglas se van añadiendo a REGLAS a lo largo de la fase 1.
"""

from datetime import datetime
from pathlib import Path

from apa_facil import datos_portada, documento, estructura
from apa_facil.documento import ErrorDocumento
from apa_facil.informe import Informe
from apa_facil.revision import comentarios, corregir, revisar
from apa_facil.reglas import (alineacion, apendices, avisos, citas, figuras, fuente, indice, interlineado, limpieza, margenes,
                              orden, paginacion, portada, referencias, resumen, sangria, secciones, tablas, titulos)

# En este orden. Los avisos van primero: después la regla de fuente lo iguala todo.
# Los títulos, después de la sangría: esa regla quita la sangría de primera línea
# de todos los estilos, y los títulos de nivel 4 y 5 la llevan. La portada, después de los
# títulos: usa el estilo «Título» con el formato que le da el paso 4.4. Las tablas, después de los
# títulos: sus títulos llevan las mayúsculas de APA, como los de los apartados. Las figuras, detrás de las
# tablas: un solo recordatorio de dónde colocarlas, que amplía el de las tablas.
REGLAS = [avisos, secciones, limpieza, margenes, fuente, interlineado, alineacion, sangria, titulos, resumen, tablas,
          figuras, citas, referencias, orden, apendices, portada, indice, paginacion]


def procesar(entrada, carpeta_salida=None, opciones=None):
    """
    Abre `entrada`, le aplica las reglas y guarda la copia y su informe en JSON
    (salida/nombre_APA.docx y salida/nombre_APA.json).
    Devuelve la ruta de la copia y el informe de lo que se hizo.
    """
    opciones = opciones or {}
    doc = documento.abrir(entrada)
    antes = documento.radiografia(doc)  # siempre antes de tocar nada

    salida = documento.ruta_salida(entrada, carpeta_salida)
    informe = Informe(archivo=Path(entrada).name, copia=salida.name, opciones=dict(opciones),
                      fecha=datetime.now().isoformat(timespec="seconds"))
    # Qué es cada párrafo (título, texto, lista…), decidido sobre el original:
    # después de la regla de fuente, por ejemplo, ya no se ve qué títulos tenían
    # la letra más grande.
    clasificados = estructura.clasificar(doc)
    # La portada que se mantiene no la toca ninguna regla de formato (revisión del propietario,
    # 2026-10-20): su formato se fija ahora, antes de que cambien los estilos.
    clasificados.intactos = portada.intactos(opciones, clasificados)
    portada.congelar(doc, clasificados.intactos)
    # Revisión de citas y referencias (fase 7): sobre el original, antes de que las
    # reglas lo cambien (reordenar las referencias, borrar vacíos…). Solo da avisos,
    # que van al final del informe, detrás de los de formato.
    revision = revisar.analizar(doc, clasificados)
    # Datos de la portada (paso 5.1): se avisa de los que faltan. La portada la pone su regla.
    if opciones.get("portada"):
        informe.regla = "portada"
        datos_portada.avisar(opciones["portada"], informe)
    for regla in REGLAS:
        informe.regla = regla.__name__.rsplit(".", 1)[-1]   # cada mensaje sabe de qué regla viene
        regla.aplicar(doc, opciones, informe, clasificados)
    informe.regla = "revision"
    # Paso 7.7: los errores de forma que se pueden corregir sin inventar nada se corrigen
    # (salvo con --citas-tal-cual); los demás, y todo el cruce, solo se avisan.
    if not opciones.get("citas_tal_cual"):
        revision.corregidos = corregir.corregir(revision, informe)
    revisar.avisar(revision, informe)
    if opciones.get("comentarios"):   # paso 7.6: los avisos, también como comentarios en la copia
        comentarios.comentar(doc, revision, informe)
    # Lo último: unir los párrafos que un salto de sección partió en dos (las correcciones de las
    # citas y los comentarios ya están puestos en cada uno de ellos).
    informe.regla = "secciones"
    secciones.unir_partidos(doc, clasificados, informe)
    informe.regla = None

    documento.guardar(doc, salida)
    ruta_informe = salida.with_suffix(".json")

    # Verificación: se vuelve a abrir la copia ya guardada y se compara con el
    # original. Solo se admiten los cambios de contenido que las reglas han anotado
    # (y que son de los permitidos). Si no coincide, la copia no se entrega (ni su
    # informe, ni el de una vez anterior).
    diferencias = documento.comparar(antes, documento.radiografia(documento.abrir(salida)),
                                     informe.cambios_previstos)
    if diferencias:
        salida.unlink(missing_ok=True)
        ruta_informe.unlink(missing_ok=True)
        detalle = "\n".join(f"     - {d}" for d in diferencias)
        raise ErrorDocumento("No se ha guardado la copia porque su contenido no coincide con el "
                             f"original. Por seguridad, no se entrega:\n{detalle}")

    informe.guardar_json(ruta_informe)
    return salida, informe
