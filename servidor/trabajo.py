"""
Paso 8.2: el núcleo se ejecuta en un proceso aparte para poder pararlo si se
pasa del tiempo máximo (un hilo de Python no se puede parar desde fuera).

El proceso devuelve por una tubería solo lo necesario: el informe ya convertido
en diccionario y en texto, o el mensaje del error. La copia queda en la carpeta
temporal, que el servidor lee y borra.
"""

import multiprocessing

from apa_facil import pipeline
from apa_facil.documento import ErrorDocumento

# «spawn» en todos los sistemas: un proceso nuevo y limpio, igual en Windows que en Linux (Docker).
CONTEXTO = multiprocessing.get_context("spawn")


class TiempoAgotado(Exception):
    pass


def _trabajar(entrada, carpeta_salida, opciones, nombre, copia, conexion):
    try:
        salida, informe = pipeline.procesar(entrada, carpeta_salida, opciones)
        informe.archivo, informe.copia = nombre, copia
        conexion.send(("bien", str(salida), informe.a_dict(), informe.como_texto()))
    except ErrorDocumento as e:
        conexion.send(("error", str(e)))
    except Exception as e:   # solo el tipo: el mensaje podría llevar texto del documento
        conexion.send(("fallo", type(e).__name__))
    finally:
        conexion.close()


def procesar(entrada, carpeta_salida, opciones, nombre, copia, tiempo_maximo):
    """
    Ejecuta pipeline.procesar en otro proceso. Devuelve ("bien", ruta, informe, resumen),
    ("error", mensaje) o ("fallo", tipo). Si tarda más de `tiempo_maximo` segundos,
    para el proceso y lanza TiempoAgotado.
    """
    recibir, enviar = CONTEXTO.Pipe(duplex=False)
    proceso = CONTEXTO.Process(target=_trabajar, daemon=True,
                               args=(str(entrada), str(carpeta_salida), opciones, nombre, copia, enviar))
    proceso.start()
    enviar.close()   # en este proceso solo se recibe
    try:
        if not recibir.poll(tiempo_maximo):
            raise TiempoAgotado
        try:
            return recibir.recv()
        except EOFError:   # el proceso murió sin responder (por ejemplo, sin memoria)
            return ("fallo", "EOFError")
    finally:
        recibir.close()
        if proceso.is_alive():
            proceso.kill()
        proceso.join(10)
