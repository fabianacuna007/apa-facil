"""
Paso 8.2: el núcleo se ejecuta en un proceso aparte para poder pararlo si se
pasa del tiempo máximo (un hilo de Python no se puede parar desde fuera).

El proceso devuelve por una tubería solo lo necesario: el informe ya convertido
en diccionario y en texto, o el mensaje del error. La copia queda en la carpeta
temporal, que el servidor lee y borra.

Paso 10.5: también devuelve lo que ha costado el documento (segundos y memoria
máxima del proceso), para saber cuántos documentos aguanta el servidor.
"""

import ctypes
import multiprocessing
import sys
import time

from apa_facil import pipeline
from apa_facil.documento import ErrorDocumento

# «spawn» en todos los sistemas: un proceso nuevo y limpio, igual en Windows que en Linux (Docker).
CONTEXTO = multiprocessing.get_context("spawn")


class TiempoAgotado(Exception):
    pass


def memoria_maxima_mb():
    """La memoria máxima que ha usado este proceso, en MB (None si el sistema no lo dice)."""
    try:
        if sys.platform == "win32":
            class Contadores(ctypes.Structure):
                _fields_ = [("cb", ctypes.c_ulong), ("PageFaultCount", ctypes.c_ulong),
                            ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t),
                            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t), ("QuotaPagedPoolUsage", ctypes.c_size_t),
                            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                            ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t),
                            ("PeakPagefileUsage", ctypes.c_size_t)]
            contadores = Contadores()
            contadores.cb = ctypes.sizeof(contadores)
            kernel32 = ctypes.WinDLL("kernel32")
            kernel32.GetCurrentProcess.restype = ctypes.c_void_p   # sin esto, el identificador se corta a 32 bits
            consultar = kernel32.K32GetProcessMemoryInfo
            consultar.argtypes = [ctypes.c_void_p, ctypes.POINTER(Contadores), ctypes.c_ulong]
            if not consultar(kernel32.GetCurrentProcess(), ctypes.byref(contadores), contadores.cb):
                return None
            return contadores.PeakWorkingSetSize / 2**20
        import resource   # Linux (Docker): ru_maxrss en KB
        return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024
    except Exception:
        return None


def _trabajar(entrada, carpeta_salida, opciones, nombre, copia, conexion):
    inicio = time.perf_counter()
    try:
        salida, informe = pipeline.procesar(entrada, carpeta_salida, opciones)
        informe.archivo, informe.copia = nombre, copia
        coste = {"segundos": time.perf_counter() - inicio, "memoria_mb": memoria_maxima_mb()}
        conexion.send(("bien", str(salida), informe.a_dict(), informe.como_texto(), coste))
    except ErrorDocumento as e:
        conexion.send(("error", str(e)))
    except Exception as e:   # solo el tipo: el mensaje podría llevar texto del documento
        conexion.send(("fallo", type(e).__name__))
    finally:
        conexion.close()


def procesar(entrada, carpeta_salida, opciones, nombre, copia, tiempo_maximo):
    """
    Ejecuta pipeline.procesar en otro proceso. Devuelve ("bien", ruta, informe, resumen, coste),
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
