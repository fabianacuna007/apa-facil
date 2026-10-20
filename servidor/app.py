"""
Servidor de APA Fácil (fase 8): la web le envía el Word y las opciones, y él
llama al núcleo igual que la línea de comandos.

    uvicorn servidor.app:app

En http://localhost:8000/docs hay una página para probarlo a mano.

POST /procesar recibe un formulario con el archivo, las opciones y los datos de
portada, y devuelve un JSON con:
  - "informe": el informe completo (el mismo que la línea de comandos guarda en salida/);
  - "resumen": el informe en texto legible (lo que se ha corregido y lo que hay que revisar);
  - "documento": la copia corregida, con su nombre y su contenido en base64.

POST /procesar-texto (paso 9.7) recibe el trabajo pegado como texto («texto») en
vez del archivo, con las mismas opciones, y devuelve lo mismo.

El archivo se guarda en una carpeta temporal con un nombre fijo (nunca el del
estudiante: «con.docx» o «com1.docx» son nombres reservados de Windows), y la
carpeta se borra al terminar, vaya bien o mal.

Seguridad (paso 8.2): solo .docx de verdad, 10 MB como máximo, nada de bombas
ZIP (servidor/seguridad.py) y un tiempo máximo por documento: el núcleo trabaja
en un proceso aparte que se para si se pasa (servidor/trabajo.py). Como mucho se
procesan SIMULTANEOS documentos a la vez; los demás esperan su turno.

Privacidad (paso 8.3): no se guarda nada. El archivo, la copia y el informe
viven solo en la carpeta temporal de la petición (APA_CARPETA_TEMPORAL, o la
del sistema), que se borra al terminar, también si hay un error o se agota el
tiempo. El registro (log) nunca lleva el nombre del archivo ni su texto: de un
error inesperado solo se anota el tipo, y el servidor solo anota la ruta y el
código de cada petición («POST /procesar 200»).

Errores (paso 11.3): si existe SENTRY_DSN, los errores inesperados (los 500) se
avisan con Sentry, solo con su tipo y su sitio en el código (servidor/errores.py).
"""

import base64
import contextlib
import logging
import os
import shutil
import tempfile
import threading
import time
from pathlib import Path
from typing import Annotated

from fastapi import Depends, FastAPI, File, Form, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from apa_facil import datos_portada, desde_texto, documento
from apa_facil.documento import ErrorDocumento
from apa_facil.reglas import fuente
from servidor import errores, limites, seguridad, trabajo

NOMBRE_TEMPORAL = "documento.docx"
TIEMPO_MAXIMO = float(os.environ.get("APA_TIEMPO_MAXIMO", 120))   # segundos por documento
SIMULTANEOS = int(os.environ.get("APA_SIMULTANEOS", 2))            # documentos procesándose a la vez
ESPERA_MAXIMA = 60                                                 # segundos en cola antes de rendirse
TURNOS = threading.BoundedSemaphore(SIMULTANEOS)
CARPETA_TEMPORAL = os.environ.get("APA_CARPETA_TEMPORAL")          # None: la del sistema
# Webs que pueden llamar al servidor desde el navegador (paso 9.1), separadas por comas.
ORIGENES = [o.strip() for o in os.environ.get("APA_ORIGENES", "http://localhost:3000").split(",") if o.strip()]
REGISTRO = logging.getLogger("apa_facil.servidor")
LIMITE = limites.Limite()                                          # paso 10.5: documentos al día por IP
# Paso 11.1, temporal: /limite enseña a quien pregunta sus propias cabeceras de IP, para
# saber cuántos proxies pone el alojamiento delante (APA_PROXIES). Apagado por defecto.
DIAGNOSTICO_IP = os.environ.get("APA_DIAGNOSTICO_IP") == "1"
if not REGISTRO.handlers:   # paso 10.5: que se vea en la terminal lo que cuesta cada documento (nivel INFO)
    _salida = logging.StreamHandler()
    _salida.setFormatter(logging.Formatter("%(levelname)s:     %(message)s"))
    REGISTRO.addHandler(_salida)
    REGISTRO.setLevel(logging.INFO)
# Lo que puede ocupar la petición entera: el archivo y, como mucho, 1 MB de opciones y cabeceras.
PETICION_MAXIMA = seguridad.TAMANO_MAXIMO + seguridad.MB
FUENTES = ", ".join(list(fuente.RECOMENDADAS)[:-1]) + " o " + list(fuente.RECOMENDADAS)[-1]

errores.activar()   # paso 11.3: solo si hay SENTRY_DSN

app = FastAPI(title="APA Fácil", description="Aplica el formato APA 7 de estudiante a un documento Word.")


class _PeticionGrande(Exception):
    pass


def _limitar_tamano(siguiente):
    """
    Corta una subida demasiado grande antes de que llegue entera: FastAPI lee el
    formulario completo (y lo guarda en disco) antes de llamar a `procesar`.
    """
    async def asgi(scope, receive, send):
        if scope["type"] != "http":
            return await siguiente(scope, receive, send)
        cabeceras = dict(scope.get("headers") or [])
        largo = cabeceras.get(b"content-length", b"")
        if largo.isdigit() and int(largo) > PETICION_MAXIMA:
            return await _error(413, _MENSAJE_GRANDE)(scope, receive, send)
        recibido = 0

        async def contar():
            nonlocal recibido
            mensaje = await receive()
            recibido += len(mensaje.get("body", b""))
            if recibido > PETICION_MAXIMA:
                raise _PeticionGrande
            return mensaje
        empezada = False

        async def enviar(mensaje):
            nonlocal empezada
            empezada = True
            await send(mensaje)
        try:
            await siguiente(scope, contar, enviar)
        except _PeticionGrande:
            await _error(413, _MENSAJE_GRANDE)(scope, receive, send)
        except Exception as e:
            # Paso 8.3: solo el tipo del error, nunca su mensaje ni la traza (podrían llevar
            # texto del documento). Sin esto, el servidor escribiría la traza completa en el registro.
            REGISTRO.error("Error inesperado en el servidor (%s).", type(e).__name__)
            errores.avisar(e)
            if not empezada:
                await _error(500, _MENSAJE_FALLO)(scope, receive, send)
    return asgi


_MENSAJE_GRANDE = (f"El archivo ocupa más de {seguridad.TAMANO_MAXIMO // seguridad.MB} MB, el máximo que se "
                   "admite. Si tiene muchas imágenes, comprímelas en Word (Formato de imagen › Comprimir "
                   "imágenes) y vuelve a subirlo.")
_MENSAJE_FALLO = ("Ha habido un error inesperado al procesar el documento. No se ha guardado nada. "
                  "Inténtalo de nuevo o prueba con otro archivo.")
app.add_middleware(_limitar_tamano)
# Paso 9.1: la web está en otro origen (localhost:3000 en local, su dominio en producción),
# y el navegador solo le deja leer la respuesta si el servidor lo permite. Va después de
# _limitar_tamano para envolverlo: así la web también puede leer sus errores 413.
app.add_middleware(CORSMiddleware, allow_origins=ORIGENES, allow_methods=["GET", "POST"])


def _error(estado, mensaje):
    return JSONResponse(status_code=estado, content={"error": mensaje})


@app.exception_handler(RequestValidationError)
def _peticion_incompleta(peticion, exc):
    """FastAPI explica en inglés lo que falta en la petición; aquí, en español."""
    if any(tuple(e.get("loc", ()))[-1:] == ("archivo",) for e in exc.errors()):
        return _error(422, "Falta el documento: sube un archivo de Word (.docx).")
    return _error(422, "La petición no es válida: revisa las opciones enviadas.")


def _nombre_original(nombre):
    """Solo el nombre del archivo, sin carpetas (el navegador podría mandar una ruta)."""
    nombre = (nombre or "").replace("\\", "/").rsplit("/", 1)[-1].strip()
    return nombre or NOMBRE_TEMPORAL


def _texto(valor):
    """Un campo vacío del formulario cuenta como no enviado."""
    if isinstance(valor, list):
        valor = [v.strip() for v in valor if v and v.strip()]
        return valor or None
    return valor.strip() if valor and valor.strip() else None


@app.get("/salud")
def salud():
    """Para Docker y el alojamiento: responde si el servidor está en marcha."""
    return {"estado": "bien"}


def _ip(peticion):
    conexion = peticion.client.host if peticion.client else ""
    return limites.ip_cliente(conexion, peticion.headers.get("x-forwarded-for"))


@app.get("/limite")
def limite(peticion: Request):
    """Paso 10.5: cuántos documentos puede procesar hoy quien pregunta (restantes: null si no hay límite)."""
    respuesta = {"maximo": LIMITE.maximo or None, "restantes": LIMITE.restantes(_ip(peticion))}
    if DIAGNOSTICO_IP:
        respuesta["diagnostico_ip"] = {
            "conexion": peticion.client.host if peticion.client else None,
            "x_forwarded_for": peticion.headers.get("x-forwarded-for"),
            "true_client_ip": peticion.headers.get("true-client-ip"),
            "cf_connecting_ip": peticion.headers.get("cf-connecting-ip"),
            "elegida": _ip(peticion),
        }
    return respuesta


class _Rechazo(Exception):
    """Una petición que no se procesa, con su código y el mensaje en español."""
    def __init__(self, estado, mensaje):
        super().__init__(mensaje)
        self.estado, self.mensaje = estado, mensaje


@app.exception_handler(_Rechazo)
def _rechazada(peticion, exc):
    return _error(exc.estado, exc.mensaje)


def _opciones(
    fuente_elegida: Annotated[str | None, Form(alias="fuente")] = None,
    quitar_numeracion: Annotated[bool, Form()] = False,
    titulos_tal_cual: Annotated[bool, Form()] = False,
    quitar_introduccion: Annotated[bool, Form()] = False,
    titulo_tal_cual: Annotated[bool, Form()] = False,
    mantener_portada: Annotated[bool, Form()] = False,
    quitar_indice: Annotated[bool, Form()] = False,
    tablas_tal_cual: Annotated[bool, Form()] = False,
    rotulo_tal_cual: Annotated[bool, Form()] = False,
    orden_tal_cual: Annotated[bool, Form()] = False,
    citas_tal_cual: Annotated[bool, Form()] = False,
    rotulo_apendice: Annotated[bool, Form()] = False,
    comentarios: Annotated[bool, Form()] = False,
    titulo: Annotated[str | None, Form()] = None,
    autor: Annotated[list[str] | None, Form(description="Se repite para cada autor")] = None,
    departamento: Annotated[str | None, Form()] = None,
    universidad: Annotated[str | None, Form()] = None,
    codigo_asignatura: Annotated[str | None, Form()] = None,
    asignatura: Annotated[str | None, Form()] = None,
    profesor: Annotated[list[str] | None, Form(description="Se repite si son varios")] = None,
    fecha: Annotated[str | None, Form(description="16/10/2026, 2026-10-16 o 16 de octubre de 2026")] = None,
):
    """
    Las opciones y los datos de portada del formulario, como los espera el núcleo.
    Las comparten /procesar y /procesar-texto (paso 9.7).
    """
    banderas = dict(quitar_numeracion=quitar_numeracion, titulos_tal_cual=titulos_tal_cual,
                    quitar_introduccion=quitar_introduccion, titulo_tal_cual=titulo_tal_cual,
                    mantener_portada=mantener_portada, quitar_indice=quitar_indice,
                    tablas_tal_cual=tablas_tal_cual, rotulo_tal_cual=rotulo_tal_cual,
                    orden_tal_cual=orden_tal_cual, citas_tal_cual=citas_tal_cual,
                    rotulo_apendice=rotulo_apendice, comentarios=comentarios)
    opciones = {opcion: True for opcion, marcada in banderas.items() if marcada}

    if _texto(fuente_elegida):
        elegida = fuente.familia(_texto(fuente_elegida))
        if elegida is None:
            raise _Rechazo(422, f"«{fuente_elegida.strip()}» no es una de las fuentes recomendadas por APA. "
                                f"Elige una de estas: {FUENTES}.")
        opciones["fuente"] = elegida

    sueltas = {"titulo": titulo, "autores": autor, "departamento": departamento, "universidad": universidad,
               "codigo_asignatura": codigo_asignatura, "asignatura": asignatura, "profesor": profesor,
               "fecha": fecha}
    try:
        portada = datos_portada.combinar(None, {clave: _texto(valor) for clave, valor in sueltas.items()})
    except datos_portada.ErrorDatos as e:
        raise _Rechazo(422, str(e))
    if portada:
        opciones["portada"] = portada
    return opciones


@app.post("/procesar")
def procesar(
    peticion: Request,
    archivo: Annotated[UploadFile, File(description="El trabajo en Word (.docx)")],
    opciones: Annotated[dict, Depends(_opciones)],
):
    # `def` y no `async def`: el núcleo tarda y FastAPI lo ejecuta aparte, sin bloquear otras peticiones.
    nombre = _nombre_original(archivo.filename)

    # Antes de copiar nada: la extensión, con el nombre original (el núcleo trabaja con el temporal).
    try:
        documento.comprobar_nombre(nombre)
    except ErrorDocumento as e:
        return _error(422, str(e))

    with _carpeta_temporal() as carpeta:
        entrada = Path(carpeta) / NOMBRE_TEMPORAL
        try:
            _copiar(archivo, entrada, nombre)
            seguridad.comprobar(entrada, nombre)
        except seguridad.ErrorEntrada as e:
            return _error(e.estado, str(e))
        return _procesar(carpeta, entrada, nombre, opciones, _ip(peticion))


@app.post("/procesar-texto")
def procesar_texto(
    peticion: Request,
    texto: Annotated[str, Form(description="El trabajo pegado como texto: un párrafo por línea")],
    opciones: Annotated[dict, Depends(_opciones)],
):
    """
    Paso 9.7: el texto pegado se convierte en un .docx (apa_facil/desde_texto.py)
    y se procesa igual que un archivo subido. La copia se llama texto_pegado_APA.docx.
    """
    try:
        desde_texto.comprobar(texto)
    except ErrorDocumento as e:
        return _error(422, str(e))
    with _carpeta_temporal() as carpeta:
        entrada = desde_texto.crear(texto, Path(carpeta) / NOMBRE_TEMPORAL)
        return _procesar(carpeta, entrada, desde_texto.NOMBRE, opciones, _ip(peticion))


def _procesar(carpeta, entrada, nombre, opciones, ip):
    """
    Llama al núcleo (en un proceso aparte, con tiempo máximo) y prepara la respuesta.
    Paso 10.5: antes, reserva uno de los documentos del día de esta IP; si algo sale
    mal, lo devuelve (solo cuentan los que se procesan bien).
    """
    try:
        reserva = LIMITE.reservar(ip)
    except limites.LimiteAlcanzado as e:
        return _error(429, str(e))
    try:
        respuesta = _procesar_con_turno(carpeta, entrada, nombre, opciones)
    except BaseException:
        LIMITE.devolver(reserva)
        raise
    if isinstance(respuesta, JSONResponse):   # un error
        LIMITE.devolver(reserva)
        return respuesta
    respuesta["limite"] = {"maximo": LIMITE.maximo or None, "restantes": LIMITE.restantes(ip)}
    return respuesta


def _procesar_con_turno(carpeta, entrada, nombre, opciones):
    copia = f"{Path(nombre).stem}{documento.SUFIJO}.docx"
    inicio = time.perf_counter()
    if not TURNOS.acquire(timeout=ESPERA_MAXIMA):
        return _error(503, "Ahora mismo hay muchos documentos en cola. Inténtalo de nuevo dentro de un momento.")
    try:
        resultado = trabajo.procesar(entrada, Path(carpeta) / "salida", opciones, nombre, copia, TIEMPO_MAXIMO)
    except trabajo.TiempoAgotado:
        return _error(504, f"«{nombre}» ha tardado demasiado en procesarse (más de {TIEMPO_MAXIMO:g} "
                           "segundos) y se ha detenido. No se ha guardado nada. Si es muy largo, prueba "
                           "a dividirlo en partes.")
    finally:
        TURNOS.release()

    if resultado[0] == "error":
        return _error(422, resultado[1].replace(NOMBRE_TEMPORAL, nombre))
    if resultado[0] == "fallo":   # sin el mensaje: podría llevar texto del documento
        _, tipo, marcos = resultado
        donde = f", en {marcos[-1][0]}:{marcos[-1][1]}" if marcos else ""
        REGISTRO.error("Error inesperado en el núcleo (%s%s).", tipo, donde)
        errores.avisar_fallo_nucleo(tipo, marcos)
        return _error(500, _MENSAJE_FALLO)
    _, salida, informe, resumen, coste = resultado
    contenido = Path(salida).read_bytes()
    # Paso 10.5: lo que ha costado, sin el nombre ni el texto (paso 8.3).
    memoria = f"{coste['memoria_mb']:.0f} MB" if coste["memoria_mb"] is not None else "desconocida"
    REGISTRO.info("Documento procesado en %.1f s (%.1f s del núcleo; memoria máxima del proceso: %s; %d KB).",
                  time.perf_counter() - inicio, coste["segundos"], memoria, Path(entrada).stat().st_size // 1024)
    return {"informe": informe, "resumen": resumen,
            "documento": {"nombre": copia, "contenido": base64.b64encode(contenido).decode("ascii")}}


@contextlib.contextmanager
def _carpeta_temporal():
    """
    Paso 8.3: una carpeta con nombre aleatorio para el archivo, su copia y su informe,
    que se borra al terminar pase lo que pase. En Windows, un archivo puede seguir
    bloqueado un instante después de parar el proceso del núcleo: se reintenta.
    """
    carpeta = Path(tempfile.mkdtemp(prefix="apa_", dir=CARPETA_TEMPORAL))
    try:
        yield carpeta
    finally:
        for intento in range(10):
            shutil.rmtree(carpeta, ignore_errors=True)
            if not carpeta.exists():
                break
            time.sleep(0.2 * (intento + 1))
        else:
            REGISTRO.error("No se ha podido borrar la carpeta temporal %s.", carpeta.name)


def _copiar(archivo, destino, nombre):
    """Copia el archivo subido a `destino`, sin pasar del tamaño máximo."""
    copiado = 0
    with destino.open("wb") as f:
        while bloque := archivo.file.read(1024 * 1024):
            copiado += len(bloque)
            if copiado > seguridad.TAMANO_MAXIMO:
                raise seguridad.demasiado_grande(nombre)
            f.write(bloque)
