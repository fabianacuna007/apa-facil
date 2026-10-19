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

El archivo se guarda en una carpeta temporal con un nombre fijo (nunca el del
estudiante: «con.docx» o «com1.docx» son nombres reservados de Windows), y la
carpeta se borra al terminar, vaya bien o mal.

Seguridad (paso 8.2): solo .docx de verdad, 10 MB como máximo, nada de bombas
ZIP (servidor/seguridad.py) y un tiempo máximo por documento: el núcleo trabaja
en un proceso aparte que se para si se pasa (servidor/trabajo.py). Como mucho se
procesan SIMULTANEOS documentos a la vez; los demás esperan su turno.
"""

import base64
import os
import tempfile
import threading
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

from apa_facil import datos_portada, documento
from apa_facil.documento import ErrorDocumento
from apa_facil.reglas import fuente
from servidor import seguridad, trabajo

NOMBRE_TEMPORAL = "documento.docx"
TIEMPO_MAXIMO = float(os.environ.get("APA_TIEMPO_MAXIMO", 120))   # segundos por documento
SIMULTANEOS = int(os.environ.get("APA_SIMULTANEOS", 2))            # documentos procesándose a la vez
ESPERA_MAXIMA = 60                                                 # segundos en cola antes de rendirse
TURNOS = threading.BoundedSemaphore(SIMULTANEOS)
# Lo que puede ocupar la petición entera: el archivo y, como mucho, 1 MB de opciones y cabeceras.
PETICION_MAXIMA = seguridad.TAMANO_MAXIMO + seguridad.MB
FUENTES = ", ".join(list(fuente.RECOMENDADAS)[:-1]) + " o " + list(fuente.RECOMENDADAS)[-1]

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
        try:
            await siguiente(scope, contar, send)
        except _PeticionGrande:
            await _error(413, _MENSAJE_GRANDE)(scope, receive, send)
    return asgi


_MENSAJE_GRANDE = (f"El archivo ocupa más de {seguridad.TAMANO_MAXIMO // seguridad.MB} MB, el máximo que se "
                   "admite. Si tiene muchas imágenes, comprímelas en Word (Formato de imagen › Comprimir "
                   "imágenes) y vuelve a subirlo.")
app.add_middleware(_limitar_tamano)


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


@app.post("/procesar")
def procesar(
    archivo: Annotated[UploadFile, File(description="El trabajo en Word (.docx)")],
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
    # `def` y no `async def`: el núcleo tarda y FastAPI lo ejecuta aparte, sin bloquear otras peticiones.
    nombre = _nombre_original(archivo.filename)
    banderas = dict(quitar_numeracion=quitar_numeracion, titulos_tal_cual=titulos_tal_cual,
                    quitar_introduccion=quitar_introduccion, titulo_tal_cual=titulo_tal_cual,
                    mantener_portada=mantener_portada, quitar_indice=quitar_indice,
                    tablas_tal_cual=tablas_tal_cual, rotulo_tal_cual=rotulo_tal_cual,
                    orden_tal_cual=orden_tal_cual, rotulo_apendice=rotulo_apendice, comentarios=comentarios)
    opciones = {opcion: True for opcion, marcada in banderas.items() if marcada}

    if _texto(fuente_elegida):
        elegida = fuente.familia(_texto(fuente_elegida))
        if elegida is None:
            return _error(422, f"«{fuente_elegida.strip()}» no es una de las fuentes recomendadas por APA. "
                               f"Elige una de estas: {FUENTES}.")
        opciones["fuente"] = elegida

    sueltas = {"titulo": titulo, "autores": autor, "departamento": departamento, "universidad": universidad,
               "codigo_asignatura": codigo_asignatura, "asignatura": asignatura, "profesor": profesor,
               "fecha": fecha}
    try:
        portada = datos_portada.combinar(None, {clave: _texto(valor) for clave, valor in sueltas.items()})
    except datos_portada.ErrorDatos as e:
        return _error(422, str(e))
    if portada:
        opciones["portada"] = portada

    # Antes de copiar nada: la extensión, con el nombre original (el núcleo trabaja con el temporal).
    try:
        documento.comprobar_nombre(nombre)
    except ErrorDocumento as e:
        return _error(422, str(e))

    with tempfile.TemporaryDirectory(prefix="apa_") as carpeta:
        entrada = Path(carpeta) / NOMBRE_TEMPORAL
        try:
            _copiar(archivo, entrada, nombre)
            seguridad.comprobar(entrada, nombre)
        except seguridad.ErrorEntrada as e:
            return _error(e.estado, str(e))

        copia = f"{Path(nombre).stem}{documento.SUFIJO}.docx"
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
        if resultado[0] == "fallo":   # sin detalles: podrían llevar texto del documento
            return _error(500, "Ha habido un error inesperado al procesar el documento. "
                               "No se ha guardado nada. Inténtalo de nuevo o prueba con otro archivo.")
        _, salida, informe, resumen = resultado
        contenido = Path(salida).read_bytes()

    return {"informe": informe, "resumen": resumen,
            "documento": {"nombre": copia, "contenido": base64.b64encode(contenido).decode("ascii")}}


def _copiar(archivo, destino, nombre):
    """Copia el archivo subido a `destino`, sin pasar del tamaño máximo."""
    copiado = 0
    with destino.open("wb") as f:
        while bloque := archivo.file.read(1024 * 1024):
            copiado += len(bloque)
            if copiado > seguridad.TAMANO_MAXIMO:
                raise seguridad.demasiado_grande(nombre)
            f.write(bloque)
