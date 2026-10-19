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
"""

import base64
import tempfile
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

from apa_facil import datos_portada, documento, pipeline
from apa_facil.documento import ErrorDocumento
from apa_facil.reglas import fuente

NOMBRE_TEMPORAL = "documento.docx"
FUENTES = ", ".join(list(fuente.RECOMENDADAS)[:-1]) + " o " + list(fuente.RECOMENDADAS)[-1]

app = FastAPI(title="APA Fácil", description="Aplica el formato APA 7 de estudiante a un documento Word.")


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

    # La extensión se comprueba con el nombre original; el núcleo trabaja con el temporal.
    extension = Path(nombre).suffix.lower()
    with tempfile.TemporaryDirectory(prefix="apa_") as carpeta:
        entrada = Path(carpeta) / (NOMBRE_TEMPORAL if extension == ".docx" else "documento" + extension)
        with entrada.open("wb") as destino:
            while bloque := archivo.file.read(1024 * 1024):
                destino.write(bloque)
        try:
            salida, informe = pipeline.procesar(entrada, Path(carpeta) / "salida", opciones)
        except ErrorDocumento as e:
            return _error(422, str(e).replace(entrada.name, nombre))
        except Exception:   # sin detalles: podrían llevar texto del documento
            return _error(500, "Ha habido un error inesperado al procesar el documento. "
                               "No se ha guardado nada. Inténtalo de nuevo o prueba con otro archivo.")
        contenido = salida.read_bytes()

    copia = f"{Path(nombre).stem}{documento.SUFIJO}.docx"
    informe.archivo, informe.copia = nombre, copia
    return {"informe": informe.a_dict(),
            "resumen": informe.como_texto(),
            "documento": {"nombre": copia, "contenido": base64.b64encode(contenido).decode("ascii")}}
