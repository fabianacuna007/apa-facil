"""
Paso 8.1: el servidor (POST /procesar) da lo mismo que la línea de comandos.
Se llama sin arrancarlo, con el cliente de pruebas de FastAPI.
"""

import base64
import io
from pathlib import Path

import docx
import pytest
from fastapi.testclient import TestClient

from apa_facil import documento, pipeline
from servidor.app import app

RAIZ = Path(__file__).resolve().parent.parent
PRUEBAS = sorted((RAIZ / "pruebas").glob("*.docx"))
cliente = TestClient(app)


def enviar(ruta, nombre=None, **campos):
    with open(ruta, "rb") as f:
        return cliente.post("/procesar", data=campos, files={"archivo": (nombre or ruta.name, f)})


@pytest.mark.parametrize("ruta", PRUEBAS, ids=lambda r: r.name)
def test_procesa_igual_que_la_linea_de_comandos(ruta, tmp_path):
    respuesta = enviar(ruta)
    assert respuesta.status_code == 200, respuesta.text
    datos = respuesta.json()
    assert datos["documento"]["nombre"] == f"{ruta.stem}_APA.docx"
    assert datos["informe"]["archivo"] == ruta.name
    assert datos["informe"]["copia"] == f"{ruta.stem}_APA.docx"

    # La copia que devuelve se abre y es la misma que guarda la línea de comandos
    # (el pipeline ya ha comprobado que conserva el texto del original).
    copia = docx.Document(io.BytesIO(base64.b64decode(datos["documento"]["contenido"])))
    original = documento.abrir(ruta)
    assert len(copia.tables) == len(original.tables)
    assert len(copia.inline_shapes) == len(original.inline_shapes)

    # El informe es el mismo que da la línea de comandos (salvo la fecha).
    _, informe = pipeline.procesar(ruta, tmp_path)
    esperado, recibido = informe.a_dict(), datos["informe"]
    for clave in ("cambios", "avisos", "errores", "resumen"):
        assert recibido[clave] == esperado[clave]
    assert datos["resumen"] == informe.como_texto()


def test_las_opciones_llegan_al_nucleo():
    datos = enviar(PRUEBAS[1], fuente="arial", comentarios="true", quitar_indice="true",
                   titulo="Mi trabajo", autor=["Ana Pérez", "Luis Gómez"], fecha="16/10/2026").json()
    opciones = datos["informe"]["opciones"]
    assert opciones["fuente"] == "Arial"
    assert opciones["comentarios"] and opciones["quitar_indice"] and not opciones["mantener_portada"]
    assert "titulo" in opciones["portada"]["dados"] and "autores" in opciones["portada"]["dados"]


def test_campos_vacios_cuentan_como_no_enviados():
    datos = enviar(PRUEBAS[1], fuente="", titulo="", fecha="").json()
    assert datos["informe"]["opciones"]["fuente"] is None
    assert datos["informe"]["opciones"]["portada"] is None


@pytest.mark.parametrize("campos, palabra", [({"fuente": "Comic Sans"}, "fuentes recomendadas"),
                                             ({"fecha": "mañana"}, "fecha")])
def test_opciones_no_validas(campos, palabra):
    respuesta = enviar(PRUEBAS[1], **campos)
    assert respuesta.status_code == 422
    assert palabra in respuesta.json()["error"]


def test_archivo_que_no_es_word(tmp_path):
    pdf = tmp_path / "trabajo.pdf"
    pdf.write_bytes(b"%PDF-1.4 no es un Word")
    respuesta = enviar(pdf)
    assert respuesta.status_code == 422
    assert respuesta.json()["error"] == "«trabajo.pdf» no es un documento de Word. Solo se admiten archivos .docx."

    renombrado = tmp_path / "trabajo.docx"
    renombrado.write_bytes(pdf.read_bytes())
    error = enviar(renombrado).json()["error"]
    assert "«trabajo.docx» está dañado" in error and "documento.docx" not in error


def test_nombre_reservado_de_windows():
    respuesta = enviar(PRUEBAS[1], nombre="con.docx")
    assert respuesta.status_code == 200
    assert respuesta.json()["documento"]["nombre"] == "con_APA.docx"


def test_falta_el_archivo():
    respuesta = cliente.post("/procesar", data={"fuente": "Arial"})
    assert respuesta.status_code == 422
    assert respuesta.json() == {"error": "Falta el documento: sube un archivo de Word (.docx)."}
