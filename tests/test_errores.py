"""
Paso 11.3: los avisos de errores con Sentry (servidor/errores.py) no llevan nada
del documento: ni su texto, ni el nombre del archivo, ni la IP, ni la petición.
Los eventos no salen del equipo: se recogen con un transporte de prueba.
"""

import json
import multiprocessing
from pathlib import Path

import pytest
import sentry_sdk
from fastapi.testclient import TestClient
from sentry_sdk.transport import Transport

from apa_facil import pipeline
from servidor import app as servidor_app, errores, limites, seguridad, trabajo
from servidor.app import app

DSN = "https://clave@o0.ingest.de.sentry.io/0"   # inventada: con el transporte de prueba no se usa
TEXTO = "La motivación intrínseca de Ana Pérez en el aula"
NOMBRE = "TFG de Ana Pérez.docx"
DOCUMENTO = Path(__file__).parent.parent / "pruebas" / "05_tfg_completo.docx"
cliente = TestClient(app)


class Recoger(Transport):
    def __init__(self, opciones=None):
        super().__init__(opciones)
        self.eventos = []

    def capture_envelope(self, sobre):
        evento = sobre.get_event()
        if evento is not None:
            self.eventos.append(evento)


@pytest.fixture
def sentry(monkeypatch):
    monkeypatch.setattr(servidor_app, "LIMITE", limites.Limite(maximo=0))
    transporte = Recoger()
    assert errores.activar(DSN, transporte)
    yield transporte.eventos
    sentry_sdk.get_client().close()
    sentry_sdk.init()   # sin clave: inactivo para las demás pruebas


def enviar(ruta, nombre=NOMBRE):
    with open(ruta, "rb") as f:
        return cliente.post("/procesar", files={"archivo": (nombre, f.read())})


def sin_datos_privados(evento):
    todo = json.dumps(evento, ensure_ascii=False)
    for privado in (TEXTO, "Ana Pérez", "testclient", "x-forwarded-for", "multipart", "C:\\\\Users"):
        assert privado not in todo, privado


def test_sin_clave_no_se_activa(monkeypatch):
    monkeypatch.delenv("SENTRY_DSN", raising=False)
    assert errores.activar() is False
    assert not sentry_sdk.get_client().is_active()


def test_un_error_del_servidor_llega_sin_texto_ni_peticion(sentry, monkeypatch):
    def falla(ruta, nombre):
        texto_del_documento = TEXTO   # una variable local: tampoco puede salir
        raise RuntimeError(f"fallo con {texto_del_documento} en {nombre}")
    monkeypatch.setattr(seguridad, "comprobar", falla)
    respuesta = enviar(DOCUMENTO)
    assert respuesta.status_code == 500
    sentry_sdk.flush()
    assert len(sentry) == 1
    evento = sentry[0]
    sin_datos_privados(evento)
    assert set(evento) <= {"event_id", "timestamp", "platform", "level", "release", "environment", "sdk",
                           "server_name", "contexts", "exception"}
    error = evento["exception"]["values"][-1]
    assert error["type"] == "RuntimeError" and error["value"] == errores.OMITIDO
    funciones = [m["function"] for m in error["stacktrace"]["frames"]]
    assert "falla" in funciones   # se sabe dónde ha fallado
    assert all("vars" not in m for m in error["stacktrace"]["frames"])


def test_un_error_del_nucleo_llega_con_sus_sitios(sentry, monkeypatch):
    def procesar(*args):
        return ("fallo", "KeyError", [("servidor/trabajo.py", 63, "_trabajar"), ("apa_facil/pipeline.py", 10, "procesar"),
                                      ("docx/oxml/xmlchemy.py", 5, "get")])
    monkeypatch.setattr(trabajo, "procesar", procesar)
    assert enviar(DOCUMENTO).status_code == 500
    sentry_sdk.flush()
    assert len(sentry) == 1
    error = sentry[0]["exception"]["values"][0]
    assert error["type"] == "KeyError" and error["value"] == errores.OMITIDO
    assert [(m["filename"], m["lineno"], m["in_app"]) for m in error["stacktrace"]["frames"]] == [
        ("servidor/trabajo.py", 63, True), ("apa_facil/pipeline.py", 10, True), ("docx/oxml/xmlchemy.py", 5, False)]
    sin_datos_privados(sentry[0])


def test_los_errores_previstos_no_se_avisan(sentry):
    # Un archivo dañado es un 422 con su mensaje: no es un fallo del programa.
    assert cliente.post("/procesar", files={"archivo": (NOMBRE, b"PK no es un zip")}).status_code == 422
    sentry_sdk.flush()
    assert sentry == []


def test_limpiar_descarta_lo_que_no_es_un_error():
    assert errores.limpiar({"message": TEXTO, "level": "info"}) is None
    evento = errores.limpiar({
        "level": "error", "user": {"ip_address": "1.2.3.4"}, "request": {"data": TEXTO},
        "extra": {"x": TEXTO}, "breadcrumbs": {"values": [{"message": TEXTO}]}, "tags": {"t": TEXTO},
        "contexts": {"runtime": {"name": "CPython", "version": "3.11"}, "os": {"name": TEXTO}},
        "exception": {"values": [{"type": "ValueError", "value": TEXTO, "stacktrace": {"frames": [
            {"filename": "apa_facil/x.py", "lineno": 1, "function": "f", "vars": {"t": TEXTO},
             "abs_path": "C:\\Users\\Ana Pérez\\x.py"}]}}]},
    })
    sin_datos_privados(evento)
    assert evento["contexts"] == {"runtime": {"name": "CPython", "version": "3.11"}}
    assert evento["exception"]["values"][0]["stacktrace"]["frames"] == [
        {"filename": "apa_facil/x.py", "lineno": 1, "function": "f"}]


def test_el_nucleo_devuelve_tipo_y_sitios_sin_mensaje(monkeypatch, tmp_path):
    def falla(*args):
        raise ValueError(TEXTO)
    monkeypatch.setattr(pipeline, "procesar", falla)
    recibir, enviar_ = multiprocessing.Pipe(duplex=False)
    trabajo._trabajar("entrada.docx", str(tmp_path), {}, NOMBRE, "copia.docx", enviar_)
    estado, tipo, marcos = recibir.recv()
    assert (estado, tipo) == ("fallo", "ValueError")
    assert marcos[0] == ("servidor/trabajo.py", marcos[0][1], "_trabajar")
    assert marcos[-1][0] == "tests/test_errores.py" and marcos[-1][2] == "falla"
    assert TEXTO not in repr(marcos)
