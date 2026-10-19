"""
Fase 9: la web repite algunas listas del núcleo y del servidor (las fuentes y las
casillas de opciones). Estas pruebas avisan si dejan de coincidir.
"""

import inspect
import re
from pathlib import Path

from apa_facil.reglas import fuente
from servidor import app as servidor_app

OPCIONES_WEB = (Path(__file__).resolve().parent.parent / "web" / "lib" / "opciones.ts").read_text(encoding="utf-8")


def test_las_fuentes_de_la_web_son_las_del_nucleo():
    fuentes = re.findall(r'\["([^"]+)", (\d+)\]', OPCIONES_WEB)
    assert [(nombre, int(tamano)) for nombre, tamano in fuentes] == list(fuente.RECOMENDADAS.items())


def test_las_casillas_de_la_web_son_las_opciones_del_servidor():
    casillas = re.findall(r'clave: "(\w+)", bandera: "(--[\w-]+)"', OPCIONES_WEB)
    parametros = inspect.signature(servidor_app._opciones).parameters
    del_servidor = {nombre for nombre, p in parametros.items() if p.annotation.__origin__ is bool}
    assert {clave for clave, _ in casillas} == del_servidor
    for clave, bandera in casillas:
        assert bandera == "--" + clave.replace("_", "-")
