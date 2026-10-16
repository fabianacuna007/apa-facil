"""
Comprueba que todas las piezas del esqueleto (paso 0.3) existen y se
pueden cargar. Las reglas de la fase 1 ya están programadas y tienen sus pruebas.
"""

import importlib
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent

MODULOS = [
    "apa_facil",
    "apa_facil.__main__",
    "apa_facil.documento",
    "apa_facil.informe",
    "apa_facil.pipeline",
    "apa_facil.estructura",
    "apa_facil.revision",
]

REGLAS = ["margenes", "fuente", "interlineado", "alineacion", "sangria", "paginacion"]


@pytest.mark.parametrize("nombre", MODULOS)
def test_modulo_se_carga(nombre):
    importlib.import_module(nombre)


@pytest.mark.parametrize("nombre", REGLAS)
def test_regla_se_carga(nombre):
    assert callable(importlib.import_module(f"apa_facil.reglas.{nombre}").aplicar)


@pytest.mark.parametrize("ruta", ["herramientas/inspeccionar.py", "herramientas/verificar.py",
                                  "pruebas", "tests"])
def test_existe(ruta):
    assert (RAIZ / ruta).exists()
