"""
Comprueba que todas las piezas del esqueleto (paso 0.3) existen y se
pueden cargar, y que las reglas sin programar avisan en vez de no hacer nada.
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
def test_regla_sin_programar_da_error(nombre):
    regla = importlib.import_module(f"apa_facil.reglas.{nombre}")
    with pytest.raises(NotImplementedError, match="se programa en el paso 1"):
        regla.aplicar(doc=None, opciones={}, informe=None)


@pytest.mark.parametrize("ruta", ["herramientas/inspeccionar.py", "herramientas/verificar.py",
                                  "pruebas", "tests"])
def test_existe(ruta):
    assert (RAIZ / ruta).exists()


def test_linea_de_comandos_arranca(capsys):
    from apa_facil.__main__ import main
    assert main() == 0
    assert "en construcción" in capsys.readouterr().out
