"""
Paso 2.3: informe estructurado.

- Cada entrada sabe qué regla la produjo (sin que las reglas lo digan).
- Junto a cada copia se guarda el informe en JSON, que la web podrá leer.
- El mismo informe se puede mostrar como texto en la terminal.
"""

import json
from datetime import datetime
from pathlib import Path

import pytest

from apa_facil import documento, pipeline
from apa_facil.documento import ErrorDocumento, w
from apa_facil.informe import NOMBRES, VERSION, Informe

CARPETA = Path(__file__).resolve().parent.parent / "pruebas"


def original(prefijo):
    return next(CARPETA.glob(f"{prefijo}_*.docx"))


@pytest.fixture(scope="module")
def procesado_03(tmp_path_factory):
    salida, informe = pipeline.procesar(original("03"), tmp_path_factory.mktemp("salida"))
    return salida, informe, json.loads(salida.with_suffix(".json").read_text(encoding="utf-8"))


# --- Cada entrada con su regla ---------------------------------------------------------

def test_cada_mensaje_sabe_de_que_regla_viene(procesado_03):
    _, informe, _ = procesado_03
    reglas = {e.mensaje: e.regla for e in informe.entradas}
    encabezado = next(m for m in informe.avisos if m.startswith("El encabezado contenía"))
    assert reglas[encabezado] == "paginacion"
    assert reglas[next(m for m in informe.cambios if m.startswith("Márgenes"))] == "margenes"
    assert reglas[next(m for m in informe.cambios if m.startswith("Fuente"))] == "fuente"
    assert all(e.regla in NOMBRES for e in informe.entradas)


def test_un_error_tambien_sabe_su_regla(tmp_path):
    _, informe = pipeline.procesar(original("01"), tmp_path, {"fuente": "Papyrus"})
    assert [e.regla for e in informe.entradas if e.tipo == "error"] == ["fuente"]


def test_las_listas_de_siempre_siguen_funcionando():
    informe = Informe()
    informe.regla = "margenes"
    informe.cambio("Uno.")
    informe.aviso("Revisa esto.")
    informe.regla = None
    informe.error("Esto no se pudo.")
    assert (informe.cambios, informe.avisos, informe.errores) == (["Uno."], ["Revisa esto."], ["Esto no se pudo."])
    assert informe.a_dict()["errores"] == [{"regla": None, "nombre": "General", "mensaje": "Esto no se pudo."}]


# --- JSON ------------------------------------------------------------------------------

def test_el_json_se_guarda_junto_a_la_copia(procesado_03):
    salida, _, datos = procesado_03
    assert salida.with_suffix(".json").exists()
    assert datos["archivo"] == original("03").name
    assert datos["copia"] == salida.name
    assert datos["version"] == VERSION
    assert datos["opciones"] == {"fuente": None, "quitar_numeracion": False, "titulos_tal_cual": False,
                                 "quitar_introduccion": False,
                                 "titulo_tal_cual": False, "mantener_portada": False, "quitar_indice": False,
                                 "tablas_tal_cual": False, "rotulo_tal_cual": False,
                                 "portada": None}
    datetime.fromisoformat(datos["fecha"])   # una fecha válida


def test_el_json_coincide_con_el_informe(procesado_03):
    _, informe, datos = procesado_03
    for clave, mensajes in (("cambios", informe.cambios), ("avisos", informe.avisos), ("errores", informe.errores)):
        assert [e["mensaje"] for e in datos[clave]] == mensajes
        assert datos["resumen"][clave] == len(mensajes)
    aviso = next(e for e in datos["avisos"] if e["mensaje"].startswith("El encabezado contenía"))
    assert aviso["regla"] == "paginacion" and aviso["nombre"] == "Número de página"
    assert "Facultad de Ciencias" in aviso["mensaje"]


def test_las_tildes_se_guardan_tal_cual(procesado_03):
    salida, _, _ = procesado_03
    texto = salida.with_suffix(".json").read_text(encoding="utf-8")
    assert "Márgenes" in texto and "\\u00e1" not in texto


def test_el_json_recoge_la_fuente_pedida(tmp_path):
    salida, _ = pipeline.procesar(original("01"), tmp_path, {"fuente": "Georgia"})
    datos = json.loads(salida.with_suffix(".json").read_text(encoding="utf-8"))
    assert datos["opciones"] == {"fuente": "Georgia", "quitar_numeracion": False, "titulos_tal_cual": False,
                                 "quitar_introduccion": False,
                                 "titulo_tal_cual": False, "mantener_portada": False, "quitar_indice": False,
                                 "tablas_tal_cual": False, "rotulo_tal_cual": False,
                                 "portada": None}


def test_sin_copia_no_hay_informe(tmp_path, monkeypatch):
    # Un informe de una vez anterior tampoco debe quedarse: no correspondería a ninguna copia.
    viejo = tmp_path / f"{original('01').stem}_APA.json"
    viejo.write_text("{}", encoding="utf-8")
    guardar_de_verdad = documento.guardar

    def guardar_estropeando(doc, ruta):
        texto = next(t for t in doc.element.body.iter(w("t")) if (t.text or "").strip())
        texto.text = "texto alterado"
        guardar_de_verdad(doc, ruta)

    monkeypatch.setattr(documento, "guardar", guardar_estropeando)
    with pytest.raises(ErrorDocumento):
        pipeline.procesar(original("01"), tmp_path)
    assert list(tmp_path.iterdir()) == []


# --- Texto para la terminal ---------------------------------------------------------------

def test_texto_con_los_tres_grupos():
    informe = Informe()
    informe.cambio("Uno.")
    informe.aviso("Revisa esto.")
    informe.error("Esto no se pudo.")
    assert informe.como_texto().splitlines() == [
        "   Lo que hemos cambiado (1):", "     · Uno.",
        "   Lo que debes revisar tú (1):", "     · Revisa esto.",
        "   Lo que no se ha podido hacer (1):", "     · Esto no se pudo.",
    ]


def test_texto_sin_cambios():
    assert Informe().como_texto().strip() == "No ha hecho falta cambiar nada: el formato ya cumplía APA."
