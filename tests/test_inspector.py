"""
Paso 0.6: el inspector muestra el formato real de cada documento.

Se comprueban datos que conocemos de los documentos de prueba, sobre todo
los que solo se ven recorriendo las tres capas de formato de Word.
"""

from pathlib import Path

import docx
import pytest

from herramientas.inspeccionar import inspeccionar, main, mostrar

CARPETA = Path(__file__).resolve().parent.parent / "pruebas"
PRUEBAS = sorted(CARPETA.glob("*.docx"))


def datos(nombre):
    return inspeccionar(CARPETA / nombre)


def fuentes(d):
    return {f for f, _ in d["fuentes"]}


def textos_de(d, tipo):
    return [t for e in d["encabezados"] if e["tipo"] == tipo for t in e["texto"]]


@pytest.mark.parametrize("ruta", PRUEBAS, ids=lambda p: p.stem)
def test_funciona_con_los_cinco(ruta):
    texto = mostrar(inspeccionar(ruta))
    for apartado in ("Márgenes", "Fuentes", "Interlineado", "Sangría", "Alineación",
                     "Estilos de párrafo", "Encabezados y pies"):
        assert apartado in texto


def test_01_arial_10_desde_los_valores_por_defecto():
    # Ningún párrafo dice "Arial": viene de la capa más profunda.
    principal = datos("01_ensayo_basico.docx")["fuentes"].most_common(1)[0][0]
    assert principal == ("Arial", 10)


def test_01_margenes_y_justificado():
    d = datos("01_ensayo_basico.docx")
    assert d["secciones"][0]["margenes"]["superior"] == "1,50"
    assert d["alineaciones"]["justificada"] == 8


def test_02_interlineado_1_15_en_la_mayoria():
    assert datos("02_informe_tablas.docx")["interlineados"].most_common(1)[0][0] == "1,15"


def test_02_numero_de_pagina_en_el_pie():
    assert textos_de(datos("02_informe_tablas.docx"), "Pie") == ["[nº de página]"]


def test_03_encabezado_con_texto_y_pie_con_numero():
    d = datos("03_trabajo_imagenes_listas.docx")
    assert textos_de(d, "Encabezado") == ["Facultad de Ciencias · Medio Ambiente y Sostenibilidad"]
    assert textos_de(d, "Pie") == ["- [nº de página] -"]


def test_03_sangria_de_un_centimetro():
    assert datos("03_trabajo_imagenes_listas.docx")["sangrias"]["primera línea 1,00 cm"] == 6


def test_04_muchas_fuentes_y_una_seccion_horizontal():
    d = datos("04_formato_caotico.docx")
    assert {"Comic Sans MS", "Courier New", "Verdana", "Arial Black"} <= fuentes(d)
    assert [s["orientacion"] for s in d["secciones"]] == ["vertical", "horizontal", "vertical"]


def test_05_pie_con_pagina_de_total():
    assert textos_de(datos("05_tfg_completo.docx"), "Pie") == ["Página [nº de página] de [total de páginas]"]


def test_05_estilos_de_titulo():
    assert {"Heading 1", "Heading 2", "Heading 3", "Heading 4"} <= set(datos("05_tfg_completo.docx")["estilos"])


def test_fuentes_del_tema(tmp_path):
    # Los Word reales suelen indicar la fuente como "la del tema"; hay que traducirla.
    doc = docx.Document()  # la plantilla de python-docx usa Cambria y Calibri del tema
    doc.add_paragraph("Texto normal")
    ruta = tmp_path / "tema.docx"
    doc.save(str(ruta))
    assert fuentes(inspeccionar(ruta)) == {"Cambria"}


def test_linea_de_comandos(capsys, tmp_path):
    assert main([str(PRUEBAS[0])]) == 0
    assert "01_ensayo_basico.docx" in capsys.readouterr().out
    assert main([str(tmp_path / "no_existe.docx")]) == 1
    assert "No se encuentra" in capsys.readouterr().out
    assert main([]) == 2
