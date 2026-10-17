"""
Paso 2.4: revisión visual en PDF.

La elección del conversor, los mensajes y LibreOffice se prueban con
conversores simulados, así que funcionan en cualquier equipo. La conversión
real con Word solo se ejecuta si Word está instalado (si no, se omite).
"""

import subprocess
from pathlib import Path

import pytest

from apa_facil import pipeline
from apa_facil.documento import ErrorDocumento
from herramientas import a_pdf

CARPETA = Path(__file__).resolve().parent.parent / "pruebas"
PDF_MINIMO = b"%PDF-1.4\n1 0 obj << /Type /Pages /Kids [2 0 R] /Count 1 >> endobj\n2 0 obj << /Type /Page >> endobj\n"


def original(prefijo):
    return next(CARPETA.glob(f"{prefijo}_*.docx"))


# --- Contar páginas ---------------------------------------------------------------------

@pytest.mark.parametrize("contenido, paginas", [
    (b"<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>", 3),
    (b"<< /Count 4 /Kids [] /Type /Pages >>", 4),                            # en otro orden
    (b"<< /Type /Outlines /Count 30 >> << /Type /Pages /Count 3 >>", 3),     # los marcadores no cuentan
    (b"<< /Type /Page >> << /Type /Page >>", 2),                             # sin total: una a una
    (b"nada", None),
])
def test_paginas_del_pdf(tmp_path, contenido, paginas):
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(contenido)
    assert a_pdf.paginas_del_pdf(pdf) == paginas


# --- Elegir el conversor ------------------------------------------------------------------

def test_word_tiene_preferencia(monkeypatch):
    monkeypatch.setattr(a_pdf, "hay_word", lambda: True)
    monkeypatch.setattr(a_pdf, "buscar_libreoffice", lambda: Path("soffice"))
    assert a_pdf.buscar_conversor() == ("Word", None)


def test_libreoffice_si_no_hay_word(monkeypatch):
    monkeypatch.setattr(a_pdf, "hay_word", lambda: False)
    monkeypatch.setattr(a_pdf, "buscar_libreoffice", lambda: Path("soffice"))
    assert a_pdf.buscar_conversor() == ("LibreOffice", Path("soffice"))


def test_sin_conversor_lo_explica(monkeypatch, capsys):
    monkeypatch.setattr(a_pdf, "hay_word", lambda: False)
    monkeypatch.setattr(a_pdf, "buscar_libreoffice", lambda: None)
    with pytest.raises(ErrorDocumento, match="LibreOffice es gratuito"):
        a_pdf.convertir([original("01")])
    assert a_pdf.main([str(original("01"))]) == 1
    assert "hace falta Word o LibreOffice" in capsys.readouterr().out


def test_sin_archivos(capsys):
    assert a_pdf.main([]) == 2
    assert "Indica qué documentos convertir" in capsys.readouterr().out


# --- LibreOffice (simulado) ------------------------------------------------------------

@pytest.fixture
def libreoffice_simulado(monkeypatch):
    """Un LibreOffice de mentira: crea el PDF en --outdir y apunta cómo se le llamó."""
    llamadas = []

    def run(orden, **_):
        llamadas.append(orden)
        docx, carpeta = Path(orden[-1]), Path(orden[orden.index("--outdir") + 1])
        (carpeta / f"{docx.stem}.pdf").write_bytes(PDF_MINIMO)
        return subprocess.CompletedProcess(orden, 0, b"", b"")

    monkeypatch.setattr(a_pdf.subprocess, "run", run)
    return llamadas


def test_libreoffice_crea_el_pdf_en_la_carpeta_de_salida(tmp_path, libreoffice_simulado):
    antes = sorted(CARPETA.iterdir())
    nombre, [conversion] = a_pdf.convertir([original("01")], tmp_path, ("LibreOffice", Path("soffice")))
    assert nombre == "LibreOffice"
    assert conversion.error is None and conversion.paginas == 1
    assert conversion.pdf == (tmp_path / f"{original('01').stem}.pdf").resolve()
    assert sorted(CARPETA.iterdir()) == antes, "no se escribe nada en pruebas/"
    orden = libreoffice_simulado[0]
    assert orden[:5] == ["soffice", "--headless", "--norestore", "--convert-to", "pdf"]


def test_libreoffice_que_falla(tmp_path, monkeypatch):
    monkeypatch.setattr(a_pdf.subprocess, "run",
                        lambda orden, **_: subprocess.CompletedProcess(orden, 1, b"", b"Error: fuente rota"))
    _, [conversion] = a_pdf.convertir([original("01")], tmp_path, ("LibreOffice", Path("soffice")))
    assert conversion.pdf is None
    assert "LibreOffice no ha podido convertirlo (Error: fuente rota)" in conversion.error


def test_archivos_que_no_se_pueden_convertir(tmp_path, libreoffice_simulado):
    texto = tmp_path / "notas.txt"
    texto.write_text("hola", encoding="utf-8")
    _, conversiones = a_pdf.convertir([tmp_path / "no_existe.docx", texto, original("01")], tmp_path,
                                      ("LibreOffice", Path("soffice")))
    assert "No se encuentra el archivo" in conversiones[0].error
    assert "no es un documento de Word" in conversiones[1].error
    assert conversiones[2].error is None
    assert len(libreoffice_simulado) == 1, "solo se convierte el documento válido"


# --- Word (simulado) ------------------------------------------------------------------------

def test_si_word_se_cuelga_solo_se_cierran_los_word_que_ha_abierto(tmp_path, monkeypatch):
    # El 111 es un Word que ya estaba abierto (puede ser el del estudiante): no se toca.
    automaticos = iter([{111}, {111, 222}])
    monkeypatch.setattr(a_pdf, "_word_automaticos", lambda: next(automaticos))

    def colgado(*_, **__):
        raise subprocess.TimeoutExpired("powershell", 1)

    cerrados = []
    monkeypatch.setattr(a_pdf, "_powershell", colgado)
    monkeypatch.setattr(a_pdf.subprocess, "run", lambda orden, **_: cerrados.append(orden))
    [conversion] = a_pdf.con_word([(original("01"), tmp_path / "a.pdf")])
    assert "Word ha tardado demasiado" in conversion.error
    assert cerrados == [["taskkill", "/PID", "222", "/F"]]


def test_si_word_no_arranca(tmp_path, monkeypatch):
    monkeypatch.setattr(a_pdf, "_word_automaticos", set)
    monkeypatch.setattr(a_pdf, "_powershell", lambda *_, **__: subprocess.CompletedProcess(
        [], 1, b"", "No se puede crear el objeto COM".encode("utf-8")))
    [conversion] = a_pdf.con_word([(original("01"), tmp_path / "a.pdf")])
    assert "No se ha podido usar Word (No se puede crear el objeto COM)" in conversion.error


# --- Word de verdad ---------------------------------------------------------------------------

@pytest.mark.skipif(not a_pdf.hay_word(), reason="Word no está instalado")
def test_conversion_real_con_word(tmp_path):
    copias = [pipeline.procesar(original(p), tmp_path)[0] for p in ("01", "05")]
    previos = a_pdf._word_automaticos()
    nombre, conversiones = a_pdf.convertir(copias + [tmp_path / "no_existe.docx"], tmp_path / "pdf")
    assert nombre == "Word"
    for copia, conversion in zip(copias, conversiones):
        assert conversion.error is None, conversion.error
        assert conversion.pdf == (tmp_path / "pdf" / f"{copia.stem}.pdf").resolve()
        assert conversion.pdf.read_bytes().startswith(b"%PDF")
    assert conversiones[0].paginas >= 1 and conversiones[1].paginas >= 5   # el TFG ocupa varias páginas
    assert "No se encuentra el archivo" in conversiones[2].error
    assert a_pdf._word_automaticos() - previos == set(), "no debe quedar ningún Word abierto"
