"""
Paso 1.7: la línea de comandos completa.

    python -m apa_facil entrada.docx --fuente "Arial"

genera salida/entrada_APA.docx y muestra un resumen del informe. Las copias se
guardan en carpetas temporales para no tocar salida/; pruebas/ no se modifica.
"""

import json
import os
import subprocess
import sys
from collections import Counter
from pathlib import Path

import docx
import pytest

from apa_facil.__main__ import expandir_comodines, main, mostrar_informe
from apa_facil.documento import Estilos, w
from apa_facil.informe import Informe

RAIZ = Path(__file__).resolve().parent.parent
PRUEBAS = sorted((RAIZ / "pruebas").glob("*.docx"))


def fuentes(ruta):
    doc = docx.Document(str(ruta))
    estilos, cuenta = Estilos(doc), Counter()
    for p in doc.element.body.iter(w("p")):
        for r in p.iter(w("r")):
            if "".join(t.text or "" for t in r.iter(w("t"))).strip():
                cuenta[(estilos.fuente(r, p), estilos.tamano(r, p))] += 1
    return set(cuenta)


# --- Opción --fuente --------------------------------------------------------------------

@pytest.mark.parametrize("pedida, esperada", [
    ("Arial", ("Arial", 11)), ("arial", ("Arial", 11)), ("Calibri Light", ("Calibri", 11)),
    ("Times New Roman", ("Times New Roman", 12)), ("Aptos", ("Aptos", 12)),
])
def test_fuente_elegida(pedida, esperada, tmp_path, capsys):
    original = PRUEBAS[2]   # el 03, en Georgia
    assert main([str(original), "--fuente", pedida], carpeta_salida=tmp_path) == 0
    assert fuentes(tmp_path / f"{original.stem}_APA.docx") == {esperada}
    assert "como has pedido" in capsys.readouterr().out


def test_fuente_no_recomendada_no_procesa_nada(tmp_path, capsys):
    with pytest.raises(SystemExit) as fin:
        main([str(PRUEBAS[0]), "--fuente", "Papyrus"], carpeta_salida=tmp_path)
    assert fin.value.code == 2
    error = capsys.readouterr().err
    assert error.startswith("«Papyrus» no es una de las fuentes recomendadas por APA. Elige una de estas: "
                            "Aptos, Calibri, Arial, Lucida Sans Unicode, Times New Roman, Georgia o Computer Modern.")
    assert not list(tmp_path.iterdir()), "no se ha guardado ninguna copia"


def test_fuente_sin_valor(capsys):
    with pytest.raises(SystemExit):
        main([str(PRUEBAS[0]), "--fuente"])
    assert "Falta indicar la fuente después de --fuente" in capsys.readouterr().err


def test_la_ayuda_explica_la_opcion(capsys):
    with pytest.raises(SystemExit):
        main(["-h"])
    ayuda = capsys.readouterr().out
    assert "--fuente" in ayuda and "Georgia 11" in ayuda and "Uso:" in ayuda


# --- Resumen del informe ----------------------------------------------------------------------

def test_resumen_con_grupos_y_cantidades(capsys):
    informe = Informe()
    informe.cambio("Uno.")
    informe.cambio("Dos.")
    informe.aviso("Revisa esto.")
    informe.error("Esto no se pudo.")
    mostrar_informe(informe)
    assert capsys.readouterr().out.splitlines() == [
        "   Lo que hemos cambiado (2):", "     · Uno.", "     · Dos.",
        "   Lo que debes revisar tú (1):", "     · Revisa esto.",
        "   Lo que no se ha podido hacer (1):", "     · Esto no se pudo.",
    ]


def test_resumen_sin_cambios(capsys):
    mostrar_informe(Informe())
    assert capsys.readouterr().out.strip() == "No ha hecho falta cambiar nada: el formato ya cumplía APA."


def test_salida_completa_de_un_documento(tmp_path, capsys):
    original = next(p for p in PRUEBAS if p.name.startswith("03"))
    assert main([str(original)], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert salida.startswith(f"✔ {original.name} → {tmp_path.name}/{original.stem}_APA.docx\n   Lo que hemos cambiado (")
    # Encabezado, un solo subapartado, numeración, «Introducción» y la portada que se mantiene (sin datos, paso 5.3);
    # y de las figuras (6.2), «(elaboración propia)» en un título, la línea en blanco y dónde colocarlas.
    assert "   Lo que debes revisar tú (8):" in salida
    assert f"   Informe completo: {tmp_path.name}/{original.stem}_APA.json" in salida
    assert "todavía incompleto" not in salida
    assert salida.rstrip().endswith("1 de 1 documento(s) procesado(s).")


# --- Comodines y salida a NUL ------------------------------------------------------------------

def test_comodines(tmp_path):
    assert expandir_comodines([str(RAIZ / "pruebas" / "*.docx")]) == [str(p) for p in PRUEBAS]
    assert expandir_comodines([str(tmp_path / "*.docx")]) == [str(tmp_path / "*.docx")], "sin coincidencias"
    assert expandir_comodines(["trabajo.docx"]) == ["trabajo.docx"]


def test_salida_a_nul(tmp_path):
    """Con la salida a NUL, Python cree que es una terminal y usaba la codificación antigua (fallaba con ✔)."""
    entorno = {k: v for k, v in os.environ.items() if k not in ("PYTHONIOENCODING", "PYTHONUTF8")}
    copia = tmp_path / "prueba_salida_nul.docx"   # nombre propio: su copia en salida/ se borra al final
    copia.write_bytes(PRUEBAS[0].read_bytes())
    generada = RAIZ / "salida" / "prueba_salida_nul_APA.docx"
    try:
        resultado = subprocess.run([sys.executable, "-m", "apa_facil", str(copia)], cwd=RAIZ, env=entorno,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        assert resultado.returncode == 0, resultado.stderr.decode("utf-8", "replace")
        assert generada.exists()
    finally:
        generada.unlink(missing_ok=True)
        generada.with_suffix(".json").unlink(missing_ok=True)


# --- Los cinco documentos -----------------------------------------------------------------------

def test_los_cinco_documentos(tmp_path, capsys):
    assert main([str(RAIZ / "pruebas" / "*.docx")], carpeta_salida=tmp_path) == 0
    assert "5 de 5 documento(s) procesado(s)." in capsys.readouterr().out
    assert len(list(tmp_path.glob("*_APA.docx"))) == 5


def test_los_cinco_en_georgia(tmp_path):
    assert main([str(p) for p in PRUEBAS] + ["--fuente", "Georgia"], carpeta_salida=tmp_path) == 0
    for original in PRUEBAS:
        assert fuentes(tmp_path / f"{original.stem}_APA.docx") == {("Georgia", 11)}, original.name


# --- Modo diagnóstico (paso 3.3) ------------------------------------------------------------------

def test_diagnostico_no_genera_copias(tmp_path, capsys):
    antes = sorted((RAIZ / "salida").iterdir()) if (RAIZ / "salida").exists() else []
    assert main([str(PRUEBAS[3]), "--diagnostico"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert f"Diagnóstico de {PRUEBAS[3].name} (33 párrafos)" in salida
    assert "── Portada " in salida and "título (nivel 2)" in salida
    assert salida.rstrip().endswith("1 de 1 documento(s) diagnosticado(s). No se ha generado ninguna copia.")
    assert list(tmp_path.iterdir()) == []
    assert (sorted((RAIZ / "salida").iterdir()) if (RAIZ / "salida").exists() else []) == antes


def test_diagnostico_de_los_cinco_con_comodin(capsys):
    assert main([str(RAIZ / "pruebas" / "*.docx"), "--diagnostico"]) == 0
    salida = capsys.readouterr().out
    assert salida.count("Diagnóstico de ") == 5
    assert "5 de 5 documento(s) diagnosticado(s)." in salida


def test_diagnostico_de_un_archivo_que_no_existe(capsys):
    assert main(["no_existe.docx", "--diagnostico"]) == 1
    assert "✘ no_existe.docx: No se encuentra el archivo" in capsys.readouterr().out


def test_diagnostico_avisa_de_que_no_usa_la_fuente(capsys):
    assert main([str(PRUEBAS[0]), "--diagnostico", "--fuente", "Arial"]) == 0
    assert capsys.readouterr().out.startswith(
        "Aviso: la opción --fuente no se usa en el modo diagnóstico, que no cambia nada.")


def test_la_ayuda_explica_el_diagnostico(capsys):
    with pytest.raises(SystemExit):
        main(["-h"])
    ayuda = capsys.readouterr().out
    assert "--diagnostico" in ayuda and "no genera ninguna copia" in ayuda


def test_opciones_de_los_titulos(tmp_path, capsys):
    original = next(p for p in PRUEBAS if p.name.startswith("05"))
    assert main([str(original), "--quitar-numeracion"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert "Títulos: se ha quitado la numeración de 12 títulos" in salida
    # Las mayúsculas de APA, por defecto (paso 4.6).
    assert "Títulos: 8 títulos llevan ahora mayúscula inicial en las palabras principales, como pide APA " \
           "(«RESUMEN» → «Resumen»" in salida
    assert "--quitar-numeracion" not in salida, "sin el aviso"
    datos = json.loads((tmp_path / f"{original.stem}_APA.json").read_text(encoding="utf-8"))
    assert datos["opciones"] == {"fuente": None, "quitar_numeracion": True, "titulos_tal_cual": False,
                                 "quitar_introduccion": False,
                                 "titulo_tal_cual": False, "mantener_portada": False, "quitar_indice": False,
                                 "tablas_tal_cual": False, "rotulo_tal_cual": False, "orden_tal_cual": False,
                                 "rotulo_apendice": False,
                                 "portada": None}


def test_titulos_tal_cual(tmp_path, capsys):
    original = next(p for p in PRUEBAS if p.name.startswith("05"))
    assert main([str(original), "--titulos-tal-cual"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert "títulos llevan ahora mayúscula inicial" not in salida
    assert "no llevan las mayúsculas de APA («RESUMEN»" in salida
    datos = json.loads((tmp_path / f"{original.stem}_APA.json").read_text(encoding="utf-8"))
    assert datos["opciones"]["titulos_tal_cual"] is True


def test_quitar_mayusculas_ya_no_existe(capsys):
    # Desde el paso 4.6 las mayúsculas de APA se ponen por defecto: la opción antigua no se reconoce.
    with pytest.raises(SystemExit):
        main([str(PRUEBAS[0]), "--quitar-mayusculas"])
    assert "Opción no reconocida: --quitar-mayusculas" in capsys.readouterr().err


def test_opciones_de_los_titulos_en_el_diagnostico(capsys):
    assert main([str(PRUEBAS[0]), "--diagnostico", "--titulos-tal-cual"]) == 0
    assert "Aviso: la opción --titulos-tal-cual no se usa en el modo diagnóstico" in capsys.readouterr().out
