"""
Paso 5.2, prueba con Word: la portada APA ocupa sola la página 1, el título está
en la cuarta línea (tres líneas en blanco por encima, «three to four lines down
from the top» en la web oficial de APA) y el texto empieza en la página 2. Eso
solo lo puede decir Word, así que se le pregunta en qué página y a qué altura
empieza cada uno de los primeros párrafos. Si Word no está instalado, la prueba
se omite.

Se usa el mismo mecanismo que herramientas/a_pdf.py: PowerShell con llamadas
dinámicas, en solo lectura, sin avisos y cerrando siempre Word.
"""

import tempfile
from pathlib import Path

import docx
import pytest

from apa_facil import pipeline
from apa_facil.datos_portada import combinar
from herramientas import a_pdf

RAIZ = Path(__file__).resolve().parent.parent
DATOS_01 = RAIZ / "tests" / "datos" / "01_datos.json"

# Para cada uno de los primeros párrafos: "PARRAFO|número|página|altura en puntos desde el borde de la página".
SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$tipo = [System.__ComObject]
$word = New-Object -ComObject Word.Application
try {
    $tipo.InvokeMember('Visible', 'SetProperty', $null, $word, @($false))
    $tipo.InvokeMember('DisplayAlerts', 'SetProperty', $null, $word, @(0))
    $documentos = $tipo.InvokeMember('Documents', 'GetProperty', $null, $word, $null)
    $doc = $tipo.InvokeMember('Open', 'InvokeMethod', $null, $documentos, @($args[0], $false, $true, $false))
    try {
        $ventana = $tipo.InvokeMember('ActiveWindow', 'GetProperty', $null, $doc, $null)
        $vista = $tipo.InvokeMember('View', 'GetProperty', $null, $ventana, $null)
        $tipo.InvokeMember('Type', 'SetProperty', $null, $vista, @(3)) | Out-Null
        $pars = $tipo.InvokeMember('Paragraphs', 'GetProperty', $null, $doc, $null)
        $n = [Math]::Min([int]$args[1], [int]$tipo.InvokeMember('Count', 'GetProperty', $null, $pars, $null))
        for ($i = 1; $i -le $n; $i++) {
            $p = $tipo.InvokeMember('Item', 'InvokeMethod', $null, $pars, @($i))
            $r = $tipo.InvokeMember('Range', 'GetProperty', $null, $p, $null)
            $ini = $tipo.InvokeMember('Start', 'GetProperty', $null, $r, $null)
            $primera = $tipo.InvokeMember('Range', 'InvokeMethod', $null, $doc, @($ini, ($ini + 1)))
            $pagina = $tipo.InvokeMember('Information', 'GetProperty', $null, $primera, @(3))
            $altura = $tipo.InvokeMember('Information', 'GetProperty', $null, $primera, @(6))
            "PARRAFO|$i|$pagina|$altura"
        }
    } finally {
        $tipo.InvokeMember('Close', 'InvokeMethod', $null, $doc, @(0)) | Out-Null
    }
} finally {
    $tipo.InvokeMember('Quit', 'InvokeMethod', $null, $word, @(0)) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""


def posiciones_segun_word(ruta, cuantos):
    """[(página, altura)] de los primeros `cuantos` párrafos, según Word."""
    with tempfile.TemporaryDirectory() as carpeta:
        script = Path(carpeta) / "posiciones.ps1"
        script.write_text(SCRIPT, encoding="ascii")
        resultado = a_pdf._powershell("-File", str(script), str(ruta), str(cuantos), tiempo=120)
    salida = resultado.stdout.decode("utf-8", "replace").splitlines()
    filas = [(int(pagina), float(altura.replace(",", "."))) for _, _, pagina, altura in
             (linea.split("|") for linea in salida if linea.startswith("PARRAFO|"))]
    assert filas, resultado.stderr.decode("utf-8", "replace")
    return filas


@pytest.mark.skipif(not a_pdf.hay_word(), reason="Word no está instalado")
def test_word_ve_la_portada_sola_en_la_primera_pagina(tmp_path):
    sin_portada = tmp_path / "01_sin_portada.docx"
    doc = docx.Document(str(RAIZ / "pruebas" / "01_ensayo_basico.docx"))
    for p in doc.paragraphs[:5]:   # su portada
        p._element.getparent().remove(p._element)
    doc.save(str(sin_portada))
    copia, _ = pipeline.procesar(sin_portada, tmp_path / "salida", {"portada": combinar(DATOS_01)})

    # 3 en blanco, título, en blanco, 5 datos; después, el título del texto y su primer párrafo.
    filas = posiciones_segun_word(copia, 12)
    portada, texto = filas[:10], filas[10:]
    assert all(pagina == 1 for pagina, _ in portada), "la portada cabe en la página 1"
    assert all(pagina == 2 for pagina, _ in texto), "el texto empieza en la página 2"
    alturas = [altura for _, altura in portada]
    linea = alturas[1] - alturas[0]                       # una línea a doble espacio
    assert linea > 0
    assert alturas[0] == pytest.approx(72, abs=1)         # la primera línea, en el margen de 2,54 cm
    assert alturas[3] == pytest.approx(alturas[0] + 3 * linea, abs=1), "el título, en la cuarta línea"
    for anterior, siguiente in zip(alturas, alturas[1:]):
        assert siguiente - anterior == pytest.approx(linea, abs=1), "todo a doble espacio, sin espacio extra"
    assert texto[0][1] == pytest.approx(72, abs=1), "el título del texto, arriba de la página 2"


# --- Paso 5.3: la portada de la universidad, sustituida ------------------------------------------

@pytest.mark.skipif(not a_pdf.hay_word(), reason="Word no está instalado")
@pytest.mark.parametrize("nombre", ["01_ensayo_basico", "05_tfg_completo"])
def test_word_ve_la_portada_sustituida_sola_en_la_primera_pagina(nombre, tmp_path):
    datos = RAIZ / "tests" / "datos" / f"{nombre[:2]}_datos.json"
    copia, _ = pipeline.procesar(RAIZ / "pruebas" / f"{nombre}.docx", tmp_path, {"portada": combinar(datos)})
    filas = posiciones_segun_word(copia, 11)
    portada, siguiente = filas[:10], filas[10]
    assert all(pagina == 1 for pagina, _ in portada), "la portada APA, sola en la página 1"
    assert siguiente[0] == 2 and siguiente[1] == pytest.approx(72, abs=1), "lo siguiente, arriba de la página 2"
    alturas = [altura for _, altura in portada]
    assert alturas[3] == pytest.approx(alturas[0] + 3 * (alturas[1] - alturas[0]), abs=1), "el título, en la 4.ª línea"
