"""
Paso 4.5, prueba específica: los títulos de nivel 4 y 5 se ven en la misma línea
que su texto. Eso solo lo puede decir Word (el XML sigue teniendo dos párrafos),
así que la copia se abre con Word y se le pregunta en qué línea empieza el título
y en cuál su texto. Si Word no está instalado, la prueba se omite.

Se usa el mismo mecanismo que herramientas/a_pdf.py: PowerShell con llamadas
dinámicas, en solo lectura, sin avisos y cerrando siempre Word.
"""

import tempfile
from pathlib import Path

import pytest

from apa_facil import pipeline
from herramientas import a_pdf

PRUEBAS = Path(__file__).resolve().parent.parent / "pruebas"
TFG = PRUEBAS / "05_tfg_completo.docx"
# El principio del párrafo que sigue a cada título de nivel 4 del 05.
TEXTOS = ["Describir la evolución", "Conocer la valoración"]

# Para cada párrafo de nivel 4 o 5 que ve Word: "LINEA|dónde está el título|dónde está el texto|posición
# del texto dentro del párrafo de Word" (-1 si Word no los ve como un solo párrafo). "Dónde" es la página y
# la altura en la página: el número de línea que da Word (Information 10) no es fiable al principio de un
# párrafo (en la prueba daba la línea 29 para un título y la 16 para su texto, a la misma altura). Solo ASCII.
SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
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
        $n = $tipo.InvokeMember('Count', 'GetProperty', $null, $pars, $null)
        for ($i = 1; $i -le $n; $i++) {
            $p = $tipo.InvokeMember('Item', 'InvokeMethod', $null, $pars, @($i))
            $nivel = $tipo.InvokeMember('OutlineLevel', 'GetProperty', $null, $p, $null)
            if ($nivel -ne 4 -and $nivel -ne 5) { continue }
            $r = $tipo.InvokeMember('Range', 'GetProperty', $null, $p, $null)
            $texto = $tipo.InvokeMember('Text', 'GetProperty', $null, $r, $null)
            $ini = $tipo.InvokeMember('Start', 'GetProperty', $null, $r, $null)
            $k = -1
            for ($j = 1; $j -lt $args.Count; $j++) { if ($k -lt 0) { $k = $texto.IndexOf($args[$j]) } }
            $primera = $tipo.InvokeMember('Range', 'InvokeMethod', $null, $doc, @($ini, ($ini + 1)))
            $lineaTitulo = [string]$tipo.InvokeMember('Information', 'GetProperty', $null, $primera, @(3)) + "/" +
                           $tipo.InvokeMember('Information', 'GetProperty', $null, $primera, @(6))
            $lineaTexto = "-"
            if ($k -ge 0) {
                $letra = $tipo.InvokeMember('Range', 'InvokeMethod', $null, $doc, @(($ini + $k), ($ini + $k + 1)))
                $lineaTexto = [string]$tipo.InvokeMember('Information', 'GetProperty', $null, $letra, @(3)) + "/" +
                              $tipo.InvokeMember('Information', 'GetProperty', $null, $letra, @(6))
            }
            "LINEA|$lineaTitulo|$lineaTexto|$k"
        }
    } finally {
        $tipo.InvokeMember('Close', 'InvokeMethod', $null, $doc, @(0)) | Out-Null
    }
} finally {
    $tipo.InvokeMember('Quit', 'InvokeMethod', $null, $word, @(0)) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""


def lineas_segun_word(docx):
    """[(página/altura del título, página/altura del texto, posición del texto en el párrafo de Word)] de cada nivel 4 y 5."""
    with tempfile.TemporaryDirectory() as carpeta:
        script = Path(carpeta) / "en_linea.ps1"
        script.write_text(SCRIPT, encoding="ascii")
        resultado = a_pdf._powershell("-File", str(script), str(docx), *TEXTOS, tiempo=120)
    salida = resultado.stdout.decode("utf-8", "replace").splitlines()
    filas = [(titulo, texto, int(posicion)) for titulo, texto, posicion in
             (linea.split("|")[1:] for linea in salida if linea.startswith("LINEA|"))]
    assert filas, resultado.stderr.decode("utf-8", "replace")
    return filas


@pytest.mark.skipif(not a_pdf.hay_word(), reason="Word no está instalado")
def test_word_ve_los_titulos_de_nivel_4_en_la_misma_linea_que_su_texto(tmp_path):
    copia, _ = pipeline.procesar(TFG, tmp_path)
    filas = lineas_segun_word(copia)
    assert len(filas) == 2
    for linea_titulo, linea_texto, posicion in filas:
        assert posicion > 0, "Word ve el título y su texto como un solo párrafo"
        assert linea_titulo == linea_texto, "y en la misma línea"

    # A la contra: en el original, Word los ve separados (el texto en otro párrafo y otra línea).
    assert all(posicion == -1 for _, _, posicion in lineas_segun_word(TFG))
