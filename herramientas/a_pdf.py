"""
Revisión visual: convierte documentos Word a PDF para revisarlos sin abrir Word.

Usa el primer conversor que encuentra:
1. Word, si está instalado (Windows). Es lo que verá el estudiante, así que el
   PDF es fiel. Se maneja desde PowerShell con llamadas dinámicas (InvokeMember):
   el acceso normal a Word por COM falla con E_NOINTERFACE en este equipo
   (ver "Problemas abiertos" en ROADMAP.md).
2. LibreOffice, si está instalado. Su resultado se parece al de Word, pero no
   es idéntico (puede cambiar dónde cae cada salto de página).

Los documentos se abren en solo lectura: nunca se modifican. Los PDF se guardan
siempre en salida/ (aunque el documento venga de pruebas/, que no se toca).

Uso:
    python herramientas/a_pdf.py salida/*.docx
"""

import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

# Para encontrar el paquete apa_facil al ejecutar este archivo directamente.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apa_facil import documento  # noqa: E402
from apa_facil.__main__ import configurar_terminal, expandir_comodines  # noqa: E402
from apa_facil.documento import ErrorDocumento  # noqa: E402

SEGUNDOS_BASE = 60          # arrancar Word puede tardar
SEGUNDOS_POR_DOCUMENTO = 60

RUTAS_LIBREOFFICE = [Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
                     Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
                     Path("/usr/bin/soffice"), Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]

# Script de PowerShell que maneja Word. Recibe pares (documento, pdf) y escribe
# una línea por documento: "OK|páginas" o "ERROR|mensaje". Solo usa caracteres
# ASCII: PowerShell 5.1 lee mal los .ps1 con tildes si no llevan BOM.
# - DisplayAlerts = 0: sin cuadros de diálogo, que dejarían a Word (invisible) esperando.
# - Open(archivo, ConfirmConversions, ReadOnly, AddToRecentFiles): en solo lectura.
# - ExportAsFixedFormat(pdf, 17): 17 = PDF. Las páginas se cuentan después en el
#   PDF: el recuento de Word (ComputeStatistics) daba una página de más en 4 de
#   los 5 documentos de prueba.
# - Close(0) y Quit(0): sin guardar. Quit va en un finally para que no quede
#   un Word abierto en segundo plano si algo falla.
SCRIPT_WORD = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$tipo = [System.__ComObject]
$word = New-Object -ComObject Word.Application
try {
    $tipo.InvokeMember('Visible', 'SetProperty', $null, $word, @($false))
    $tipo.InvokeMember('DisplayAlerts', 'SetProperty', $null, $word, @(0))
    $documentos = $tipo.InvokeMember('Documents', 'GetProperty', $null, $word, $null)
    for ($i = 0; $i -lt $args.Count; $i += 2) {
        $doc = $null
        try {
            $doc = $tipo.InvokeMember('Open', 'InvokeMethod', $null, $documentos, @($args[$i], $false, $true, $false))
            $tipo.InvokeMember('ExportAsFixedFormat', 'InvokeMethod', $null, $doc, @($args[$i + 1], 17))
            "OK|"
        } catch {
            $fallo = $_.Exception
            while ($fallo.InnerException) { $fallo = $fallo.InnerException }
            "ERROR|" + ($fallo.Message -replace "[\r\n]+", " ")
        } finally {
            if ($doc) { $tipo.InvokeMember('Close', 'InvokeMethod', $null, $doc, @(0)) | Out-Null }
        }
    }
} finally {
    $tipo.InvokeMember('Quit', 'InvokeMethod', $null, $word, @(0)) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""

# Los Word abiertos por automatización (los que abre este script), para cerrar
# solo esos si la conversión se queda colgada. Nunca el Word del usuario.
SCRIPT_WORD_AUTOMATICOS = ("Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE'\" | "
                           "Where-Object { $_.CommandLine -match '/Automation' } | "
                           "ForEach-Object { $_.ProcessId }")


@dataclass
class Conversion:
    archivo: str               # nombre del documento
    pdf: Path = None           # el PDF creado, si se ha podido
    paginas: int = None
    error: str = None


# --- Conversores ---------------------------------------------------------------------

def _powershell(*argumentos, tiempo=None):
    return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                           *argumentos], capture_output=True, timeout=tiempo)


def hay_word():
    if sys.platform != "win32":
        return False
    import winreg
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID").Close()
        return True
    except OSError:
        return False


def buscar_libreoffice():
    for nombre in ("soffice", "libreoffice"):
        ruta = shutil.which(nombre)
        if ruta:
            return Path(ruta)
    return next((r for r in RUTAS_LIBREOFFICE if r.exists()), None)


def buscar_conversor():
    """("Word", None), ("LibreOffice", ruta de soffice) o None si no hay ninguno."""
    if hay_word():
        return "Word", None
    soffice = buscar_libreoffice()
    return ("LibreOffice", soffice) if soffice else None


def _word_automaticos():
    resultado = _powershell("-Command", SCRIPT_WORD_AUTOMATICOS, tiempo=30)
    return {int(linea) for linea in resultado.stdout.decode("utf-8", "replace").split() if linea.isdigit()}


def con_word(pares):
    """Convierte cada (docx, pdf) con Word, en una sola sesión de Word."""
    previos = _word_automaticos()
    with tempfile.TemporaryDirectory() as carpeta:
        script = Path(carpeta) / "a_pdf.ps1"
        script.write_text(SCRIPT_WORD, encoding="ascii")
        argumentos = [str(ruta) for par in pares for ruta in par]
        try:
            resultado = _powershell("-File", str(script), *argumentos,
                                    tiempo=SEGUNDOS_BASE + SEGUNDOS_POR_DOCUMENTO * len(pares))
        except subprocess.TimeoutExpired:
            # Word se ha quedado colgado (un cuadro de diálogo, un documento enorme…):
            # se cierran los Word que ha abierto esta conversión.
            for pid in _word_automaticos() - previos:
                subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
            return [Conversion(docx.name, error="Word ha tardado demasiado y se ha cerrado sin terminar.")
                    for docx, _ in pares]

    lineas = [l for l in resultado.stdout.decode("utf-8", "replace").splitlines() if l.startswith(("OK|", "ERROR|"))]
    if len(lineas) != len(pares):   # Word no ha llegado a arrancar
        detalle = " ".join(resultado.stderr.decode("utf-8", "replace").split())[:300]
        return [Conversion(docx.name, error=f"No se ha podido usar Word ({detalle or 'sin más detalles'}).")
                for docx, _ in pares]
    conversiones = []
    for (docx, pdf), linea in zip(pares, lineas):
        estado, dato = linea.split("|", 1)
        if estado == "OK":
            conversiones.append(Conversion(docx.name, pdf, paginas_del_pdf(pdf)))
        else:
            conversiones.append(Conversion(docx.name, error=f"Word no ha podido convertirlo: {dato.strip()}"))
    return conversiones


def paginas_del_pdf(pdf):
    """
    Las páginas de un PDF, o None si no se sabe. Se lee el total que declara el
    árbol de páginas (/Type /Pages … /Count N); si no está a la vista (puede ir
    comprimido), se cuentan las páginas una a una (/Type /Page).
    """
    datos = Path(pdf).read_bytes()
    totales = [int(n) for n in re.findall(rb"/Type\s*/Pages\b[^>]*?/Count\s+(\d+)", datos)]
    totales += [int(n) for n in re.findall(rb"/Count\s+(\d+)[^>]*?/Type\s*/Pages\b", datos)]
    if totales:
        return max(totales)
    return len(re.findall(rb"/Type\s*/Page(?![a-zA-Z])", datos)) or None


def con_libreoffice(soffice, pares):
    """Convierte cada (docx, pdf) con LibreOffice, que guarda el PDF con el nombre del documento."""
    conversiones = []
    for docx, pdf in pares:
        try:
            resultado = subprocess.run([str(soffice), "--headless", "--norestore", "--convert-to", "pdf",
                                        "--outdir", str(pdf.parent), str(docx)],
                                       capture_output=True, timeout=SEGUNDOS_BASE + SEGUNDOS_POR_DOCUMENTO)
        except subprocess.TimeoutExpired:
            conversiones.append(Conversion(docx.name, error="LibreOffice ha tardado demasiado."))
            continue
        creado = pdf.parent / f"{docx.stem}.pdf"
        if resultado.returncode != 0 or not creado.exists():
            detalle = " ".join(resultado.stderr.decode("utf-8", "replace").split())[:300]
            conversiones.append(Conversion(docx.name, error=f"LibreOffice no ha podido convertirlo ({detalle})."))
            continue
        if creado != pdf:
            creado.replace(pdf)
        conversiones.append(Conversion(docx.name, pdf, paginas_del_pdf(pdf)))
    return conversiones


# --- Conversión ------------------------------------------------------------------------

def convertir(archivos, carpeta=None, conversor=None):
    """
    Convierte los documentos a PDF en `carpeta` (salida/ por defecto).
    Devuelve (nombre del conversor, lista de Conversion).
    Lanza ErrorDocumento si no hay ningún conversor.
    """
    conversor = conversor or buscar_conversor()
    if conversor is None:
        raise ErrorDocumento("Para crear los PDF hace falta Word o LibreOffice, y no se ha encontrado "
                             "ninguno. LibreOffice es gratuito: https://es.libreoffice.org/descarga/")
    nombre, soffice = conversor
    carpeta = Path(carpeta) if carpeta is not None else documento.CARPETA_SALIDA
    carpeta.mkdir(parents=True, exist_ok=True)

    conversiones = [None] * len(archivos)   # en el mismo orden que `archivos`
    pendientes, pares = [], []              # posición y (docx, pdf) de los que se pueden convertir
    for i, archivo in enumerate(archivos):
        ruta = Path(archivo)
        try:
            documento.abrir(ruta)   # mismas comprobaciones y mensajes que al procesar
        except ErrorDocumento as e:
            conversiones[i] = Conversion(ruta.name, error=str(e))
            continue
        pendientes.append(i)
        pares.append((ruta.resolve(), (carpeta / f"{ruta.stem}.pdf").resolve()))

    if pares:
        hechas = con_word(pares) if nombre == "Word" else con_libreoffice(soffice, pares)
        for i, conversion in zip(pendientes, hechas):
            conversiones[i] = conversion
    return nombre, conversiones


def main(argv=None, carpeta=None):
    configurar_terminal()
    archivos = expandir_comodines(sys.argv[1:] if argv is None else argv)
    if not archivos:
        print("Indica qué documentos convertir a PDF. Ejemplo: python herramientas/a_pdf.py salida/*.docx")
        return 2
    try:
        nombre, conversiones = convertir(archivos, carpeta)
    except ErrorDocumento as e:
        print(f"✘ {e}")
        return 1

    print(f"Conversor: {nombre}")
    for c in conversiones:
        if c.error:
            print(f"✘ {c.archivo}: {c.error}")
        else:
            paginas = f" ({c.paginas} página{'s' if c.paginas != 1 else ''})" if c.paginas else ""
            print(f"✔ {c.archivo} → {c.pdf.parent.name}/{c.pdf.name}{paginas}")
    bien = sum(1 for c in conversiones if not c.error)
    print(f"\n{bien} de {len(conversiones)} PDF creado(s).")
    return 0 if bien == len(conversiones) else 1


if __name__ == "__main__":
    sys.exit(main())
