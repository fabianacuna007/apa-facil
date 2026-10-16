"""
Línea de comandos:

    python -m apa_facil mi_trabajo.docx [otro.docx ...]

Guarda cada copia en salida/ como nombre_APA.docx. Se completa en el paso 1.7.
"""

import argparse
import sys
from pathlib import Path

from apa_facil import pipeline
from apa_facil.documento import ErrorDocumento

EJEMPLO = "python -m apa_facil mi_trabajo.docx"


class _Formato(argparse.HelpFormatter):
    def add_usage(self, usage, actions, groups, prefix=None):
        return super().add_usage(usage, actions, groups, prefix="Uso: ")


class _Ordenes(argparse.ArgumentParser):
    """argparse escribe sus mensajes en inglés; aquí se traducen."""

    def error(self, message):
        if "required" in message:
            texto = f"Falta indicar qué documento procesar. Ejemplo: {EJEMPLO}"
        elif "unrecognized arguments" in message:
            texto = "Opción no reconocida: " + message.split(":", 1)[1].strip()
        else:
            texto = message
        self.exit(2, f"{texto}\n")


def crear_ordenes():
    ordenes = _Ordenes(prog="python -m apa_facil", formatter_class=_Formato, add_help=False,
                       description="Aplica el formato APA 7 de estudiante a una copia de un documento Word.")
    ordenes._positionals.title = "argumentos"
    ordenes._optionals.title = "opciones"
    ordenes.add_argument("archivos", nargs="+", metavar="archivo.docx",
                         help="uno o varios documentos de Word")
    ordenes.add_argument("-h", "--help", action="help", help="muestra esta ayuda")
    return ordenes


def configurar_terminal():
    """
    En una terminal de Windows, Python ya escribe bien acentos y símbolos (✔, ✘).
    Si la salida va a otro programa o a un archivo, se fuerza UTF-8 para que
    no lleguen desfigurados.
    """
    for flujo in (sys.stdout, sys.stderr):
        if hasattr(flujo, "reconfigure") and not flujo.isatty():
            flujo.reconfigure(encoding="utf-8", errors="replace")


def mostrar_informe(informe):
    """Resumen provisional del informe. El definitivo llega en el paso 1.7."""
    print("   Formato APA todavía incompleto: de momento solo se aplican los márgenes y la fuente.")
    for titulo, mensajes in (("Cambios", informe.cambios), ("Revisa tú", informe.avisos),
                             ("No se ha podido hacer", informe.errores)):
        if mensajes:
            print(f"   {titulo}:")
            for mensaje in mensajes:
                print(f"     · {mensaje}")
    if not informe.cambios:
        print("   No ha hecho falta cambiar nada.")


def main(argv=None, carpeta_salida=None):
    configurar_terminal()
    args = crear_ordenes().parse_args(argv)

    bien = 0
    for archivo in args.archivos:
        nombre = Path(archivo).name
        try:
            salida, informe = pipeline.procesar(archivo, carpeta_salida)
        except ErrorDocumento as e:
            print(f"✘ {nombre}: {e}")
            continue
        except Exception as e:  # un fallo inesperado no debe parar los demás archivos
            print(f"✘ {nombre}: error inesperado al procesarlo ({type(e).__name__}: {e}).")
            continue
        print(f"✔ {nombre} → {salida.parent.name}/{salida.name}")
        mostrar_informe(informe)
        bien += 1

    total = len(args.archivos)
    print(f"\n{bien} de {total} documento(s) procesado(s).")
    return 0 if bien == total else 1


if __name__ == "__main__":
    sys.exit(main())
