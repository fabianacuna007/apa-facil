"""
Línea de comandos:

    python -m apa_facil mi_trabajo.docx [otro.docx ...] [--fuente "Arial"]

Guarda cada copia en salida/ como nombre_APA.docx y muestra un resumen de lo
que se ha cambiado y de lo que el estudiante debe revisar.
"""

import argparse
import glob
import sys
from pathlib import Path

from apa_facil import pipeline
from apa_facil.documento import ErrorDocumento
from apa_facil.reglas import fuente

EJEMPLO = "python -m apa_facil mi_trabajo.docx"
FUENTES = ", ".join(list(fuente.RECOMENDADAS)[:-1]) + " o " + list(fuente.RECOMENDADAS)[-1]


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
        elif "--fuente" in message and "expected one argument" in message:
            texto = f'Falta indicar la fuente después de --fuente. Ejemplo: {EJEMPLO} --fuente "Arial"'
        elif message.startswith("argument "):   # mensaje propio de una opción (p. ej. --fuente)
            texto = message.split(": ", 1)[1]
        else:
            texto = message
        self.exit(2, f"{texto}\n")


def _fuente(valor):
    """Comprueba la opción --fuente: una de las recomendadas por APA (o una variante de su familia)."""
    elegida = fuente.familia(valor.strip())
    if elegida is None:
        raise argparse.ArgumentTypeError(f"«{valor}» no es una de las fuentes recomendadas por APA. "
                                         f"Elige una de estas: {FUENTES}. No se ha procesado ningún documento.")
    return elegida


def crear_ordenes():
    ordenes = _Ordenes(prog="python -m apa_facil", formatter_class=_Formato, add_help=False,
                       description="Aplica el formato APA 7 de estudiante a una copia de un documento Word.")
    ordenes._positionals.title = "argumentos"
    ordenes._optionals.title = "opciones"
    ordenes.add_argument("archivos", nargs="+", metavar="archivo.docx",
                         help="uno o varios documentos de Word")
    tamanos = ", ".join(f"{nombre} {puntos}" for nombre, puntos in fuente.RECOMENDADAS.items())
    ordenes.add_argument("--fuente", type=_fuente, metavar='"NOMBRE"',
                         help=f"fuente para todo el trabajo, con su tamaño: {tamanos}. Si no se indica, se "
                              "mantiene la del documento si es una de estas, o se usa la más parecida")
    ordenes.add_argument("-h", "--help", action="help", help="muestra esta ayuda")
    return ordenes


def configurar_terminal():
    """
    En una terminal de Windows, Python ya escribe en UTF-8, y los acentos y
    símbolos (✔, ✘) salen bien. Si la salida va a un archivo, a otro programa o
    a NUL, Python usa la codificación antigua de Windows, que no tiene "✔", y
    fallaría al escribirlo: en esos casos se cambia a UTF-8. (No basta con
    mirar si es una terminal: Windows hace creer a Python que NUL lo es.)
    """
    for flujo in (sys.stdout, sys.stderr):
        codificacion = (getattr(flujo, "encoding", None) or "").lower().replace("-", "")
        if hasattr(flujo, "reconfigure") and codificacion != "utf8":
            flujo.reconfigure(encoding="utf-8", errors="replace")


def mostrar_informe(informe):
    """Resumen del informe en la terminal: cambios, lo que hay que revisar y lo que no se pudo hacer."""
    print(informe.como_texto())


def expandir_comodines(archivos):
    """
    En la terminal de Windows (cmd o PowerShell), "pruebas\\*.docx" llega tal
    cual: Windows no expande los comodines. Se expanden aquí. Si un comodín no
    encuentra nada, se deja como está para que el mensaje de error lo nombre.
    """
    resultado = []
    for archivo in archivos:
        coincidencias = sorted(glob.glob(archivo)) if any(c in archivo for c in "*?") else []
        resultado += coincidencias or [archivo]
    return resultado


def main(argv=None, carpeta_salida=None):
    configurar_terminal()
    args = crear_ordenes().parse_args(argv)
    opciones = {"fuente": args.fuente} if args.fuente else {}

    archivos = expandir_comodines(args.archivos)
    bien = 0
    for archivo in archivos:
        nombre = Path(archivo).name
        try:
            salida, informe = pipeline.procesar(archivo, carpeta_salida, opciones)
        except ErrorDocumento as e:
            print(f"✘ {nombre}: {e}")
            continue
        except Exception as e:  # un fallo inesperado no debe parar los demás archivos
            print(f"✘ {nombre}: error inesperado al procesarlo ({type(e).__name__}: {e}).")
            continue
        print(f"✔ {nombre} → {salida.parent.name}/{salida.name}")
        mostrar_informe(informe)
        json = salida.with_suffix(".json")
        print(f"   Informe completo: {json.parent.name}/{json.name}")
        bien += 1

    total = len(archivos)
    print(f"\n{bien} de {total} documento(s) procesado(s).")
    return 0 if bien == total else 1


if __name__ == "__main__":
    sys.exit(main())
