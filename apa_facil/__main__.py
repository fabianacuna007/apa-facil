"""
Línea de comandos:

    python -m apa_facil mi_trabajo.docx [otro.docx ...] [--fuente "Arial"]
                        [--quitar-numeracion] [--titulos-tal-cual] [--quitar-introduccion] [--quitar-indice]
                        [--tablas-tal-cual] [--rotulo-tal-cual] [--orden-tal-cual]
                        [--rotulo-apendice]
                        [--datos datos.json] [--titulo "…"] [--autor "…"] … [--fecha 16/10/2026]
                        [--titulo-tal-cual] [--mantener-portada]
    python -m apa_facil mi_trabajo.docx --diagnostico
    python -m apa_facil mi_trabajo.docx --citas
    python -m apa_facil mi_trabajo.docx --referencias
    python -m apa_facil mi_trabajo.docx --cruce

Guarda cada copia en salida/ como nombre_APA.docx y muestra un resumen de lo
que se ha cambiado y de lo que el estudiante debe revisar.

Los datos de la portada llegan en un archivo JSON (--datos) o con opciones
sueltas (--titulo, --autor…), que mandan sobre el archivo dato a dato.

Con --diagnostico no se genera ninguna copia: solo se muestra qué es cada
párrafo (título, cuerpo, lista…) según el clasificador, para revisarlo.

Con --citas tampoco se genera ninguna copia: se muestran las citas que hay en
el texto (autores, año, páginas), tal como se han reconocido. Con --referencias,
las partes de cada referencia de la lista (autores, año, título, fuente…). Con
--cruce, qué citas no tienen referencia, qué referencias no se citan y en cuáles
no coinciden el año o los autores.
"""

import argparse
import glob
import shutil
import sys
from pathlib import Path

from apa_facil import datos_portada, documento, estructura, pipeline
from apa_facil.documento import ErrorDocumento
from apa_facil.reglas import fuente
from apa_facil.revision import citas, cruce, referencias

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
        elif "expected one argument" in message:
            opcion = message.split(":", 1)[0].removeprefix("argument ").split("/")[-1]
            texto = f"Falta indicar el valor después de {opcion}."
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
    ordenes.add_argument("--quitar-numeracion", action="store_true",
                         help="quita la numeración de los títulos («1.1. Objetivos» → «Objetivos»), como pide APA. "
                              "Sin esta opción solo se avisa, porque muchas universidades exigen numerarlos")
    ordenes.add_argument("--titulos-tal-cual", action="store_true",
                         help="deja las mayúsculas de los títulos de los apartados como están. Sin esta opción "
                              "llevan mayúscula inicial en las palabras principales, como pide APA («RESUMEN» → "
                              "«Resumen», «Marco teórico» → «Marco Teórico»), respetando los nombres propios y las "
                              "siglas que el trabajo escribe así en el texto. También los títulos de las tablas")
    ordenes.add_argument("--quitar-introduccion", action="store_true",
                         help="quita el rótulo «Introducción» del primer apartado, que APA no usa (el título del "
                              "trabajo hace sus veces), si el apartado no tiene subapartados. Sin esta opción solo "
                              "se avisa")
    portada = ordenes.add_argument_group(
        "datos de la portada", "En un archivo JSON (--datos) o con estas opciones, que mandan sobre el archivo. "
        "Lo que falte no se rellena: se avisa")
    portada.add_argument("--datos", metavar="datos.json",
                         help="archivo con los datos de la portada (ver el ejemplo en docs/datos_ejemplo.json)")
    portada.add_argument("--titulo", metavar='"…"', help="título del trabajo")
    portada.add_argument("--autor", action="append", metavar='"…"',
                         help="autor o autora; se repite para cada uno (--autor \"Ana Pérez\" --autor \"Luis Gómez\")")
    portada.add_argument("--departamento", metavar='"…"', help="departamento o facultad")
    portada.add_argument("--universidad", metavar='"…"', help="universidad")
    portada.add_argument("--codigo-asignatura", metavar='"…"', help="código de la asignatura, p. ej. \"HIS 201\"")
    portada.add_argument("--asignatura", metavar='"…"', help="nombre de la asignatura")
    portada.add_argument("--profesor", action="append", metavar='"…"',
                         help="profesor o profesora; se repite si son varios")
    portada.add_argument("--fecha", metavar="FECHA",
                         help=f"fecha de entrega: {datos_portada.FORMATOS_FECHA.replace('«', '').replace('»', '')}")
    portada.add_argument("--mantener-portada", action="store_true",
                         help="mantiene la portada que ya tiene el documento (por ejemplo, la que exige tu "
                              "universidad). Sin esta opción, si das tus datos, se sustituye por la portada APA "
                              "de estudiante")
    portada.add_argument("--titulo-tal-cual", action="store_true",
                         help="deja el título de la portada como lo escribes. Sin esta opción lleva mayúscula "
                              "inicial en las palabras principales, como pide APA")
    ordenes.add_argument("--quitar-indice", action="store_true",
                         help="quita el índice, que APA no usa en los trabajos de estudiante. Sin esta opción "
                              "se deja y se avisa, porque muchas universidades lo exigen")
    ordenes.add_argument("--tablas-tal-cual", action="store_true",
                         help="deja los bordes y los colores de las tablas como están. Sin esta opción llevan el "
                              "estilo de APA: solo líneas horizontales (arriba, debajo de los encabezados y abajo), "
                              "sin bordes verticales ni colores de fondo")
    ordenes.add_argument("--rotulo-tal-cual", action="store_true",
                         help="deja el rótulo de la lista de referencias como está («Bibliografía», «Webgrafía»…). "
                              "Sin esta opción pasa a ser «Referencias», como pide APA")
    ordenes.add_argument("--orden-tal-cual", action="store_true",
                         help="deja las referencias en el orden en que están. Sin esta opción se ordenan "
                              "alfabéticamente, como pide APA")
    ordenes.add_argument("--rotulo-apendice", action="store_true",
                         help="cambia «Anexo I» por «Apéndice» («Apéndice A», «B»… si hay varios), como pide APA. "
                              "Sin esta opción solo se avisa, porque muchas universidades piden «Anexo»")
    ordenes.add_argument("--diagnostico", action="store_true",
                         help="no genera ninguna copia: muestra qué es cada párrafo (título, cuerpo, lista…), "
                              "con qué confianza y por qué, para revisar cómo se ha entendido el documento")
    ordenes.add_argument("--citas", action="store_true",
                         help="no genera ninguna copia: muestra las citas que hay en el texto (autores, año y "
                              "páginas) tal como se han reconocido")
    ordenes.add_argument("--referencias", action="store_true",
                         help="no genera ninguna copia: muestra las partes de cada referencia de la lista (autores, "
                              "año, título, fuente, DOI o URL) tal como se han leído, y las que no se han podido leer")
    ordenes.add_argument("--cruce", action="store_true",
                         help="no genera ninguna copia: cruza las citas del texto con la lista de referencias y "
                              "muestra las citas sin referencia, las referencias que no se citan y las que no "
                              "coinciden en el año o en los autores")
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


def diagnosticar(archivos, ancho=None):
    """Muestra la clasificación de los párrafos de cada documento, sin generar copias."""
    ancho = ancho or shutil.get_terminal_size((100, 24)).columns - 1

    def mostrar(doc, nombre):
        clasificados = estructura.clasificar(doc)
        print(f"\nDiagnóstico de {nombre} ({len(clasificados)} párrafos)\n")
        print(estructura.como_texto(clasificados, ancho))
    return _sin_copias(archivos, mostrar, "diagnosticado(s)")


def mostrar_citas(archivos, ancho=None):
    """Muestra las citas del texto de cada documento, sin generar copias (paso 7.1)."""
    ancho = ancho or shutil.get_terminal_size((100, 24)).columns - 1

    def mostrar(doc, nombre):
        print(f"\nCitas en el texto de {nombre}\n")
        print(citas.como_texto(citas.detectar(doc), ancho))
    return _sin_copias(archivos, mostrar, "revisado(s)")


def mostrar_referencias(archivos, ancho=None):
    """Muestra las partes de cada referencia de cada documento, sin generar copias (paso 7.2)."""
    ancho = ancho or shutil.get_terminal_size((100, 24)).columns - 1

    def mostrar(doc, nombre):
        print(f"\nReferencias de {nombre}\n")
        print(referencias.como_texto(referencias.leer(doc), ancho))
    return _sin_copias(archivos, mostrar, "revisado(s)")


def mostrar_cruce(archivos, ancho=None):
    """Cruza las citas con las referencias de cada documento, sin generar copias (paso 7.3)."""
    ancho = ancho or shutil.get_terminal_size((100, 24)).columns - 1

    def mostrar(doc, nombre):
        clasificacion = estructura.clasificar(doc)
        resultado = cruce.cruzar(citas.detectar(doc, clasificacion), referencias.leer(doc, clasificacion))
        print(f"\nCruce de citas y referencias de {nombre}\n")
        print(cruce.como_texto(resultado, ancho))
    return _sin_copias(archivos, mostrar, "revisado(s)")


def _sin_copias(archivos, mostrar, hecho):
    """Abre cada documento y lo muestra con `mostrar(doc, nombre)`, sin generar copias."""
    bien = 0
    for archivo in archivos:
        nombre = Path(archivo).name
        try:
            mostrar(documento.abrir(archivo), nombre)
        except ErrorDocumento as e:
            print(f"✘ {nombre}: {e}")
            continue
        except Exception as e:  # un fallo inesperado no debe parar los demás archivos
            print(f"✘ {nombre}: error inesperado al leerlo ({type(e).__name__}: {e}).")
            continue
        bien += 1

    total = len(archivos)
    print(f"\n{bien} de {total} documento(s) {hecho}. No se ha generado ninguna copia.")
    return 0 if bien == total else 1


def opciones_de_portada(args):
    """{dato: valor} de las opciones sueltas de portada; None si no se usó la opción."""
    return {clave: getattr(args, opcion.removeprefix("--").replace("-", "_"))
            for clave, (opcion, _) in datos_portada.DATOS.items()}


def main(argv=None, carpeta_salida=None):
    configurar_terminal()
    ordenes = crear_ordenes()
    args = ordenes.parse_args(argv)
    opciones = {"fuente": args.fuente} if args.fuente else {}
    for opcion in ("quitar_numeracion", "titulos_tal_cual", "quitar_introduccion", "titulo_tal_cual",
                   "mantener_portada", "quitar_indice", "tablas_tal_cual",
                   "rotulo_tal_cual", "orden_tal_cual", "rotulo_apendice"):
        if getattr(args, opcion):
            opciones[opcion] = True

    archivos = expandir_comodines(args.archivos)
    if args.diagnostico or args.citas or args.referencias or args.cruce:
        modo = ("el modo diagnóstico" if args.diagnostico else "la revisión de citas" if args.citas
                else "la revisión de referencias" if args.referencias else "el cruce de citas y referencias")
        for nombre, usada in (("--fuente", args.fuente), ("--quitar-numeracion", args.quitar_numeracion),
                              ("--titulos-tal-cual", args.titulos_tal_cual),
                              ("--mantener-portada", args.mantener_portada),
                              ("--quitar-indice", args.quitar_indice),
                              ("--tablas-tal-cual", args.tablas_tal_cual),
                              ("--rotulo-tal-cual", args.rotulo_tal_cual),
                              ("--orden-tal-cual", args.orden_tal_cual),
                              ("--rotulo-apendice", args.rotulo_apendice),
                              ("--quitar-introduccion", args.quitar_introduccion),
                              ("--titulo-tal-cual", args.titulo_tal_cual)):
            if usada:
                print(f"Aviso: la opción {nombre} no se usa en {modo}, que no cambia nada.")
        if args.datos or any(v is not None for v in opciones_de_portada(args).values()):
            print(f"Aviso: los datos de la portada no se usan en {modo}, que no cambia nada.")
        resultado = diagnosticar(archivos) if args.diagnostico else 0
        resultado = max(resultado, mostrar_citas(archivos)) if args.citas else resultado
        resultado = max(resultado, mostrar_referencias(archivos)) if args.referencias else resultado
        return max(resultado, mostrar_cruce(archivos)) if args.cruce else resultado

    try:
        portada = datos_portada.combinar(args.datos, opciones_de_portada(args))
    except datos_portada.ErrorDatos as e:
        ordenes.error(f"{e} No se ha procesado ningún documento.")
    if portada:
        opciones["portada"] = portada
        if len(archivos) > 1:
            print(f"Aviso: se usarán los mismos datos de portada en los {len(archivos)} documentos.")
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
