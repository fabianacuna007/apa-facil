"""
Paso 10.5: mide lo que cuesta procesar cada documento, como lo hace el servidor
(en un proceso aparte), y cuántos documentos aguantaría un servidor:

    python herramientas/medir_costes.py                  # los documentos de pruebas/
    python herramientas/medir_costes.py mi_trabajo.docx  # o los que se indiquen

Para cada documento: los segundos, la memoria máxima del proceso y el tamaño.
Al final, la media y cuántos documentos por hora salen con 1 procesador y con los
que procesa a la vez el servidor (APA_SIMULTANEOS). Con eso y el precio del
alojamiento (por hora o por mes) se calcula el coste por documento. No se
guarda nada: las copias van a una carpeta temporal que se borra.
"""

import argparse
import os
import sys
import tempfile
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from servidor import trabajo   # noqa: E402

SIMULTANEOS = int(os.environ.get("APA_SIMULTANEOS", 2))


def medir(ruta):
    """Lo que cuesta el documento: el núcleo solo, y en total (con el arranque del proceso aparte)."""
    inicio = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix="apa_coste_") as carpeta:
        resultado = trabajo.procesar(ruta, Path(carpeta), {}, ruta.name, ruta.stem + "_APA.docx", 300)
    if resultado[0] != "bien":
        return None
    return dict(resultado[4], total=time.perf_counter() - inicio)


def main(argv=None):
    for flujo in (sys.stdout, sys.stderr):
        flujo.reconfigure(encoding="utf-8", errors="replace")
    ordenes = argparse.ArgumentParser(description="Mide lo que cuesta procesar cada documento.")
    ordenes.add_argument("archivos", nargs="*", help="documentos .docx (por defecto, los de pruebas/)")
    args = ordenes.parse_args(argv)
    rutas = [Path(a) for a in args.archivos] or sorted((RAIZ / "pruebas").glob("*.docx"))

    print(f"{'Documento':<34} {'Núcleo':>8} {'Total':>8} {'Memoria':>9} {'Tamaño':>8}")
    tiempos, memorias = [], []
    for ruta in rutas:
        coste = medir(ruta)
        tamano = f"{ruta.stat().st_size // 1024} KB"
        if coste is None:
            print(f"{ruta.name[:34]:<34} {'no se ha podido procesar':>36}")
            continue
        memoria = f"{coste['memoria_mb']:.0f} MB" if coste["memoria_mb"] is not None else "—"
        print(f"{ruta.name[:34]:<34} {coste['segundos']:>7.2f}s {coste['total']:>7.2f}s {memoria:>9} {tamano:>8}")
        tiempos.append(coste["total"])
        if coste["memoria_mb"] is not None:
            memorias.append(coste["memoria_mb"])
    if not tiempos:
        return 1

    media = sum(tiempos) / len(tiempos)
    print()
    print(f"Media: {media:.2f} s por documento en total" + (f", {max(memorias):.0f} MB de memoria como máximo" if memorias else ""))
    print(f"Con 1 procesador: unos {3600 / media:,.0f} documentos por hora.".replace(",", "."))
    print(f"Con {SIMULTANEOS} a la vez (APA_SIMULTANEOS) y {SIMULTANEOS} procesadores: unos "
          f"{SIMULTANEOS * 3600 / media:,.0f} por hora.".replace(",", "."))
    print("Coste por documento = precio del servidor por hora / documentos por hora.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
