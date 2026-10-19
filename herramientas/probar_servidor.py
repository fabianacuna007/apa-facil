"""
Envía un Word al servidor y guarda lo que devuelve en salida/:

    python herramientas/probar_servidor.py mi_trabajo.docx [--fuente Arial] [--comentarios] …
                                           [--url http://localhost:8000]

Guarda la copia corregida (nombre_APA.docx), el informe completo (nombre_APA.json)
y el resumen legible (nombre_APA.txt). Antes hay que arrancar el servidor:

    uvicorn servidor.app:app
"""

import argparse
import base64
import json
import sys
from pathlib import Path

import httpx

RAIZ = Path(__file__).resolve().parent.parent
BANDERAS = ("quitar_numeracion", "titulos_tal_cual", "quitar_introduccion", "titulo_tal_cual",
            "mantener_portada", "quitar_indice", "tablas_tal_cual", "rotulo_tal_cual", "orden_tal_cual",
            "rotulo_apendice", "comentarios")


def main(argv=None):
    for flujo in (sys.stdout, sys.stderr):
        flujo.reconfigure(encoding="utf-8", errors="replace")
    ordenes = argparse.ArgumentParser(description="Prueba el servidor de APA Fácil con un documento.")
    ordenes.add_argument("archivo")
    ordenes.add_argument("--url", default="http://localhost:8000")
    ordenes.add_argument("--salida", default=str(RAIZ / "salida"))
    ordenes.add_argument("--fuente")
    for bandera in BANDERAS:
        ordenes.add_argument("--" + bandera.replace("_", "-"), action="store_true")
    args = ordenes.parse_args(argv)

    archivo = Path(args.archivo)
    campos = {b: "true" for b in BANDERAS if getattr(args, b)}
    if args.fuente:
        campos["fuente"] = args.fuente
    try:
        with archivo.open("rb") as f:
            respuesta = httpx.post(f"{args.url}/procesar", data=campos, timeout=300,
                                   files={"archivo": (archivo.name, f)})
    except httpx.ConnectError:
        print(f"No se puede conectar con el servidor en {args.url}. Arráncalo con: uvicorn servidor.app:app")
        return 1
    datos = respuesta.json()
    if respuesta.status_code != 200:
        print(f"✘ {archivo.name}: {datos['error']}")
        return 1

    carpeta = Path(args.salida)
    carpeta.mkdir(parents=True, exist_ok=True)
    copia = carpeta / datos["documento"]["nombre"]
    copia.write_bytes(base64.b64decode(datos["documento"]["contenido"]))
    copia.with_suffix(".json").write_text(json.dumps(datos["informe"], ensure_ascii=False, indent=2) + "\n",
                                          encoding="utf-8")
    copia.with_suffix(".txt").write_text(datos["resumen"] + "\n", encoding="utf-8")
    print(f"✔ {archivo.name} → {copia}")
    print(datos["resumen"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
