"""
Inspector: muestra el formato actual de un documento Word, sin cambiar nada.

Para cada documento indica los márgenes por sección, las fuentes usadas, los
interlineados, las sangrías, la alineación, los estilos de párrafo y lo que
hay en encabezados y pies. Sirve para diagnosticar los documentos de prueba.

Uso:
    python herramientas/inspeccionar.py archivo.docx [otro.docx ...]
"""

import sys
from collections import Counter
from pathlib import Path

# Para encontrar el paquete apa_facil al ejecutar este archivo directamente.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apa_facil import documento  # noqa: E402
from apa_facil.__main__ import configurar_terminal  # noqa: E402
from apa_facil.documento import ErrorDocumento, Estilos, texto_legible, texto_parrafo, w  # noqa: E402

TWIPS_POR_CM = 567     # Word mide casi todo en "twips": veinteavos de punto


def cm(twips):
    return f"{int(twips) / TWIPS_POR_CM:.2f}".replace(".", ",")


def numero(x):
    """1.5 → "1,5"; 2.0 → "2"."""
    return f"{x:g}".replace(".", ",")


# --- Recogida de datos ----------------------------------------------------------

def describir_interlineado(estilos, p):
    linea = estilos.parrafo(p, "spacing", "line")
    regla = estilos.parrafo(p, "spacing", "lineRule") or "auto"
    if linea is None:
        return "sencillo (por defecto)"
    if regla == "exact":
        return f"exacto {numero(int(linea) / 20)} pt"
    if regla == "atLeast":
        return f"mínimo {numero(int(linea) / 20)} pt"
    veces = round(int(linea) / 240, 2)
    return {1: "sencillo", 1.15: "1,15", 1.5: "1,5", 2: "doble"}.get(veces, f"{numero(veces)} líneas")


def describir_espaciado(estilos, p):
    antes = estilos.parrafo(p, "spacing", "before") or "0"
    despues = estilos.parrafo(p, "spacing", "after") or "0"
    return f"antes {numero(int(antes) / 20)} pt · después {numero(int(despues) / 20)} pt"


def describir_sangria(estilos, p):
    partes = []
    izquierda = estilos.parrafo(p, "ind", "left") or estilos.parrafo(p, "ind", "start")
    francesa = estilos.parrafo(p, "ind", "hanging")
    primera = estilos.parrafo(p, "ind", "firstLine")
    if izquierda and int(izquierda):
        partes.append(f"izquierda {cm(izquierda)} cm")
    if francesa and int(francesa):
        partes.append(f"francesa {cm(francesa)} cm")
    elif primera and int(primera):
        partes.append(f"primera línea {cm(primera)} cm")
    return " + ".join(partes) or "sin sangría"


ALINEACIONES = {"left": "izquierda", "start": "izquierda", "both": "justificada",
                "distribute": "justificada", "center": "centrada", "right": "derecha", "end": "derecha"}


def describir_encabezados(doc):
    """Encabezados y pies que se ven en cada sección."""
    pares = doc.settings.odd_and_even_pages_header_footer
    resultado = []
    for i, seccion in enumerate(doc.sections, start=1):
        tipos = [("Encabezado", seccion.header, True), ("Pie", seccion.footer, True)]
        if seccion.different_first_page_header_footer:
            tipos += [("Encabezado de la primera página", seccion.first_page_header, True),
                      ("Pie de la primera página", seccion.first_page_footer, True)]
        if pares:
            tipos += [("Encabezado de páginas pares", seccion.even_page_header, True),
                      ("Pie de páginas pares", seccion.even_page_footer, True)]
        for tipo, parte, _ in tipos:
            if parte.is_linked_to_previous:
                estado = "no tiene" if i == 1 else "igual que en la sección anterior"
                resultado.append({"seccion": i, "tipo": tipo, "texto": [], "estado": estado})
                continue
            lineas = [texto_legible(p).strip() for p in parte._element.iter(w("p"))]
            lineas = [l for l in lineas if l]
            imagenes = sum(1 for _ in parte._element.iter(w("drawing"), w("pict")))
            estado = "vacío" if not lineas and not imagenes else ""
            if imagenes:
                lineas.append(f"[{imagenes} imagen(es)]")
            resultado.append({"seccion": i, "tipo": tipo, "texto": lineas, "estado": estado})
    return resultado


def inspeccionar(ruta):
    """Reúne los datos de formato de un documento. Devuelve un diccionario."""
    doc = documento.abrir(ruta)
    estilos = Estilos(doc)
    datos = {
        "archivo": Path(ruta).name,
        "secciones": [],
        "fuentes": Counter(),        # (fuente, tamaño) → número de letras
        "interlineados": Counter(),  # descripción → número de párrafos
        "espaciados": Counter(),
        "sangrias": Counter(),
        "alineaciones": Counter(),
        "estilos": Counter(),
        "parrafos": 0,
        "parrafos_en_tablas": 0,
        "parrafos_vacios": 0,
    }

    for i, s in enumerate(doc.sections, start=1):
        datos["secciones"].append({
            "numero": i,
            "orientacion": "horizontal" if s.page_width > s.page_height else "vertical",
            "pagina": f"{cm(s.page_width.twips)} × {cm(s.page_height.twips)} cm",
            "margenes": {"superior": cm(s.top_margin.twips), "inferior": cm(s.bottom_margin.twips),
                         "izquierdo": cm(s.left_margin.twips), "derecho": cm(s.right_margin.twips)},
        })

    for p in doc.element.body.iter(w("p")):
        for r in p.iter(w("r")):
            letras = sum(len(t.text or "") for t in r.iter(w("t")))
            if letras:
                datos["fuentes"][(estilos.fuente(r, p), estilos.tamano(r, p))] += letras

        if not texto_parrafo(p).strip():
            datos["parrafos_vacios"] += 1
            continue
        if next(p.iterancestors(w("tc")), None) is not None:
            datos["parrafos_en_tablas"] += 1  # las tablas tienen sus propias reglas
            continue
        datos["parrafos"] += 1
        datos["interlineados"][describir_interlineado(estilos, p)] += 1
        datos["espaciados"][describir_espaciado(estilos, p)] += 1
        datos["sangrias"][describir_sangria(estilos, p)] += 1
        alineacion = estilos.parrafo(p, "jc", "val")
        datos["alineaciones"][ALINEACIONES.get(alineacion, "izquierda (por defecto)"
                                               if alineacion is None else alineacion)] += 1
        # Sin estilo asignado, Word lo muestra como "Normal".
        datos["estilos"][estilos.nombre(estilos.estilo_de(p)) or "Normal"] += 1

    datos["encabezados"] = describir_encabezados(doc)
    return datos


# --- Presentación ---------------------------------------------------------------

def _lista(contador, unidad):
    total = sum(contador.values()) or 1
    return [f"   {n:>4} {unidad} ({n * 100 / total:.0f} %)  {clave}" for clave, n in contador.most_common()]


def mostrar(datos):
    """Convierte los datos en un texto legible para la terminal."""
    L = ["=" * 72, f"📄 {datos['archivo']}", "=" * 72]

    L.append("\nMárgenes")
    for s in datos["secciones"]:
        m = s["margenes"]
        L.append(f"   Sección {s['numero']} ({s['orientacion']}, {s['pagina']}): "
                 f"superior {m['superior']} · inferior {m['inferior']} · "
                 f"izquierdo {m['izquierdo']} · derecho {m['derecho']} cm")

    L.append("\nFuentes (por cantidad de letras, incluidas las tablas)")
    total = sum(datos["fuentes"].values()) or 1
    for (fuente, tamano), letras in datos["fuentes"].most_common():
        L.append(f"   {letras * 100 / total:5.1f} %  {fuente} {numero(tamano)} pt".replace(".", ",", 1))

    L.append(f"\nPárrafos con texto: {datos['parrafos']} en el cuerpo, "
             f"{datos['parrafos_en_tablas']} dentro de tablas, y {datos['parrafos_vacios']} vacíos.")
    L.append("Los apartados siguientes cuentan solo los del cuerpo.")
    for titulo, clave in (("Interlineado", "interlineados"), ("Espacio entre párrafos", "espaciados"),
                          ("Sangría", "sangrias"), ("Alineación", "alineaciones"),
                          ("Estilos de párrafo", "estilos")):
        L.append(f"\n{titulo}")
        L += _lista(datos[clave], "párr.")

    L.append("\nEncabezados y pies de página")
    for e in datos["encabezados"]:
        if e["texto"]:
            L.append(f"   Sección {e['seccion']} · {e['tipo']}: " + " | ".join(
                t.replace("\t", " → ").replace("\n", " ⏎ ") for t in e["texto"]))
        else:
            L.append(f"   Sección {e['seccion']} · {e['tipo']}: ({e['estado']})")
    return "\n".join(L)


def main(argv=None):
    configurar_terminal()
    archivos = sys.argv[1:] if argv is None else argv
    if not archivos:
        print("Indica qué documento inspeccionar. Ejemplo: "
              "python herramientas/inspeccionar.py pruebas/01_ensayo_basico.docx")
        return 2
    errores = 0
    for archivo in archivos:
        try:
            print(mostrar(inspeccionar(archivo)))
        except ErrorDocumento as e:
            print(f"✘ {Path(archivo).name}: {e}")
            errores += 1
        print()
    return 1 if errores else 0


if __name__ == "__main__":
    sys.exit(main())
