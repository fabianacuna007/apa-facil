"""
Paso 3.6: problemas que solo se avisan, sin corregirlos (corregirlos
cambiaría el texto del estudiante).

- Saltos de línea manuales (Mayús+Intro) dentro de un párrafo. Los saltos de
  página y de columna no cuentan.
- Viñetas escritas a mano ("•", "-"…). Las listas numeradas a mano no se
  avisan: APA admite listas numeradas.
- Dos o más espacios seguidos entre palabras. No cuentan los del principio (los
  quita la limpieza) ni los que van detrás de una viñeta o un número de lista.
- Texto en otra fuente, que suele venir de pegarlo de una web o de otro
  documento: párrafos de texto, listas, citas y referencias con al menos 20
  letras en una fuente distinta de la principal. No cuentan los títulos, las
  leyendas ni el código.

Va la primera del pipeline: después, la regla de fuente lo pone todo en la
misma fuente y ya no se vería qué estaba en otra. Como la limpieza, no mira la
portada (paso 5.3) ni las tablas.
"""

import re
from collections import Counter

from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import LISTA_A_MANO, VINETA, clasificar
from apa_facil.reglas import fuente

LETRAS_EN_OTRA_FUENTE = 20
TIPOS_DE_TEXTO = ("cuerpo", "lista", "cita_bloque", "referencia")
DOBLE_ESPACIO = re.compile(r"(\S+) {2,}(\S+)")
EJEMPLOS = 2


def aplicar(doc, opciones, informe, estructura=None):
    estilos = Estilos(doc)
    clasificados = estructura if estructura is not None else clasificar(doc, estilos)
    revisados = [c for c in clasificados if c.zona != "portada"]
    _saltos_de_linea(revisados, informe)
    _vinetas_a_mano(revisados, informe)
    _dobles_espacios(revisados, informe)
    _otra_fuente(revisados, estilos, informe)


# --- Saltos de línea manuales -----------------------------------------------------------

def _saltos_manuales(p):
    return [b for b in p.iter(w("br"), w("cr"))
            if next(b.iterancestors(w("p")), None) is p and b.get(w("type")) in (None, "textWrapping")]


def _saltos_de_linea(revisados, informe):
    con_saltos = [c for c in revisados if c.texto and _saltos_manuales(c.elemento)]
    if not con_saltos:
        return
    ejemplos = []
    for c in con_saltos[:EJEMPLOS]:
        texto = texto_parrafo(c.elemento)
        donde = texto.index("\n")
        ejemplos.append(f"«{_alrededor(texto, donde, donde + 1, ' ⏎ ')}»")
    n = len(con_saltos)
    informe.aviso(f"{_parrafos(n)} {'tiene' if n == 1 else 'tienen'} saltos de línea manuales (Mayús+Intro) "
                  f"dentro del texto ({_y(ejemplos, n)}). En APA el texto de un párrafo va seguido: quita el "
                  "salto, o pulsa Intro si querías empezar un párrafo nuevo.")


# --- Viñetas escritas a mano ------------------------------------------------------------

def _vinetas_a_mano(revisados, informe):
    lineas = [c for c in revisados if c.tipo == "lista" and VINETA.match(texto_parrafo(c.elemento))]
    if not lineas:
        return
    simbolos = []
    for c in lineas:
        simbolo = texto_parrafo(c.elemento).strip()[0]
        if simbolo not in simbolos:
            simbolos.append(simbolo)
    n = len(lineas)
    mezcla = (f" Además, se mezclan símbolos distintos ({_lista_simple([f'«{s}»' for s in simbolos])})."
              if len(simbolos) > 1 else "")
    informe.aviso(f"{'1 línea empieza' if n == 1 else f'{n} líneas empiezan'} con una viñeta escrita a mano "
                  f"(«{_extracto(lineas[0].texto)}»).{mezcla} Es mejor usar la lista con viñetas de Word "
                  "(Inicio › Viñetas): así la sangría y el espaciado quedan bien. No se ha cambiado porque "
                  "habría que modificar el texto.")


# --- Dobles espacios --------------------------------------------------------------------

def _dobles_espacios(revisados, informe):
    sitios = []
    for c in revisados:
        texto = texto_parrafo(c.elemento).strip()
        marca = LISTA_A_MANO.match(texto)
        if marca:
            texto = texto[marca.end():].lstrip()
        sitios += [(encontrado.group(1), encontrado.group(2)) for encontrado in DOBLE_ESPACIO.finditer(texto)]
    if not sitios:
        return
    n = len(sitios)
    ejemplos = [f"entre «{antes}» y «{despues}»" for antes, despues in sitios[:EJEMPLOS]]
    informe.aviso(f"Hay {'1 sitio' if n == 1 else f'{n} sitios'} con dos o más espacios seguidos entre "
                  f"palabras ({_y(ejemplos, n)}). APA usa un solo espacio, también después del punto. En Word "
                  "puedes corregirlos todos a la vez con Ctrl+L: busca dos espacios y reemplázalos por uno.")


# --- Texto en otra fuente ---------------------------------------------------------------

def _familia(nombre):
    """Las variantes de una familia recomendada ("Calibri Light") cuentan como la misma fuente."""
    return fuente.familia(nombre) or nombre


def _letras_por_fuente(c, estilos):
    letras = Counter()
    for r in fuente._runs_propios(c.elemento):
        nombre = estilos.fuente(r, c.elemento)
        if not fuente.es_de_simbolos(nombre):
            letras[_familia(nombre)] += fuente._letras(r)
    return letras


def _otra_fuente(revisados, estilos, informe):
    por_parrafo = {c.numero: _letras_por_fuente(c, estilos) for c in revisados}
    total = sum(por_parrafo.values(), Counter())
    if not total:
        return
    principal = total.most_common(1)[0][0]

    pegados = []
    for c in revisados:
        if c.tipo not in TIPOS_DE_TEXTO or fuente.es_codigo(c.elemento, estilos):
            continue
        otras = Counter({nombre: n for nombre, n in por_parrafo[c.numero].items() if nombre != principal})
        if otras and otras.most_common(1)[0][1] >= LETRAS_EN_OTRA_FUENTE:
            pegados.append((otras.most_common(1)[0][0], c))
    if not pegados:
        return
    n = len(pegados)
    ejemplos = [f"{nombre} en «{_extracto(c.texto, 25)}»" for nombre, c in pegados[:EJEMPLOS]]
    informe.aviso(f"{_parrafos(n)} {'estaba' if n == 1 else 'estaban'} en otra fuente, distinta de la del resto "
                  f"del trabajo ({principal}): {_y(ejemplos, n)}. Suele pasar al pegar texto de una página web o "
                  f"de otro documento. La fuente ya se ha igualado; comprueba que ese texto está citado y que se "
                  "ve bien.")


# --- Mensajes ---------------------------------------------------------------------------

def _parrafos(n):
    return "1 párrafo" if n == 1 else f"{n} párrafos"


def _extracto(texto, largo=35):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo].rstrip() + "…"


def _alrededor(texto, desde, hasta, en_medio, margen=20):
    """El texto de alrededor de texto[desde:hasta], con `en_medio` en su lugar."""
    antes, despues = texto[max(0, desde - margen):desde], texto[hasta:hasta + margen]
    cortado_antes, cortado_despues = desde > margen, hasta + margen < len(texto)
    if cortado_antes and " " in antes.strip():
        antes = antes.strip().split(" ", 1)[1]          # sin la palabra que ha quedado a medias
    if cortado_despues and " " in despues.strip():
        despues = despues.strip().rsplit(" ", 1)[0]
    antes, despues = " ".join(antes.split()), " ".join(despues.split())
    return ("…" if cortado_antes else "") + antes + en_medio + despues + ("…" if cortado_despues else "")


def _lista_simple(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


def _y(ejemplos, total):
    """Los ejemplos, y cuántos más hay."""
    resto = total - len(ejemplos)
    return _lista_simple(ejemplos + ([f"{resto} más"] if resto else []))
