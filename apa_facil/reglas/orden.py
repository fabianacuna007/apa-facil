"""
Paso 6.5: el orden alfabético de las referencias.

Según la web oficial de APA (Reference list setup y Citing works with the same
author and date):
- las referencias se ordenan por su primera palabra, normalmente el apellido del
  primer autor; una obra sin autor, por su título (que es lo primero);
- las del mismo autor y el mismo año llevan una letra tras el año («2018a»,
  «2018b»), asignada por el orden de los títulos.

El resto lo remite al Manual; aquí se aplica lo que sale de la regla general:
palabra a palabra, de modo que «nada va antes que algo» («Brown» antes que
«Browning») y la obra de un solo autor va antes que la escrita con otros; luego
el año (sin fecha, «s. f.», primero) y luego el título. Las tildes no cuentan y
la «ñ» va después de la «n», como en el alfabeto español.

Reordenar es por defecto y queda en el informe (decisión del propietario);
--orden-tal-cual deja el orden y avisa. Una línea que solo es una dirección web
no es una referencia completa: va al final, con aviso. Con varias listas no se
ordena nada (hay que unirlas: lo avisa también el paso 6.4).
"""

import re
import unicodedata

from lxml import etree

from apa_facil.documento import texto_parrafo
from apa_facil.estructura import clasificar
from apa_facil.reglas.leyendas import _extracto, _lista, _plural
from apa_facil.reglas.referencias import listas

# El año de la referencia: «(2019)», «(2019a)», «(2019, 3 de mayo)», «(s. f.)»; o, sin paréntesis, «. 2019.».
FECHA = re.compile(r"\(\s*((?:1[5-9]|20)\d{2}[a-z]?|s\.\s?f\.(?:-[a-z])?|n\.\s?d\.(?:-[a-z])?)[^)]*\)"
                   r"|(?<![\w./:-])((?:1[5-9]|20)\d{2}[a-z]?)(?![\w/])")
SOLO_UNA_URL = re.compile(r"^\s*(https?://|www\.)\S+\s*$", re.IGNORECASE)
CONJUNCIONES = {"y", "e", "and"}   # entre autores: no cuentan para el orden


def _letras(texto):
    """En minúsculas, sin tildes, y la «ñ» después de la «n» (n + un carácter que va detrás de la z)."""
    texto = unicodedata.normalize("NFD", texto.lower().replace("ñ", "n\x7f"))
    return "".join(c for c in texto if not unicodedata.combining(c))


def _palabras(texto, sin_conjunciones=False):
    palabras = re.findall(r"[a-z0-9\x7f]+", _letras(texto))
    return tuple(p for p in palabras if not (sin_conjunciones and p in CONJUNCIONES))


def partes(texto):
    """(autores, año, título y resto) de una referencia; sin año, todo es «autores» (o el título, si no tiene autor)."""
    fecha = FECHA.search(texto)
    if fecha is None:
        return texto, "", ""
    return texto[:fecha.start()], fecha.group(1) or fecha.group(2), texto[fecha.end():]


def clave(texto):
    """La clave de orden de una referencia (las direcciones web sueltas, al final)."""
    if SOLO_UNA_URL.match(texto):
        return (1, (), (), ())
    autores, anio, resto = partes(texto)
    anio = _letras(anio)
    del_anio = (0, anio) if anio.startswith(("s", "n")) else (1, anio) if anio else (2, "")
    return (0, _palabras(autores, sin_conjunciones=bool(anio)), del_anio, _palabras(resto))


def _mismo_autor_y_anio(textos):
    """Las referencias con los mismos autores y el mismo año sin letra («2018», «2018»)."""
    vistas, repetidas = {}, []
    for texto in textos:
        autores, anio, _ = partes(texto)
        if anio and anio[-1].isdigit():
            k = (_palabras(autores, sin_conjunciones=True), anio)
            if k in vistas:
                repetidas.append(texto)
            vistas[k] = texto
    return repetidas


def _reordenar(elementos, nuevos):
    """Pone los párrafos `nuevos` en los sitios que ocupaban `elementos`, en ese orden."""
    marcas = []
    for el in elementos:
        marca = etree.Element("marca")
        el.addprevious(marca)
        marcas.append(marca)
    for el in elementos:
        el.getparent().remove(el)
    for marca, el in zip(marcas, nuevos):
        marca.addnext(el)
        marca.getparent().remove(marca)


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    todas = listas(clasificados)
    if not todas:
        return
    if len(todas) > 1:
        informe.aviso(f"Como hay {len(todas)} listas de referencias, no se han ordenado: cuando las unas en una sola, "
                      "la herramienta las ordenará alfabéticamente.")
        return
    _, referencias = todas[0]
    textos = [texto_parrafo(c.elemento) for c in referencias]
    ordenadas = sorted(referencias, key=lambda c: clave(texto_parrafo(c.elemento)))

    urls = [t for t in textos if SOLO_UNA_URL.match(t)]
    if urls:
        informe.aviso(f"{_lista(urls)} {'es' if len(urls) == 1 else 'son'} solo una dirección web: APA pide la "
                      "referencia completa (autor, fecha, título y dónde se publicó). "
                      + ("En la lista va al final, porque no se puede ordenar: complétala tú."
                         if not opciones.get("orden_tal_cual") else "Complétala tú."))
    repetidas = _mismo_autor_y_anio(textos)
    if repetidas:
        informe.aviso(f"Hay obras de los mismos autores y el mismo año ({_lista(repetidas)}): APA las distingue con una "
                      "letra tras el año («2018a», «2018b»), por el orden de sus títulos, también en las citas del texto. "
                      "Añádela tú.")
    if ordenadas == referencias:
        return
    if opciones.get("orden_tal_cual"):
        informe.aviso(f"Las referencias no están en orden alfabético (la primera debería ser "
                      f"«{_extracto(texto_parrafo(ordenadas[0].elemento))}»): APA las ordena por el apellido del primer "
                      "autor. Se han dejado como estaban porque has usado la opción --orden-tal-cual.")
        return
    fuera_de_sitio = sum(1 for a, b in zip(referencias, ordenadas) if a is not b)
    _reordenar([c.elemento for c in referencias], [c.elemento for c in ordenadas])
    informe.cambio_previsto(min(c.numero for c in referencias), "reordenar_referencias",
                            desde=tuple(c.numero for c in ordenadas))
    informe.cambio(f"Referencias: se han ordenado alfabéticamente, como pide APA (por el apellido del primer autor; las "
                   f"obras sin autor, por el título): {_plural(fuera_de_sitio, 'referencia')} "
                   f"{'ha' if fuera_de_sitio == 1 else 'han'} cambiado de sitio, y ahora la primera es "
                   f"«{_extracto(texto_parrafo(ordenadas[0].elemento))}». Si prefieres el orden que tenían, usa la "
                   "opción --orden-tal-cual.")
