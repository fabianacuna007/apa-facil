"""
Saltos de sección a mitad del texto y párrafos partidos (revisión del propietario,
2026-10-20, con un trabajo que era un PDF convertido a Word).

Un PDF convertido a Word (con Word, Acrobat o un conversor en línea) pone un salto
de sección al final de cada página del PDF, y parte en dos los párrafos que
pasaban de una página a otra. Con el interlineado doble de APA el texto de cada
página ya no cabe en una, y el salto obliga a empezar otra: quedaban medias
páginas en blanco por todo el trabajo.

1. Los saltos de sección que caen a mitad del texto pasan a ser continuos (no
   empiezan página). Solo cambia la maquetación: las secciones siguen ahí, con
   sus encabezados. Se mantienen los que van antes de un título de nivel 1, del
   rótulo de las referencias o de un apéndice, los que separan zonas (portada,
   índice, resumen y texto) y los que cambian el tamaño o la orientación de la
   página (una tabla en horizontal).
2. Al final de todo (después de las correcciones de las citas: `unir_partidos`, lo
   llama el pipeline), se une un párrafo con el siguiente si el primero acaba sin
   punto, el segundo empieza en minúscula (o el primero acaba en «el», «de», «y»…,
   con las que nunca acaba un párrafo) y entre los dos solo hay un salto de
   sección y párrafos vacíos (decisión del propietario). Entre los dos se pone un
   espacio; ninguna palabra cambia, y la integridad lo comprueba
   («unir_parrafos»).
"""

import copy

from docx.oxml import OxmlElement

from apa_facil.documento import partidos, texto_parrafo, w

TIPOS_DE_TEXTO = ("cuerpo", "lista", "cita_bloque", "referencia")
# El orden de los hijos de w:sectPr: «type» va antes de todo esto.
DESPUES_DE_TYPE = ("pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
                   "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
                   "printerSettings", "sectPrChange")


def _salto(p):
    return p.find(f"{w('pPr')}/{w('sectPr')}")


def _empieza_pagina(seccion):
    tipo = seccion.find(w("type"))
    return tipo is None or tipo.get(w("val")) != "continuous"


def _pagina(seccion):
    """(ancho, alto) del papel de la sección."""
    tamano = seccion.find(w("pgSz")) if seccion is not None else None
    return (None, None) if tamano is None else (tamano.get(w("w")), tamano.get(w("h")))


def _hacer_continuo(seccion):
    tipo = seccion.find(w("type"))
    if tipo is None:
        tipo = OxmlElement("w:type")
        siguiente = next((h for h in seccion if h.tag in [w(e) for e in DESPUES_DE_TYPE]), None)
        if siguiente is not None:
            siguiente.addprevious(tipo)
        else:
            seccion.append(tipo)
    tipo.set(w("val"), "continuous")


def _abre_algo(c):
    """¿Empieza aquí algo que va en página nueva? Un título de nivel 1, las referencias o un apéndice."""
    return c.tipo in ("titulo_referencias", "anexo") or (c.tipo == "titulo" and c.nivel == 1)


def aplicar(doc, opciones, informe, estructura=None):
    if estructura is None:
        return
    con_texto = [c for c in estructura if c.texto.strip()]
    posicion = {c.elemento: i for i, c in enumerate(estructura)}
    secciones = [c for c in estructura if _salto(c.elemento) is not None]
    siguiente_seccion = {s.elemento: (secciones[k + 1].elemento if k + 1 < len(secciones) else None)
                         for k, s in enumerate(secciones)}
    final = doc.element.body.find(w("sectPr"))
    continuos = 0
    for s in secciones:
        seccion = _salto(s.elemento)
        if not _empieza_pagina(seccion):
            continue
        antes = next((c for c in reversed(con_texto) if posicion[c.elemento] < posicion[s.elemento]), None)
        despues = next((c for c in con_texto if posicion[c.elemento] > posicion[s.elemento]), None)
        otra = siguiente_seccion[s.elemento]
        pagina_siguiente = _pagina(_salto(otra) if otra is not None else final)
        if (antes is None or despues is None or antes.zona != despues.zona or _abre_algo(despues)
                or _pagina(seccion) != pagina_siguiente):
            continue
        _hacer_continuo(seccion)
        continuos += 1
    if continuos:
        varios = continuos > 1
        informe.cambio(f"Saltos de sección: {continuos} salto{'s' if varios else ''} de sección que caía{'n' if varios else ''} "
                       "a mitad del texto ya no empieza" + ("n" if varios else "") + " una página nueva, para que "
                       "no queden páginas a medias. Los de antes de un apartado principal, de las referencias y de los "
                       "apéndices se mantienen.")
    if estructura.convertido_de_pdf:
        # El conversor separa o junta las letras (espaciado entre caracteres) para imitar el texto
        # justificado del PDF: con el texto a la izquierda, solo deja espacios raros entre las palabras.
        intactos = getattr(estructura, "intactos", frozenset())
        espaciados = 0
        for p in doc.element.body.iter(w("p")):
            if p in intactos:
                continue
            for rpr in p.iter(w("rPr")):
                for etiqueta in ("spacing", "w"):
                    el = rpr.find(w(etiqueta))
                    if el is not None and el.get(w("val")) not in (None, "0", "100"):
                        rpr.remove(el)
                        espaciados += 1
        if espaciados:
            informe.cambio("Espaciado entre letras: se ha quitado el que el conversor de PDF puso para imitar el texto "
                           "justificado (separaba o juntaba las letras de algunas líneas).")
        informe.aviso("Tu documento parece un PDF convertido a Word (tiene una sección por página). Revisa con "
                      "especial cuidado el resultado: en estos documentos la sangría de cada párrafo sale de la "
                      "posición del texto en el PDF, así que la herramienta no puede saber qué párrafos son citas en "
                      "bloque (los ha tratado como texto normal), y las tablas, las imágenes y los índices pueden "
                      "necesitar retoques. Si tienes el documento original de Word, procesa ese.")


def _espacio_como(r):
    """Un fragmento con un espacio y el mismo formato que `r`."""
    nuevo = OxmlElement("w:r")
    if r is not None and r.find(w("rPr")) is not None:
        nuevo.append(copy.deepcopy(r.find(w("rPr"))))
    t = OxmlElement("w:t")
    t.text = " "
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    nuevo.append(t)
    return nuevo


def unir_partidos(doc, estructura, informe):
    """Une los párrafos que un salto de sección partió en dos (ver la explicación del módulo)."""
    unidos = []
    cuerpo = doc.element.body
    for s in [c for c in estructura if _salto(c.elemento) is not None]:
        bloques = list(cuerpo)
        if s.elemento not in bloques:
            continue
        i = bloques.index(s.elemento)
        a = next((b for b in reversed(bloques[:i]) if b.tag != w("p") or texto_parrafo(b).strip()), None)
        b = next((b for b in bloques[i + 1:] if b.tag != w("p") or texto_parrafo(b).strip()), None)
        if a is None or b is None or a.tag != w("p") or b.tag != w("p"):
            continue   # una tabla en medio: no se une
        clase = {c.elemento: c for c in estructura}
        ca, cb = clase.get(a), clase.get(b)
        if ca is None or cb is None or ca.tipo not in TIPOS_DE_TEXTO or cb.tipo not in TIPOS_DE_TEXTO:
            continue
        entre = [texto_parrafo(x) for x in bloques[bloques.index(a) + 1:bloques.index(b)]]
        if not partidos(texto_parrafo(a), texto_parrafo(b), entre):
            continue
        ultimo = next((r for r in reversed(list(a.iter(w("r")))) if r.getparent() is a), None)
        a.append(_espacio_como(ultimo))
        for hijo in list(b):
            if hijo.tag != w("pPr"):
                a.append(hijo)
        b.getparent().remove(b)
        informe.cambio_previsto(ca.numero, "unir_parrafos", desde=(cb.numero,))
        unidos.append(texto_parrafo(a))
    if unidos:
        varios = len(unidos) > 1
        informe.cambio(f"Párrafos partidos: se ha{'n' if varios else ''} unido {len(unidos)} párrafo{'s' if varios else ''} "
                       "que estaba" + ("n" if varios else "") + " partido" + ("s" if varios else "") + " en dos por un "
                       "salto de sección (el primero se cortaba a mitad de frase). Entre los dos "
                       "se ha puesto un espacio; el texto no cambia. Por ejemplo: «"
                       + _extracto(unidos[0]) + "».")


def _extracto(texto, largo=70):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"
