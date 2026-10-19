"""
Paso 7.6: los avisos de citas y referencias (paso 7.5), como comentarios de
Word en la copia, cada uno en su sitio. Solo con la opción --comentarios
(decisión del propietario): los comentarios van dentro del trabajo que se
entrega, y si el estudiante olvida borrarlos, el profesor los ve.

Un comentario por caso, de «APA Fácil», con el mensaje del aviso, la explicación
propia del caso y la corrección. Se ponen al final, después de las reglas: cada
caso recuerda su párrafo (el objeto del documento, no su número), que sigue
siendo el mismo aunque las reglas lo hayan movido (las referencias reordenadas
se llevan su comentario). Dentro del párrafo se busca el texto del caso; si una
referencia ya no tiene ese texto, se marca entera, y si una cita no se encuentra
o el párrafo ya no está, el caso se queda sin comentario (nunca se adivina el
sitio). Las notas al pie no llevan comentarios: el aviso sigue en el informe.

El texto no cambia: el comentario va en otra parte del archivo, y en el texto
solo quedan sus marcas (sin letras), así que la integridad no lo nota.
"""

from docx.text.paragraph import Paragraph
from docx.text.run import Run

from apa_facil.documento import aislar, texto_parrafo

AUTOR, INICIALES = "APA Fácil", "AF"


def comentar(doc, revision, informe):
    """Pone en `doc` un comentario por cada caso de los avisos de la revisión. Devuelve cuántos ha puesto."""
    puestos, en_notas, sin_sitio = 0, 0, 0
    for mensaje, caso, ancla in revision.anclas:
        if ancla.elemento is None:
            en_notas += 1
            continue
        sitio = _sitio(doc, ancla)
        fragmentos = aislar(ancla.elemento, *sitio) if sitio else None
        if not fragmentos:
            sin_sitio += 1
            continue
        parrafo = Paragraph(ancla.elemento, doc._body)
        doc.add_comment([Run(fragmentos[0], parrafo), Run(fragmentos[1], parrafo)], text=texto_del_comentario(
            mensaje, caso), author=AUTOR, initials=INICIALES)
        puestos += 1
    if puestos:
        informe.cambio(f"Citas y referencias: se han puesto {_plural(puestos, 'comentario')} en la copia, en el sitio "
                       "de cada aviso. Bórralos antes de entregar el trabajo (en Word: Revisar › Eliminar › Eliminar "
                       "todos los comentarios del documento).")
    if en_notas or sin_sitio:
        partes = []
        if en_notas:
            partes.append(f"{_plural(en_notas, 'caso')} de las notas al pie (en ellas no se ponen comentarios)")
        if sin_sitio:
            partes.append(f"{_plural(sin_sitio, 'caso')} cuyo texto ya no se encuentra en la copia")
        informe.aviso("No llevan comentario " + " ni ".join(partes) + ": están en este informe.")
    return puestos


def texto_del_comentario(mensaje, caso):
    texto = f"{AUTOR}: {mensaje}"
    if caso.get("explicacion"):
        explicacion = caso["explicacion"]
        texto += f" En este caso, {explicacion[0].lower()}{explicacion[1:]}" + ("" if explicacion.endswith(".") else ".")
    if caso.get("correccion"):
        texto += f" Corrección: «{caso['correccion']}»."
    return texto


def _sitio(doc, ancla):
    """(inicio, fin) del texto que se marca en el párrafo, o None."""
    if doc.element.body not in ancla.elemento.iterancestors():
        return None   # las reglas han quitado el párrafo
    texto = texto_parrafo(ancla.elemento)
    if ancla.buscar:
        posiciones = _todas(texto, ancla.buscar)
        if posiciones:
            if ancla.cerca is not None:
                inicio = min(posiciones, key=lambda i: abs(i - ancla.cerca))
            else:
                inicio = posiciones[-1] if ancla.ultima else posiciones[0]
            return inicio, inicio + len(ancla.buscar)
        if not ancla.referencia:
            return None
    principio = len(texto) - len(texto.lstrip())
    fin = len(texto.rstrip())
    return (principio, fin) if fin > principio else None


def _todas(texto, buscado):
    posiciones, i = [], texto.find(buscado)
    while i >= 0:
        posiciones.append(i)
        i = texto.find(buscado, i + 1)
    return posiciones


def _plural(n, palabra):
    return f"{n} {palabra}" if n == 1 else f"{n} {palabra}s"
