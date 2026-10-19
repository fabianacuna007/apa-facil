"""
Paso 6.3: las citas en bloque.

Según la web oficial de APA (Quotations), una cita de 40 palabras o más va en
bloque:
- en un párrafo propio, con todo el bloque sangrado 1,27 cm desde el margen
  izquierdo (del derecho no dice nada: no lleva sangría) y sin sangría de
  primera línea;
- si tiene varios párrafos, la primera línea del segundo y los siguientes lleva
  1,27 cm más;
- a doble espacio y sin espacio extra antes ni después (lo hace el paso 1.3);
- sin comillas, y con la fuente entre paréntesis después del punto final de la
  cita y sin punto detrás.

Qué es una cita en bloque lo decide el clasificador: un párrafo con sangría a
la izquierda (y sin sangría francesa). Varios seguidos son una sola cita, si
entre todos tienen 40 palabras o más. La cursiva que va en todo el bloque es
decorativa y se quita; la de una parte (el título de un libro) se respeta.

Lo que cambiaría el texto solo se avisa (ROADMAP y decisión del propietario): las
comillas, el punto detrás del paréntesis, el párrafo con sangría y menos de 40
palabras (si es una cita, va dentro del texto y entre comillas) y la cita de 40
palabras o más escrita entre comillas dentro de un párrafo.
"""

import re

from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import PALABRAS_DE_CITA_EN_BLOQUE, clasificar
from apa_facil.reglas import titulos
from apa_facil.reglas.leyendas import _lista, _n, _plural, _xml
from apa_facil.reglas.resumen import _poner

SANGRIA = 720   # 1,27 cm en veinteavos de punto
ABRE = "\"«“'‘"
CIERRA = "\"»”'’"
# Una cita entre comillas dentro de un párrafo: lo que va entre «…», "…" o “…”.
ENTRE_COMILLAS = re.compile(r"[«\"“]([^«»\"“”]+)[»\"”]")
PUNTO_TRAS_LA_FUENTE = re.compile(r"\)\s*\.\s*$")


def bloques_de_cita(clasificados):
    """Las citas en bloque: listas de párrafos `cita_bloque` seguidos (de 40 palabras o más entre todos)."""
    bloques, actual = [], []
    for c in clasificados:
        if c.elemento.getparent() is None:
            continue
        if c.tipo == "cita_bloque":
            actual.append(c)
            continue
        if actual:
            bloques.append(actual)
        actual = []
    if actual:
        bloques.append(actual)
    return bloques


def palabras(bloque):
    return sum(len(texto_parrafo(c.elemento).split()) for c in bloque)


def _sangria(p, primera_linea):
    """1,27 cm a la izquierda, nada a la derecha y `primera_linea` en la primera línea. Devuelve si cambia algo."""
    ppr = p.get_or_add_pPr()
    antes = _xml(ppr)
    ind = ppr.get_or_add_ind()
    for atributo in ("right", "end", "rightChars", "endChars", "hanging", "hangingChars", "firstLineChars",
                     "leftChars", "startChars", "start"):
        ind.attrib.pop(w(atributo), None)
    ind.set(w("left"), str(SANGRIA))
    ind.set(w("firstLine"), str(primera_linea))
    return _xml(ppr) != antes


def _cursiva_en_todo(estilos, bloque):
    fragmentos = [(r, c.elemento) for c in bloque for r in titulos.fragmentos_con_letras(c.elemento)]
    return bool(fragmentos) and all(titulos.cursiva_efectiva(estilos, r, p) for r, p in fragmentos)


def entre_comillas(bloque):
    """¿Va la cita entre comillas? (empieza con unas y las cierra antes de la fuente o al final)"""
    texto = " ".join(texto_parrafo(c.elemento).strip() for c in bloque)
    return bool(texto) and texto[0] in ABRE and any(ch in CIERRA for ch in texto[1:])


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    estilos = Estilos(doc)
    con_sangria, sin_cursiva, con_comillas, con_punto, cortas = [], [], [], [], []

    for bloque in bloques_de_cita(clasificados):
        extracto = texto_parrafo(bloque[0].elemento)
        if palabras(bloque) < PALABRAS_DE_CITA_EN_BLOQUE:
            cortas.append(extracto)
            continue
        cambiado = False
        for n, c in enumerate(bloque):
            cambiado |= _sangria(c.elemento, 0 if n == 0 else SANGRIA)
        if cambiado:
            con_sangria.append(extracto)
        if _cursiva_en_todo(estilos, bloque):
            for c in bloque:
                for r in titulos.fragmentos_con_letras(c.elemento):
                    _poner(r, "i", False)
            sin_cursiva.append(extracto)
        if entre_comillas(bloque):
            con_comillas.append(extracto)
        if PUNTO_TRAS_LA_FUENTE.search(texto_parrafo(bloque[-1].elemento).strip()):
            con_punto.append(extracto)

    largas = [cita for c in clasificados if c.tipo == "cuerpo" and c.elemento.getparent() is not None
              for cita in ENTRE_COMILLAS.findall(texto_parrafo(c.elemento))
              if len(cita.split()) >= PALABRAS_DE_CITA_EN_BLOQUE]

    if con_sangria:
        n = _n(len(con_sangria))
        informe.cambio(f"Citas en bloque: {_plural(len(con_sangria), 'cita')} ({_lista(con_sangria)}) lleva{n} ahora "
                       "todo el bloque sangrado 1,27 cm desde el margen izquierdo, sin sangría a la derecha ni en la "
                       "primera línea, como pide APA.")
    if sin_cursiva:
        n = _n(len(sin_cursiva))
        informe.cambio(f"Citas en bloque: se ha quitado la cursiva de {_plural(len(sin_cursiva), 'cita')} "
                       f"({_lista(sin_cursiva)}), que la llevaba{n} en todo el texto: APA no la usa en las citas.")
    if con_comillas:
        n = _n(len(con_comillas))
        informe.aviso(f"{_plural(len(con_comillas), 'cita')} en bloque va{n} entre comillas ({_lista(con_comillas)}): "
                      "APA no las pone en una cita en bloque, porque la sangría ya indica que es una cita. Quita tú "
                      "las de apertura y cierre (no las que haya dentro de la cita).")
    if con_punto:
        informe.aviso(f"{_plural(len(con_punto), 'cita')} en bloque acaba{_n(len(con_punto))} con un punto después "
                      f"del paréntesis de la fuente ({_lista(con_punto)}): APA pone el punto al final de la cita, "
                      "antes del paréntesis, y nada detrás. Cámbialo tú.")
    if cortas:
        n = _n(len(cortas))
        informe.aviso(f"{_plural(len(cortas), 'párrafo')} con sangría a la izquierda tiene{n} menos de "
                      f"{PALABRAS_DE_CITA_EN_BLOQUE} palabras ({_lista(cortas)}). Si es una cita, APA la pone dentro "
                      "del texto y entre comillas: la cita en bloque es para las de 40 palabras o más. Se ha dejado "
                      "como estaba.")
    if largas:
        n = _n(len(largas))
        informe.aviso(f"{_plural(len(largas), 'cita')} entre comillas dentro de un párrafo tiene{n} "
                      f"{PALABRAS_DE_CITA_EN_BLOQUE} palabras o más ({_lista(largas)}): APA la pone en un párrafo "
                      "aparte, sin comillas y con todo el bloque sangrado 1,27 cm. Hazlo tú (la herramienta le dará "
                      "el formato).")
