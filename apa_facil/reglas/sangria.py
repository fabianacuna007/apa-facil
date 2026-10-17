"""
Regla 1.5, versión definitiva (paso 3.4): sangría de 1,27 cm en la primera
línea de los párrafos de texto del cuerpo.

La sangría se pone en cada párrafo, no en el estilo Normal (que también usan
títulos y tablas). Si algún estilo traía sangría de primera línea, se le quita.

Qué lleva sangría lo decide el clasificador de párrafos (estructura.py): solo
los de tipo "cuerpo", salvo
  - el primer párrafo del resumen (APA lo pide sin sangría; la línea de
    palabras clave sí la lleva);
  - los centrados o alineados a la derecha (la regla 1.4 no los toca);
  - los que empiezan con un tabulador o espacios usados como sangría: se
    dejan sin ella y se avisa.
"""

from apa_facil import documento
from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import NOMBRES_DE_TIPO, clasificar

SANGRIA = 720                 # 1,27 cm en twips (exacto: media pulgada)


def _sangria(estilos, p, atributo):
    try:
        return int(estilos.parrafo(p, "ind", atributo) or 0)
    except ValueError:
        return 0


def motivos_sin_sangria(doc, estilos=None, estructura=None):
    """
    Para cada párrafo del cuerpo, por qué no lleva sangría (o None si la lleva).
    Devuelve una lista de (párrafo, motivo) en el orden del documento.
    `estructura` es la clasificación de los párrafos; si no se da, se clasifica
    el documento tal como está.
    """
    estilos = estilos or Estilos(doc)
    clasificados = estructura if estructura is not None else clasificar(doc, estilos)
    primero_del_resumen = next((c.elemento for c in clasificados if c.zona == "resumen" and c.tipo == "cuerpo"),
                               None)
    resultado = []
    for c in clasificados:
        p, texto = c.elemento, texto_parrafo(c.elemento)
        if c.tipo != "cuerpo":
            motivo = NOMBRES_DE_TIPO[c.tipo]
        elif p is primero_del_resumen:
            motivo = "primer párrafo del resumen"
        elif (estilos.parrafo(p, "jc", "val") or "left") in ("center", "right", "end"):
            motivo = "centrado o a la derecha"
        elif texto.startswith("\t") or texto.startswith("  "):
            motivo = "tabulador"
        else:
            motivo = None
        resultado.append((p, motivo))
    return resultado


def _poner_sangria(ppr):
    ind = ppr.get_or_add_ind()
    for atributo in ("hanging", "hangingChars", "firstLineChars"):
        ind.attrib.pop(w(atributo), None)
    ind.set(w("firstLine"), str(SANGRIA))


def _quitar_primera_linea(ppr):
    ind = ppr.find(w("ind")) if ppr is not None else None
    if ind is None:
        return
    for atributo in ("firstLine", "firstLineChars"):
        ind.attrib.pop(w(atributo), None)


def aplicar(doc, opciones, informe, estructura=None):
    estilos = Estilos(doc)
    clasificados = motivos_sin_sangria(doc, estilos, estructura)
    con_sangria = [p for p, motivo in clasificados if motivo is None]
    ya_estaban = sum(1 for p in con_sangria if _sangria(estilos, p, "firstLine") == SANGRIA
                     and _sangria(estilos, p, "hanging") == 0)
    con_tabulador = [p for p, motivo in clasificados if motivo == "tabulador"]

    # Ninguna sangría de primera línea "heredada": solo la llevan los párrafos de texto.
    _quitar_primera_linea(documento.parrafo_por_defecto(doc))
    for estilo in doc.styles.element.iter(w("style")):
        _quitar_primera_linea(estilo.find(w("pPr")))
    for p in con_sangria:
        _poner_sangria(p.get_or_add_pPr())

    nuevos = len(con_sangria) - ya_estaban
    if nuevos:
        informe.cambio(f"Sangría: {nuevos} párrafo{'s' if nuevos > 1 else ''} de texto "
                       f"lleva{'n' if nuevos > 1 else ''} ahora 1,27 cm en la primera línea.")
    if con_tabulador:
        varios = len(con_tabulador) > 1
        extracto = " ".join(texto_parrafo(con_tabulador[0]).split())[:35]
        informe.aviso(f"{len(con_tabulador)} párrafo{'s' if varios else ''} empieza{'n' if varios else ''} con "
                      f"un tabulador o espacios usados como sangría («{extracto}…»). Se ha{'n' if varios else ''} "
                      "dejado así: quita el tabulador o los espacios del principio y la herramienta pondrá la "
                      "sangría correcta.")
