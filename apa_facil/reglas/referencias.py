"""
Paso 6.4: la lista de referencias.

Según la web oficial de APA (Reference list setup):
- empieza en una página nueva, después del texto;
- el rótulo «Referencias» («Referencia» si solo hay una), arriba, centrado y en
  negrita: APA no admite «Bibliografía», «Obras citadas» ni «Lista de referencias»;
- cada referencia, un párrafo alineado a la izquierda, con sangría francesa de
  1,27 cm (la primera línea sin sangría y las demás sangradas);
- a doble espacio y sin espacio extra entre referencias (lo hace el paso 1.3; la
  alineación, el 1.4);
- los DOI y las URL pueden ir en azul y subrayados o como texto normal: se dejan
  como estén.

El rótulo lleva el estilo «Título 1» de Word (centrado y en negrita, como los de
nivel 1), así que sale en el panel de navegación. Si antes de él hay un párrafo
vacío que solo lleva un salto de página, se borra y el rótulo empieza página con
«Salto de página anterior»: así no queda una línea en blanco arriba.

Cambiar el rótulo cambia el texto: se hace por defecto, se anota en el informe y
lo comprueba la integridad (decisión del propietario); con --rotulo-tal-cual se
deja como está y se avisa. Si hay varias listas (una bibliografía y una
webgrafía), no se cambia ningún rótulo: unirlas es del paso 6.5.
"""

from types import SimpleNamespace

from apa_facil.documento import (ROTULOS_APA_DE_REFERENCIAS, Estilos, texto_parrafo, w)
from apa_facil.estructura import clasificar
from apa_facil.reglas import titulos
from apa_facil.reglas.fuente import _runs_propios
from apa_facil.reglas.leyendas import _lista, _n, _plural, _se_puede_cambiar, _xml
from apa_facil.reglas.portada import _termina_en_salto

SANGRIA = 720   # 1,27 cm en veinteavos de punto
# Lo que impide borrar el párrafo vacío de antes del rótulo (además de su texto).
NO_SE_BORRA_SI_LLEVA = ("sectPr", "drawing", "pict", "bookmarkStart", "fldChar", "fldSimple", "footnoteReference",
                        "endnoteReference", "commentReference", "txbxContent")


def listas(clasificados):
    """Cada lista de referencias: (rótulo, [referencias]), en orden. Las referencias son las de su zona."""
    resultado = []
    for c in clasificados:
        if c.elemento.getparent() is None:
            continue
        if c.tipo == "titulo_referencias":
            resultado.append((c, []))
        elif c.tipo == "referencia" and c.zona == "referencias" and resultado:
            resultado[-1][1].append(c)
    return resultado


def rotulo_apa(referencias):
    return ROTULOS_APA_DE_REFERENCIAS[1] if len(referencias) == 1 else ROTULOS_APA_DE_REFERENCIAS[0]


def _cambiar_texto(p, nuevo):
    """El texto del párrafo pasa a ser `nuevo`, en su primer fragmento (con su formato). Devuelve si se ha podido."""
    texto = texto_parrafo(p)
    if not _se_puede_cambiar(p, len(texto)):
        return False
    piezas = [el for r in _runs_propios(p) for el in r if el.tag in (w("t"), w("tab"))]
    primera = next((el for el in piezas if el.tag == w("t") and el.text), None)
    if primera is None:
        return False
    for el in piezas:
        if el is primera:
            el.text = nuevo
        elif el.tag == w("t"):
            el.text = ""
        else:
            el.getparent().remove(el)
    return texto_parrafo(p) == nuevo


def _estilo_de_nivel_1(doc):
    """El estilo «Título 1» de Word, con el formato APA (creándolo si no existe)."""
    estilos = titulos.estilos_de_titulo(doc)[1]
    estilo = titulos._quitar_duplicados(estilos) if estilos else titulos._crear(doc, 1)
    titulos._dar_formato(doc, estilo, 1)
    return estilo.get(w("styleId"))


def _anterior(p):
    return next((s for s in p.itersiblings(preceding=True) if s.tag in (w("p"), w("tbl"), w("sdt"))), None)


def _solo_salto(p):
    """¿Es un párrafo sin texto que solo lleva un salto de página (y nada que haya que conservar)?"""
    return (p is not None and p.tag == w("p") and not texto_parrafo(p).strip() and _termina_en_salto(p)
            and not any(next(p.iter(w(e)), None) is not None for e in NO_SE_BORRA_SI_LLEVA))


def _pagina_nueva(doc, informe, rotulo, clasificados):
    """
    La lista empieza en una página nueva. Devuelve "borrado" si se ha quitado el
    párrafo vacío del salto, "puesto" si se ha puesto el salto y None si ya la empezaba.
    """
    p = rotulo.elemento
    if titulos._salto_de_pagina_antes(Estilos(doc), p):
        return None
    anterior = _anterior(p)
    hecho = "puesto"
    if _solo_salto(anterior):
        vacio = next((c for c in clasificados if c.elemento is anterior), None)
        if vacio is None:
            return None
        informe.cambio_previsto(vacio.numero, "borrar")
        anterior.getparent().remove(anterior)
        clasificados[:] = [c for c in clasificados if c is not vacio]
        hecho = "borrado"
    elif anterior is not None and anterior.tag == w("p") and _termina_en_salto(anterior):
        return None   # el párrafo de antes, con texto, ya acaba en un salto de página
    salto = p.get_or_add_pPr().get_or_add_pageBreakBefore()
    salto.attrib.pop(w("val"), None)
    return hecho


def _sangria_francesa(p):
    """1,27 cm de sangría francesa (la primera línea, sin sangría). Devuelve si cambia algo."""
    ppr = p.get_or_add_pPr()
    antes = _xml(ppr)
    ind = ppr.get_or_add_ind()
    for atributo in ("firstLine", "firstLineChars", "hangingChars", "leftChars", "startChars", "start", "right", "end",
                     "rightChars", "endChars"):
        ind.attrib.pop(w(atributo), None)
    ind.set(w("left"), str(SANGRIA))
    ind.set(w("hanging"), str(SANGRIA))
    return _xml(ppr) != antes


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    todas = listas(clasificados)
    if not todas:
        return
    tal_cual = opciones.get("rotulo_tal_cual")
    id_estilo = _estilo_de_nivel_1(doc)
    renombrados, sin_cambiar, con_estilo, paginas, francesas = [], [], [], [], 0

    for rotulo, referencias in todas:
        p = rotulo.elemento
        antes = texto_parrafo(p)
        nuevo = rotulo_apa(referencias)
        if antes != nuevo:
            if tal_cual or len(todas) > 1:
                sin_cambiar.append(antes)
            elif _cambiar_texto(p, nuevo):
                informe.cambio_previsto(rotulo.numero, "rotulo_de_referencias", nuevo=nuevo)
                renombrados.append((antes, nuevo))
            else:
                sin_cambiar.append(antes)
        pstyle = p.get_or_add_pPr().get_or_add_pStyle()
        if pstyle.get(w("val")) != id_estilo:
            pstyle.set(w("val"), id_estilo)
            con_estilo.append(texto_parrafo(p))
        titulos._formato_a_mano(Estilos(doc), SimpleNamespace(elemento=p, nivel=1))
        if rotulo is todas[0][0]:   # la lista (la primera, si hay varias) empieza en una página nueva
            hecho = _pagina_nueva(doc, informe, rotulo, clasificados)
            if hecho:
                paginas.append(hecho)
        francesas += sum(_sangria_francesa(c.elemento) for c in referencias)

    if renombrados:
        antes, nuevo = renombrados[0]
        informe.cambio(f"Referencias: el rótulo «{antes}» pasa a ser «{nuevo}», como pide APA, que no admite "
                       "«Bibliografía», «Obras citadas» ni «Lista de referencias». Si tu universidad pide otro, usa la "
                       "opción --rotulo-tal-cual.")
    if con_estilo:
        informe.cambio(f"Referencias: el rótulo {_lista(con_estilo)} lleva ahora el estilo «Título 1» de Word: va "
                       "centrado y en negrita, como pide APA, y aparece en el panel de navegación.")
    if paginas:
        informe.cambio("Referencias: la lista empieza ahora en una página nueva, como pide APA"
                       + (" (se ha quitado el párrafo vacío que hacía el salto, para que no quede una línea en blanco "
                          "arriba)." if "borrado" in paginas else "."))
    if francesas:
        informe.cambio(f"Referencias: {_plural(francesas, 'referencia')} lleva{_n(francesas)} ahora sangría francesa "
                       "de 1,27 cm (la primera línea sin sangría y las demás sangradas), como pide APA.")
    if len(todas) > 1:
        informe.aviso(f"Hay {len(todas)} listas de referencias ({_lista([texto_parrafo(r.elemento) for r, _ in todas])}). "
                      "APA usa una sola, «Referencias», con todas las obras en orden alfabético: se han dejado los "
                      "rótulos como estaban. Únelas tú.")
    elif sin_cambiar and tal_cual:
        informe.aviso(f"El rótulo {_lista(sin_cambiar)} se ha dejado como estaba porque has usado la opción "
                      f"--rotulo-tal-cual: APA lo llama «{rotulo_apa(todas[0][1])}».")
    elif sin_cambiar:
        informe.aviso(f"El rótulo {_lista(sin_cambiar)} lleva un campo de Word o cambios marcados y no se ha podido "
                      f"cambiar: APA lo llama «{rotulo_apa(todas[0][1])}». Cámbialo tú.")
    for rotulo, referencias in todas:
        if not referencias:
            informe.aviso(f"Debajo del rótulo «{texto_parrafo(rotulo.elemento)}» no se ha encontrado ninguna referencia. "
                          "Revisa la lista.")
