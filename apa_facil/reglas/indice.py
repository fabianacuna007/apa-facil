"""
Paso 5.4: el índice.

La web oficial de APA (Order of pages) no incluye el índice en un trabajo de
estudiante, y para las tesis y los TFG remite a las normas de cada universidad.
Muchas universidades españolas lo exigen, así que por defecto se mantiene y se
avisa (decisión del propietario). Con --quitar-indice se quita:

- los párrafos que el clasificador marca como índice: el control de contenido
  «Tabla de contenido» de Word entero, o el rótulo («Índice») y las líneas de un
  índice escrito a mano. Se conservan los párrafos que llevan un salto de sección
  (si no, se perdería una sección);
- el párrafo vacío con un salto de página que va justo antes, si después ya
  empieza una página: si no, quedaría una página en blanco (decisión del
  propietario);
- «actualizar campos al abrir», si ya no queda ningún índice: solo servía para
  rehacerlo, y Word deja de preguntar al abrir (decisión del propietario).

No se quita, y se avisa, si algo no parece de un índice (una línea sin número de
página, una tabla, una imagen, una nota…): el índice escrito a mano se reconoce
por su rótulo, y el clasificador podría haber metido en él texto del trabajo.
"""

from apa_facil.documento import es_parte_de_un_indice, texto_parrafo, w
from apa_facil.estructura import clasificar

NO_SE_QUITA_SI_LLEVA = ("drawing", "pict", "footnoteReference", "endnoteReference", "commentReference",
                        "txbxContent")
BLOQUES = (w("p"), w("tbl"), w("sdt"))


def _extracto(texto, largo=40):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _salto_de_seccion(p):
    return p.find(f"{w('pPr')}/{w('sectPr')}") is not None


def _por_que_no_se_puede_quitar(parrafos):
    """None si el índice se puede quitar; si no, el motivo, para el aviso."""
    for c in parrafos:
        if next(c.elemento.iterancestors(w("tc")), None) is not None:
            return "está dentro de una tabla"
        if any(next(c.elemento.iter(w(e)), None) is not None for e in NO_SE_QUITA_SI_LLEVA):
            return "lleva una imagen, una nota, un comentario o un cuadro de texto"
        if not es_parte_de_un_indice(texto_parrafo(c.elemento)):
            return f"hay líneas que no parecen de un índice («{_extracto(texto_parrafo(c.elemento))}»)"
    # Los campos (el índice automático es un campo TOC) tienen que quitarse enteros.
    principios = sum(1 for c in parrafos for f in c.elemento.iter(w("fldChar")) if f.get(w("fldCharType")) == "begin")
    finales = sum(1 for c in parrafos for f in c.elemento.iter(w("fldChar")) if f.get(w("fldCharType")) == "end")
    if principios != finales:
        return "su campo de Word sigue fuera de las líneas del índice"
    return None


def _bloque(p):
    """El bloque del cuerpo en que está el párrafo: él mismo o el control de contenido que lo contiene."""
    while p.getparent() is not None and p.getparent().tag != w("body"):
        p = p.getparent()
    return p


def _empieza_pagina(bloque):
    """¿Empieza página lo que hay en este bloque? (salto de página anterior o un párrafo vacío con salto de sección)"""
    if bloque is None or bloque.tag != w("p"):
        return False
    salto = bloque.find(f"{w('pPr')}/{w('pageBreakBefore')}")
    if salto is not None and salto.get(w("val")) not in ("0", "false", "off"):
        return True
    seccion = bloque.find(f"{w('pPr')}/{w('sectPr')}")
    tipo = seccion.find(w("type")) if seccion is not None else None
    sin_texto = not texto_parrafo(bloque).strip()
    return seccion is not None and sin_texto and (tipo is None or tipo.get(w("val")) != "continuous")


def _vacio_con_salto(bloque):
    return (bloque is not None and bloque.tag == w("p") and not texto_parrafo(bloque).strip()
            and not _salto_de_seccion(bloque) and any(b.get(w("type")) == "page" for b in bloque.iter(w("br"))))


def _queda_algun_indice(doc):
    cuerpo = doc.element.body
    return (any((t.text or "").split()[:1] == ["TOC"] for t in cuerpo.iter(w("instrText")))
            or any((f.get(w("instr")) or "").split()[:1] == ["TOC"] for f in cuerpo.iter(w("fldSimple"))))


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    del_indice = [c for c in clasificados if c.zona == "indice" and c.elemento.getparent() is not None]
    if not del_indice:
        return
    if not opciones.get("quitar_indice"):
        informe.aviso("Tu trabajo tiene un índice. APA no lo pide en los trabajos de estudiante, pero muchas "
                      "universidades lo exigen, así que se ha dejado. Si la tuya no lo pide, usa la opción "
                      "--quitar-indice.")
        return

    quitar = [c for c in del_indice if not _salto_de_seccion(c.elemento)]
    motivo = _por_que_no_se_puede_quitar(quitar) if quitar else "solo quedan sus saltos de sección"
    if motivo:
        informe.aviso(f"No se ha quitado el índice porque {motivo}. Si tu universidad no lo pide, quítalo tú en Word.")
        return

    numero_de = {c.elemento: c.numero for c in clasificados}
    bloques = list(dict.fromkeys(_bloque(c.elemento) for c in quitar))
    elementos = {c.elemento for c in quitar}
    # Un control de contenido se quita entero solo si todo lo que lleva es del índice.
    a_quitar = [b for b in bloques if b.tag == w("p") or
                (all(p in elementos for p in b.iter(w("p"))) and next(b.iter(w("tbl")), None) is None)]
    if len(a_quitar) != len(bloques):
        a_quitar = [c.elemento for c in quitar]
    anterior = next((s for s in a_quitar[0].itersiblings(preceding=True) if s.tag in BLOQUES), None)
    siguiente = next((s for s in a_quitar[-1].itersiblings() if s.tag in BLOQUES), None)

    quitados = [c for c in quitar]
    for bloque in a_quitar:
        bloque.getparent().remove(bloque)
    # El salto de página de antes del índice, si lo que queda detrás ya empieza página: si no, una página en blanco.
    hueco = _vacio_con_salto(anterior) and _empieza_pagina(siguiente) and anterior in numero_de
    if hueco:
        anterior.getparent().remove(anterior)
        informe.cambio_previsto(numero_de[anterior], "quitar_indice")
    for c in quitados:
        informe.cambio_previsto(c.numero, "quitar_indice")

    con_texto = [_extracto(c.texto) for c in quitados if c.texto.strip()]
    detalle = (f": {len(con_texto)} línea{'s' if len(con_texto) > 1 else ''} («{con_texto[0]}»"
               f"{'…' if len(con_texto) > 1 else ''})" if con_texto else " (el índice automático de Word)")
    informe.cambio(f"Índice: se ha quitado el índice{detalle}, como has pedido: APA no lo usa en los trabajos de "
                   "estudiante." + (" También se ha quitado el salto de página que había antes, para que no quede "
                                    "una página en blanco." if hueco else ""))
    actualizar = doc.settings.element.find(w("updateFields"))
    if actualizar is not None and not _queda_algun_indice(doc):
        doc.settings.element.remove(actualizar)
        informe.cambio("Índice: Word ya no preguntará al abrir el trabajo si quieres actualizar los campos: solo "
                       "hacía falta para rehacer el índice.")
