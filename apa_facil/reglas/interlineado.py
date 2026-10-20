"""
Regla 1.3: interlineado doble, sin espacio antes ni después del párrafo.

- Cuerpo del documento (texto, títulos, portada, listas, citas): doble espacio
  y 0 pt antes y después.
- Tablas: se mantiene el interlineado de cada párrafo si es sencillo, 1,5 o
  doble (APA admite los tres); si es otro, se pone sencillo. Sin espacio antes
  ni después.
- No se tocan las notas al pie (APA pide el formato por defecto), los
  encabezados y pies (paso 1.6) ni los cuadros de texto (forman parte de
  figuras). Como el estilo Normal pasa a doble espacio y ellos lo heredarían,
  su espaciado actual se deja fijado en cada párrafo y en sus estilos.
"""

from collections import Counter

from docx.oxml import OxmlElement

from apa_facil import documento
from apa_facil.documento import Estilos, poner_espaciado, w

SENCILLO, UNO_Y_MEDIO, DOBLE = 240, 360, 480
ADMITIDOS_EN_TABLAS = {SENCILLO, UNO_Y_MEDIO, DOBLE}
TWIPS_POR_PUNTO = 20
SANGRIA_POR_NIVEL = 220   # sangría de cada nivel del índice en Word (0,39 cm)

ESTILOS_QUE_NO_SE_TOCAN = {"footnote text", "endnote text", "header", "footer",
                           "annotation text", "balloon text"}
PARTES_QUE_NO_SE_TOCAN = ("/word/header", "/word/footer", "/word/footnotes.xml", "/word/endnotes.xml")


def _entero(valor, por_defecto=0):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return por_defecto


def _espaciado(valores):
    """(line, lineRule, before, after) leídos de Word → números, con los valores por defecto de Word."""
    linea, regla, antes, despues = valores
    return _entero(linea, SENCILLO), regla or "auto", _entero(antes), _entero(despues)


def _efectivo(estilos, p):
    return _espaciado([estilos.parrafo(p, "spacing", a) for a in ("line", "lineRule", "before", "after")])


def _efectivo_de_estilo(estilos, sid):
    """El espaciado que da un estilo por sí solo (con lo que hereda y los valores por defecto)."""
    capas = [e.find(w("pPr")) for e in estilos.cadena(sid)] + [estilos.ppr_defecto]
    valores = []
    for atributo in ("line", "lineRule", "before", "after"):
        valor = None
        for capa in capas:
            espaciado = capa.find(w("spacing")) if capa is not None else None
            if espaciado is not None and espaciado.get(w(atributo)) is not None:
                valor = espaciado.get(w(atributo))
                break
        valores.append(valor)
    return _espaciado(valores)


def _dentro_de(p, *etiquetas):
    return next(p.iterancestors(*etiquetas), None) is not None


def aplicar(doc, opciones, informe, estructura=None):
    estilos = Estilos(doc)
    cuerpo = doc.element.body
    intactos = getattr(estructura, "intactos", frozenset())   # la portada que se mantiene

    # 1. Antes de cambiar nada: cómo está ahora el cuerpo (para el informe) y las tablas.
    antes_cuerpo = [_efectivo(estilos, p) for p in cuerpo.iter(w("p"))
                    if not _dentro_de(p, w("tc"), w("txbxContent"))]
    antes_tablas = [_efectivo(estilos, p) for p in cuerpo.iter(w("p")) if _dentro_de(p, w("tc"))]

    # 2. Fijar lo que no debe heredar el doble espacio: tablas, cuadros de texto,
    #    encabezados, pies y notas. Se hace antes de cambiar los estilos.
    for p in cuerpo.iter(w("p")):
        if p in intactos:
            continue
        if _dentro_de(p, w("txbxContent")):
            poner_espaciado(p.get_or_add_pPr(), *_efectivo(estilos, p))
        elif _dentro_de(p, w("tc")):
            linea, regla, _, _ = _efectivo(estilos, p)
            conservado = linea if regla == "auto" and linea in ADMITIDOS_EN_TABLAS else SENCILLO
            poner_espaciado(p.get_or_add_pPr(), conservado)

    def fijar_parte(nombre_parte, raiz):
        if nombre_parte.startswith(PARTES_QUE_NO_SE_TOCAN):
            for p in raiz.iter(w("p")):
                poner_espaciado(p.get_or_add_pPr(), *_efectivo(estilos, p))
    documento.editar_partes_de_texto(doc, fijar_parte)

    for estilo in doc.styles.element.iter(w("style")):
        nombre = estilo.find(w("name"))
        if nombre is not None and nombre.get(w("val"), "").lower() in ESTILOS_QUE_NO_SE_TOCAN:
            poner_espaciado(estilo.get_or_add_pPr(), *_efectivo_de_estilo(estilos, estilo.get(w("styleId"))))

    # 3. Doble espacio en los valores por defecto, en los estilos de párrafo y en
    #    los párrafos del cuerpo que tienen espaciado propio.
    poner_espaciado(documento.parrafo_por_defecto(doc), DOBLE)
    for estilo in doc.styles.element.iter(w("style")):
        nombre = estilo.find(w("name"))
        if (estilo.get(w("type")) == "paragraph" and estilo.find(f"{w('pPr')}/{w('spacing')}") is not None
                and (nombre is None or nombre.get(w("val"), "").lower() not in ESTILOS_QUE_NO_SE_TOCAN)):
            poner_espaciado(estilo.find(w("pPr")), DOBLE)
    for p in cuerpo.iter(w("p")):
        if _dentro_de(p, w("tc"), w("txbxContent")) or p in intactos:
            continue
        ppr = p.find(w("pPr"))
        if ppr is not None and ppr.find(w("spacing")) is not None:
            poner_espaciado(ppr, DOBLE)
    if _tiene_indice(cuerpo):
        _estilos_de_indice(doc)

    _informar(antes_cuerpo, antes_tablas, informe)


def _tiene_indice(cuerpo):
    codigos = [t.text or "" for t in cuerpo.iter(w("instrText"))]
    codigos += [f.get(w("instr"), "") for f in cuerpo.iter(w("fldSimple"))]
    return any(c.strip().upper().startswith("TOC") for c in codigos)


def _estilos_de_indice(doc):
    """
    Word rehace el índice con sus estilos "toc 1" a "toc 9" (en español, "TDC").
    Si el documento no los define, Word usa los suyos, con 5 pt de espacio
    después. Se definen (o se corrigen) a doble espacio, con la sangría
    habitual de cada nivel.
    """
    raiz = doc.styles.element
    por_nombre = {(e.find(w("name")).get(w("val")) or "").lower(): e
                  for e in raiz.iter(w("style")) if e.find(w("name")) is not None}
    ids = {e.get(w("styleId")) for e in raiz.iter(w("style"))}
    for nivel in range(1, 10):
        estilo = por_nombre.get(f"toc {nivel}")
        if estilo is None:
            estilo = OxmlElement("w:style")
            estilo.set(w("type"), "paragraph")
            estilo.set(w("styleId"), next(i for i in (f"TOC{nivel}", f"TDC{nivel}", f"APA_TOC{nivel}")
                                          if i not in ids))
            for etiqueta, valor in (("w:name", f"toc {nivel}"), ("w:uiPriority", "39")):
                hijo = OxmlElement(etiqueta)
                hijo.set(w("val"), valor)
                estilo.append(hijo)
            estilo.append(OxmlElement("w:unhideWhenUsed"))
            raiz.append(estilo)
            estilo.get_or_add_pPr().get_or_add_ind().set(w("left"), str(SANGRIA_POR_NIVEL * (nivel - 1)))
        poner_espaciado(estilo.get_or_add_pPr(), DOBLE)


# --- Informe ------------------------------------------------------------------

def _describir(linea, regla):
    if regla == "exact":
        return f"exacto de {_pt(linea / TWIPS_POR_PUNTO)} pt"
    if regla == "atLeast":
        return f"mínimo de {_pt(linea / TWIPS_POR_PUNTO)} pt"
    return {SENCILLO: "sencillo", UNO_Y_MEDIO: "1,5", DOBLE: "doble"}.get(linea, _pt(linea / SENCILLO))


def _pt(numero):
    return f"{round(numero, 2):g}".replace(".", ",")


def _lista(elementos):
    elementos = list(elementos)
    return elementos[0] if len(elementos) == 1 else ", ".join(elementos[:-1]) + " y " + elementos[-1]


def _informar(antes_cuerpo, antes_tablas, informe):
    if antes_cuerpo and any(e != (DOBLE, "auto", 0, 0) for e in antes_cuerpo):
        interlineados = Counter(_describir(linea, regla) for linea, regla, _, _ in antes_cuerpo)
        mensaje = f"antes: {_lista([d for d, _ in interlineados.most_common()])}"
        separacion = max(max(antes, despues) for _, _, antes, despues in antes_cuerpo)
        if separacion:
            mensaje += f", con hasta {_pt(separacion / TWIPS_POR_PUNTO)} pt de separación entre párrafos"
        informe.cambio("Interlineado: todo el texto pasa a doble espacio, sin espacio extra antes ni "
                       f"después de los párrafos ({mensaje}).")

    if not antes_tablas:
        return
    conservados = Counter(_describir(l, r) for l, r, _, _ in antes_tablas
                          if r == "auto" and l in ADMITIDOS_EN_TABLAS)
    cambiados = Counter(_describir(l, r) for l, r, _, _ in antes_tablas
                        if not (r == "auto" and l in ADMITIDOS_EN_TABLAS))
    con_espacio = any(antes or despues for _, _, antes, despues in antes_tablas)
    partes = []
    if conservados:
        partes.append(f"se mantiene el interlineado {_lista(list(conservados))}, que APA admite")
    if cambiados:
        partes.append(f"el interlineado {_lista(list(cambiados))} pasa a sencillo "
                      "(APA admite sencillo, 1,5 o doble)")
    if con_espacio:
        partes.append("se quita el espacio extra entre párrafos")
    if cambiados or con_espacio:
        informe.cambio("Tablas: " + "; ".join(partes) + ".")
