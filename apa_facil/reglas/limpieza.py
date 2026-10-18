"""
Paso 3.5: limpieza de formato. Todo lo que se hace queda en el informe.

1. Párrafos vacíos usados como espaciado: se borran (en APA el espacio lo da
   el interlineado doble). Se conservan los que hacen algo: los que llevan un
   salto de página o de sección, una imagen, un marcador, una nota o un campo,
   y los que van justo antes o después de una tabla o una imagen (APA quiere
   una línea en blanco entre la tabla o la figura y el texto).
2. Tabuladores y espacios al principio de los párrafos de texto, usados como
   sangría: se quitan, y la regla de sangría pone después la de 1,27 cm.
3. Color, resaltado, fondo y subrayado del texto: se quitan del formato
   puesto a mano. Se conservan en los enlaces (APA admite el azul subrayado).

Las dos primeras cambian el contenido: cada cambio se anota en el informe
(informe.cambio_previsto) para que la comprobación de integridad lo acepte.

No se toca: la portada (paso 5.3: se sustituye o se mantiene tal cual), las
tablas (paso 6.1: sus colores van con sus fondos), los cuadros de texto, el
código de programación, las notas al pie (paso 6.6), los encabezados y pies,
ni los estilos (los de título los rehace el paso 4.1).
"""

from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.estructura import URL, _bloques, clasificar
from apa_facil.reglas.alineacion import _salto_antes
from apa_facil.reglas.fuente import es_codigo

# Lo que puede llevar un párrafo vacío de relleno, además del formato.
DE_RELLENO_EN_PARRAFO = {w("pPr"), w("r"), w("proofErr")}
DE_RELLENO_EN_FRAGMENTO = {w("rPr"), w("t"), w("tab"), w("lastRenderedPageBreak")}
MARCAS = ("color", "highlight", "shd", "u")
SIN_EFECTO = {None, "auto", "000000", "none", "clear", "FFFFFF"}
ESTILOS_DE_ENLACE = ("hyperlink", "hipervínculo", "hipervinculo")


def aplicar(doc, opciones, informe, estructura=None):
    estilos = Estilos(doc)
    clasificados = estructura if estructura is not None else clasificar(doc, estilos)
    _parrafos_vacios(doc, estilos, clasificados, informe)
    _sangria_a_mano(clasificados, informe)
    _color_y_subrayado(doc, estilos, clasificados, informe)


# --- 1. Párrafos vacíos -----------------------------------------------------------------

def _solo_formato(p):
    """¿No lleva el párrafo nada más que formato y espacio en blanco?"""
    for hijo in p:
        if hijo.tag not in DE_RELLENO_EN_PARRAFO:
            return False
        if hijo.tag == w("pPr") and hijo.find(w("sectPr")) is not None:
            return False   # salto de sección
        if hijo.tag == w("r"):
            for pieza in hijo:
                salto_de_linea = pieza.tag == w("br") and pieza.get(w("type")) in (None, "textWrapping")
                if pieza.tag not in DE_RELLENO_EN_FRAGMENTO and not salto_de_linea:
                    return False   # salto de página, imagen, nota, campo…
    return True


def _parrafos_vacios(doc, estilos, clasificados, informe):
    bloques = _bloques(doc)
    posicion = {b: i for i, b in enumerate(bloques)}

    def junto_a_tabla_o_imagen(p):
        i = posicion[p]
        vecinos = [bloques[j] for j in (i - 1, i + 1) if 0 <= j < len(bloques)]
        return any(v.tag == w("tbl") or _es_imagen(v) for v in vecinos)

    borrar = [c for c in clasificados
              if c.tipo == "vacio" and c.zona != "portada" and _solo_formato(c.elemento)
              and not _salto_antes(estilos, c.elemento) and not junto_a_tabla_o_imagen(c.elemento)]
    for c in borrar:
        informe.cambio_previsto(c.numero, "borrar")
        c.elemento.getparent().remove(c.elemento)
    # La estructura sigue describiendo el documento: sin los párrafos que ya no existen.
    clasificados[:] = [c for c in clasificados if c not in borrar]

    if borrar:
        n = len(borrar)
        informe.cambio(f"Limpieza: se {'ha' if n == 1 else 'han'} eliminado {_parrafos(n)} "
                       f"{'vacío' if n == 1 else 'vacíos'} que se usaba{'n' if n > 1 else ''} para dejar espacio. "
                       "En APA no se dejan líneas en blanco de más: el espacio lo da el interlineado doble.")


def _es_imagen(bloque):
    return bloque.tag == w("p") and not texto_parrafo(bloque).strip() and \
        next(bloque.iter(w("drawing"), w("pict")), None) is not None


# --- 2. Tabuladores y espacios usados como sangría --------------------------------------

def _quitar_del_principio(p):
    """Quita los tabuladores y espacios del principio del párrafo. Devuelve lo quitado."""
    quitado = []
    for pieza in p.iter(w("t"), w("tab"), w("br"), w("cr"), w("sym"), w("noBreakHyphen"), w("delText")):
        if next(pieza.iterancestors(w("p")), None) is not p:
            continue   # de un cuadro de texto que va dentro del párrafo
        if pieza.tag == w("tab"):
            if pieza.getparent().tag != w("r"):
                continue   # una tabulación del formato del párrafo, no del texto
            quitado.append("\t")
            pieza.getparent().remove(pieza)
            continue
        if pieza.tag != w("t"):
            break
        texto = pieza.text or ""
        limpio = texto.lstrip(" \t")
        quitado.append(texto[:len(texto) - len(limpio)])
        pieza.text = limpio
        if limpio:
            break
    return "".join(quitado)


def _sangria_a_mano(clasificados, informe):
    limpiados = []
    for c in clasificados:
        texto = texto_parrafo(c.elemento)
        if c.tipo == "cuerpo" and (texto.startswith("\t") or texto.startswith("  ")):
            quitado = _quitar_del_principio(c.elemento)
            informe.cambio_previsto(c.numero, "recortar", quitado)
            limpiados.append(c)
    if limpiados:
        n = len(limpiados)
        informe.cambio(f"Limpieza: se han quitado los tabuladores o espacios del principio de {_parrafos(n)} "
                       f"(«{_extracto(limpiados[0].texto)}»), que hacía{'n' if n > 1 else ''} de sangría. "
                       f"Ahora lleva{'n' if n > 1 else ''} la sangría de 1,27 cm.")


# --- 3. Color, resaltado y subrayado ----------------------------------------------------

def _es_enlace(r, estilos):
    if next(r.iterancestors(w("hyperlink")), None) is not None:
        return True
    estilo = r.find(f"{w('rPr')}/{w('rStyle')}")
    if estilo is not None and any(e in estilos.nombre(estilo.get(w("val"))).lower() for e in ESTILOS_DE_ENLACE):
        return True
    return bool(URL.search("".join(t.text or "" for t in r.iter(w("t")))))   # enlace escrito a mano


def _color_y_subrayado(doc, estilos, clasificados, informe):
    cuenta = {"color": [], "resaltado": [], "subrayado": []}
    for c in clasificados:
        if c.zona == "portada" or es_codigo(c.elemento, estilos):
            continue
        for r in c.elemento.iter(w("r")):
            rpr = r.find(w("rPr"))
            if rpr is None or next(r.iterancestors(w("p")), None) is not c.elemento or _es_enlace(r, estilos):
                continue
            texto = "".join(t.text or "" for t in r.iter(w("t"))).strip()
            for etiqueta in MARCAS:
                marca = rpr.find(w(etiqueta))
                if marca is None:
                    continue
                valor = marca.get(w("fill")) if etiqueta == "shd" else marca.get(w("val"))
                if texto and (valor not in SIN_EFECTO or marca.get(w("themeColor"))):
                    grupo = {"color": "color", "u": "subrayado"}.get(etiqueta, "resaltado")
                    if not cuenta[grupo] or cuenta[grupo][-1][0] is not r:
                        cuenta[grupo].append((r, texto))
                rpr.remove(marca)

    for grupo, explicacion in (("color", "APA usa texto negro"),
                               ("resaltado", "APA no usa resaltados ni fondos de color"),
                               ("subrayado", "APA no usa subrayado")):
        if cuenta[grupo]:
            n = len(cuenta[grupo])
            articulo = "el" if grupo != "resaltado" else "el resaltado o"
            que = "fondo de color" if grupo == "resaltado" else grupo
            informe.cambio(f"Limpieza: se ha quitado {articulo} {que} de {n} fragmento{'s' if n > 1 else ''} de "
                           f"texto («{_extracto(cuenta[grupo][0][1])}»). {explicacion}; los enlaces "
                           "conservan el suyo.")


# --- Mensajes ---------------------------------------------------------------------------

def _parrafos(n):
    return "1 párrafo" if n == 1 else f"{n} párrafos"


def _extracto(texto, largo=35):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo].rstrip() + "…"
