"""
Regla 1.5 (versión simple): sangría de 1,27 cm en la primera línea de los
párrafos de texto del cuerpo.

La sangría se pone en cada párrafo, no en el estilo Normal (que también usan
títulos y tablas). Si algún estilo traía sangría de primera línea, se le quita.

Qué no lleva sangría se decide con reglas sencillas (la fase 3 las sustituye
por el clasificador de párrafos, y el paso 3.4 rehace esta regla). En caso de
duda, mejor sin sangría: se nota menos que un título sangrado.
"""

import re

from apa_facil import documento
from apa_facil.documento import Estilos, texto_parrafo, w
from apa_facil.reglas.alineacion import parece_portada

SANGRIA = 720                 # 1,27 cm en twips (exacto: media pulgada)
PALABRAS_DE_TITULO = 15       # un párrafo más corto que no acaba en punto parece un título

ESTILOS_DE_TITULO = ("heading", "título", "titulo", "title", "subtitle", "subtítulo", "toc", "caption")
ROTULOS_DE_REFERENCIAS = {"referencias", "referencias bibliográficas", "referencias bibliograficas",
                          "bibliografía", "bibliografia", "webgrafía", "webgrafia", "fuentes",
                          "fuentes consultadas", "bibliografía y webgrafía", "obras citadas",
                          "lista de referencias", "references", "bibliography"}
ROTULOS_DE_RESUMEN = {"resumen", "abstract"}
FIN_DE_REFERENCIAS = re.compile(r"^(anexo|anexos|apéndice|apendice|apéndices|apendices)\b", re.IGNORECASE)
LISTA_A_MANO = re.compile(r"^\s*([•·▪◦‣●○■□\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z]\))\s")
LEYENDA = re.compile(r"^(tabla|figura|gráfico|grafico|imagen|ilustración|ilustracion|cuadro|mapa|"
                     r"fotografía|fotografia)\s+\d|^(nota\.|fuente:|fuentes:)", re.IGNORECASE)
FINAL_DE_FRASE = re.compile(r"[.…:][\"'»”)]*$")


def _normalizar(texto):
    return " ".join(texto.split()).rstrip(":.").lower()


def _capa(estilos, p, etiqueta):
    """El primer elemento `etiqueta` en las capas de formato del párrafo (directo, estilos, por defecto)."""
    for capa in estilos._capas_parrafo(p):
        el = capa.find(w(etiqueta)) if capa is not None else None
        if el is not None:
            return el
    return None


def es_titulo_de_word(estilos, p):
    nivel = _capa(estilos, p, "outlineLvl")
    if nivel is not None and nivel.get(w("val")) not in (None, "9"):
        return True
    nombres = [estilos.nombre(e.get(w("styleId"))).lower() for e in estilos.cadena(estilos.estilo_de(p))]
    return any(c in nombre for nombre in nombres for c in ESTILOS_DE_TITULO)


def _todo_en_negrita(estilos, p):
    runs = [r for r in p.iter(w("r")) if "".join(t.text or "" for t in r.iter(w("t"))).strip()]
    if not runs:
        return False
    for r in runs:
        negrita = next((capa.find(w("b")) for capa in estilos._capas_texto(r, p)
                        if capa is not None and capa.find(w("b")) is not None), None)
        if negrita is None or negrita.get(w("val")) in ("0", "false", "off"):
            return False
    return True


def parece_titulo(estilos, p, texto):
    """Estilo de título de Word, o párrafo corto que no acaba como una frase o va entero en negrita o mayúsculas."""
    if es_titulo_de_word(estilos, p):
        return True
    if len(texto.split()) >= PALABRAS_DE_TITULO:
        return False
    if not FINAL_DE_FRASE.search(texto.rstrip()):
        return True
    return texto.isupper() or _todo_en_negrita(estilos, p)


def _es_lista_de_word(estilos, p):
    numeracion = _capa(estilos, p, "numPr")
    if numeracion is None:
        return False
    num_id = numeracion.find(w("numId"))
    return num_id is None or num_id.get(w("val")) != "0"


def _sangria(estilos, p, atributo):
    try:
        return int(estilos.parrafo(p, "ind", atributo) or 0)
    except ValueError:
        return 0


def motivos_sin_sangria(doc, estilos=None):
    """
    Para cada párrafo del cuerpo, por qué no lleva sangría (o None si la lleva).
    Devuelve una lista de (párrafo, motivo) en el orden del documento.
    """
    estilos = estilos or Estilos(doc)
    parrafos = [p for p in doc.element.body.iter(w("p"))
                if next(p.iterancestors(w("tc"), w("txbxContent")), None) is None]
    portada = set(parece_portada(parrafos, estilos))
    resultado, en_referencias, despues_de_resumen = [], False, False

    for p in parrafos:
        texto = texto_parrafo(p)
        limpio = texto.strip()
        normalizado = _normalizar(limpio)
        titulo = bool(limpio) and parece_titulo(estilos, p, limpio)

        rotulo_de_referencias = normalizado in ROTULOS_DE_REFERENCIAS
        if rotulo_de_referencias:
            en_referencias = True
        elif en_referencias and (es_titulo_de_word(estilos, p) or FIN_DE_REFERENCIAS.match(limpio)):
            en_referencias = False

        if not limpio:
            motivo = "vacío"
        elif p in portada:
            motivo = "portada"
        elif (estilos.parrafo(p, "jc", "val") or "left") in ("center", "right", "end"):
            motivo = "centrado o a la derecha"
        elif es_titulo_de_word(estilos, p) or rotulo_de_referencias:
            motivo = "título"
        elif en_referencias or _sangria(estilos, p, "hanging") > 0:
            motivo = "referencia"
        elif LEYENDA.match(limpio):
            motivo = "leyenda o nota"
        elif _es_lista_de_word(estilos, p):
            motivo = "lista"
        elif titulo and (limpio.isupper() or _todo_en_negrita(estilos, p)):
            motivo = "título"        # "1. Cambios en…" en negrita es un título, no una lista
        elif LISTA_A_MANO.match(texto):
            motivo = "lista"
        elif titulo:
            motivo = "título"
        elif max(_sangria(estilos, p, "left"), _sangria(estilos, p, "start")) > 0:
            motivo = "cita en bloque"
        elif despues_de_resumen:
            motivo = "resumen"
        elif texto.startswith("\t") or texto.startswith("  "):
            motivo = "tabulador"
        else:
            motivo = None

        if limpio:
            despues_de_resumen = titulo and normalizado in ROTULOS_DE_RESUMEN
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


def aplicar(doc, opciones, informe):
    estilos = Estilos(doc)
    clasificados = motivos_sin_sangria(doc, estilos)
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
