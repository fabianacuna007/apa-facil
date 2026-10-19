"""
Paso 7.4: errores de forma en las citas (paso 7.1) y en las referencias (paso
7.2), según la web oficial de APA y la especificación en español del ROADMAP
(«y» entre autores, «s. f.»). Cada error lleva lo escrito y la forma correcta
que se **propone**: no se cambia nada del documento (la fase 7 solo analiza).

En las citas (Parenthetical versus narrative, Author–date citation system,
Direct quotation of material with page numbers):
- falta la coma entre el autor y el año: «(Martínez 2019)»;
- «et al.» mal escrito: «et. al», «et al»;
- «s. f.» mal escrito («s.f», «s.f.») o en inglés («n.d.»; «in press» → «en prensa»);
- «&» en lugar de «y»;
- falta el espacio en «p.45»;
- «p.» con varias páginas o «pp.» con una sola.

En las referencias (Elements of reference list entries, Basic principles of
reference list entries, DOIs and URLs, Missing reference information):
- «&» en lugar de «y», o autores separados por «;»;
- DOI con la forma antigua («doi:», «http://dx.doi.org/»): APA pide https://doi.org/…;
- la ciudad de la editorial, que APA 7 ya no pone;
- «Recuperado de» delante del enlace (sí vale con la fecha de consulta: «Recuperado
  el 30 de septiembre de 2020, de…», para las obras que cambian);
- falta el año, o va sin paréntesis o fuera de su sitio;
- falta el punto tras el autor o tras la fecha;
- «nº 7» en lugar del número entre paréntesis tras el volumen;
- referencias incompletas: un artículo sin volumen ni páginas, un libro de una
  persona sin editorial ni enlace, sin título;
- referencias que no se han podido leer (el 7.2 las marca);
- fuera del orden alfabético (con la clave de orden del paso 6.5).

No se repite lo que ya dice el cruce (7.3): «et al.» con menos de tres personas y
los coautores que faltan o sobran. Una organización sin editorial no se avisa:
APA omite la editorial cuando es la misma que el autor.

Fuera de este paso (decisión del propietario, posible ampliación): la cursiva de
títulos y revistas, y la raya corta (–) en los rangos de páginas.
"""

import re
import shutil
import textwrap
from dataclasses import dataclass
from typing import Optional

from apa_facil.reglas.orden import clave as clave_de_orden
from apa_facil.revision.referencias import ARTICULO, FECHA, OBRA

CITA, REFERENCIA = "cita", "referencia"

COMA = "coma"
ET_AL = "et_al"
SIN_FECHA = "sin_fecha"
AMPERSAND = "ampersand"
ESPACIO_EN_PAGINA = "espacio_en_pagina"
P_O_PP = "p_o_pp"
PUNTO_Y_COMA = "punto_y_coma"
DOI_ANTIGUO = "doi_antiguo"
CIUDAD = "ciudad"
RECUPERADO_DE = "recuperado_de"
FALTA_AÑO = "falta_año"
AÑO_SIN_PARENTESIS = "año_sin_parentesis"
PUNTO_TRAS_AUTOR = "punto_tras_autor"
PUNTO_TRAS_FECHA = "punto_tras_fecha"
NUMERO_ABREVIADO = "numero_abreviado"
INCOMPLETA = "incompleta"
NO_RECONOCIDA = "no_reconocida"
ORDEN = "orden"

TITULOS = {
    COMA: "Falta la coma antes del año",
    ET_AL: "«et al.» mal escrito",
    SIN_FECHA: "Fecha mal escrita",
    AMPERSAND: "«&» en lugar de «y»",
    ESPACIO_EN_PAGINA: "Falta el espacio después de «p.»",
    P_O_PP: "«p.» y «pp.»",
    PUNTO_Y_COMA: "Autores separados por «;»",
    DOI_ANTIGUO: "DOI con la forma antigua",
    CIUDAD: "Ciudad de la editorial",
    RECUPERADO_DE: "«Recuperado de» innecesario",
    FALTA_AÑO: "Falta el año",
    AÑO_SIN_PARENTESIS: "Año sin paréntesis",
    PUNTO_TRAS_AUTOR: "Falta el punto después del autor",
    PUNTO_TRAS_FECHA: "Falta el punto después de la fecha",
    NUMERO_ABREVIADO: "Número de la revista con «nº»",
    INCOMPLETA: "Referencia incompleta",
    NO_RECONOCIDA: "Referencia que no se ha podido leer",
    ORDEN: "Fuera del orden alfabético",
}

PAGINA_SIN_ESPACIO = re.compile(r"^(pp?|págs?)\.(?=\S)")
VARIAS_PAGINAS = re.compile(r"\d\s*[-–—,]\s*\S")
NUMERO_CON_ABREVIATURA = re.compile(r",\s*(?P<escrito>(?:nº|n\.º|n°|núm\.|no\.)\s*(?P<numero>\d+))", re.I)
NUMERO_DE_DOI = re.compile(r"10\.\S+")


@dataclass
class ErrorDeForma:
    tipo: str
    en: str                         # CITA o REFERENCIA
    ubicacion: str                  # «Párrafo 7», «Párrafo 58 (dentro de una tabla)», «Nota al pie 1»
    numero: Optional[int]           # el número del párrafo (None en las notas)
    escrito: str                    # lo que está mal, tal cual
    correccion: Optional[str]       # la forma correcta que se propone (None si no se puede proponer)
    explicacion: str                # por qué, en español
    origen: object = None           # la Cita o la Referencia


def revisar(citas, referencias):
    """Los errores de forma de las citas y de las referencias, en el orden del documento."""
    errores, vistos = [], set()
    for c in citas:
        for error in _errores_de_la_cita(c):
            clave = (error.tipo, error.ubicacion, c.inicio)
            if clave not in vistos:   # las obras de un mismo paréntesis comparten el fragmento
                vistos.add(clave)
                errores.append(error)
    de_las_referencias = [e for r in referencias for e in _errores_de_la_referencia(r)] + _fuera_de_orden(referencias)
    return errores + sorted(de_las_referencias, key=lambda e: e.numero)


# --- Citas ----------------------------------------------------------------------------

def _errores_de_la_cita(c):
    def error(tipo, correccion, explicacion):
        return ErrorDeForma(tipo, CITA, c.ubicacion, c.numero, c.fragmento, correccion, explicacion, c)

    errores = []
    if c.coma_antes_del_año is False:   # solo se mira en las citas entre paréntesis (en las narrativas es None)
        correccion = re.sub(re.escape(c.autores_escritos) + r"\s*" + re.escape(c.año),
                            lambda m: f"{c.autores_escritos}, {c.año}", c.fragmento, count=1)
        errores.append(error(COMA, correccion, "Entre el autor y el año va una coma."))
    if c.et_al and c.et_al_escrito != "et al.":
        errores.append(error(ET_AL, c.fragmento.replace(c.et_al_escrito, "et al.", 1),
                             "Se escribe «et al.»: «et» sin punto y «al.» con punto."))
    correcta = _fecha_correcta(c.año)
    if correcta and correcta != c.año:
        errores.append(error(SIN_FECHA, c.fragmento.replace(c.año, correcta, 1),
                             f"En APA en español se escribe «{correcta}»"
                             + (", con un espacio entre «s.» y «f.»." if correcta == "s. f." else ".")))
    if c.conector == "&":
        errores.append(error(AMPERSAND, re.sub(r"\s*&\s*", " y ", c.fragmento, count=1),
                             "En APA en español, los autores se unen con «y»."))
    if c.localizador:
        sin_espacio = PAGINA_SIN_ESPACIO.match(c.localizador)
        if sin_espacio:
            bien = sin_espacio.group(1) + ". " + c.localizador[sin_espacio.end():]
            errores.append(error(ESPACIO_EN_PAGINA, c.fragmento.replace(c.localizador, bien, 1),
                                 "Entre «p.» (o «pp.») y el número va un espacio."))
        abreviatura = re.match(r"(pp?)\.\s*", c.localizador)
        if abreviatura:
            varias = bool(VARIAS_PAGINAS.search(c.localizador[abreviatura.end():]))
            if varias != (abreviatura.group(1) == "pp"):
                bien = ("pp." if varias else "p.") + " " + c.localizador[abreviatura.end():]
                errores.append(error(P_O_PP, c.fragmento.replace(c.localizador, bien, 1),
                                     "«p.» es para una página, y «pp.» para varias (pp. 34–36)."))
    return errores


def _fecha_correcta(año):
    """«s. f.» o «en prensa» si el año es una forma de ellos; None si es un año."""
    if not año:
        return None
    texto = año.lower().replace(" ", "").replace(".", "")
    if texto in ("sf", "s/f", "nd"):
        return "s. f."
    if texto in ("enprensa", "inpress"):
        return "en prensa"
    return None


# --- Referencias ----------------------------------------------------------------------

def _errores_de_la_referencia(r):
    def error(tipo, escrito, correccion, explicacion):
        return ErrorDeForma(tipo, REFERENCIA, r.ubicacion, r.numero, escrito, correccion, explicacion, r)

    if not r.reconocida:
        return [error(NO_RECONOCIDA, " ".join(r.texto.split()), None,
                      f"{r.motivo} APA pide autor, fecha, título y fuente (y, al final, el DOI o la URL).")]
    errores = []
    if r.conector in ("&", ";"):
        tipo = AMPERSAND if r.conector == "&" else PUNTO_Y_COMA
        explicacion = ("En APA en español, el último autor va unido con «y»." if r.conector == "&" else
                       "Los autores se separan con comas, y el último va unido con «y».")
        errores.append(error(tipo, r.autores_escritos, _autores_en_español(r), explicacion))
    if r.doi and not r.doi.startswith("https://doi.org/"):
        numero = NUMERO_DE_DOI.search(r.doi)
        errores.append(error(DOI_ANTIGUO, r.doi, f"https://doi.org/{numero.group()}" if numero else None,
                             "APA pide el DOI como enlace, con la forma https://doi.org/ seguida del número."))
    if r.lugar:
        errores.append(error(CIUDAD, r.lugar, r.fuente, "APA 7 ya no pone la ciudad de la editorial: quítala."))
    if r.recuperado_de and not re.search(r"\d", r.recuperado_de):
        errores.append(error(RECUPERADO_DE, f"{r.recuperado_de} {r.url or r.doi}", r.url or r.doi,
                             "El enlace basta: APA no pone «Recuperado de» (solo con la fecha de consulta, "
                             "en las obras que cambian con el tiempo)."))
    if r.año is None:
        errores.append(error(FALTA_AÑO, r.autores_escritos or " ".join(r.texto.split()),
                             _con_fecha(r, "s. f.") if r.autores_escritos else None,
                             "Detrás del autor va el año entre paréntesis; si la obra no tiene fecha, «(s. f.)»."))
    elif r.año_entre_parentesis is False:
        errores.append(error(AÑO_SIN_PARENTESIS, r.año, _con_fecha(r, r.año),
                             "El año va entre paréntesis, seguido de un punto"
                             + ("." if r.año_tras_el_autor else ", justo detrás del autor.")))
    fecha = FECHA.search(r.texto) if r.año_entre_parentesis else None
    if fecha:
        antes = r.texto[:fecha.start()].rstrip()
        if antes and not antes.endswith("."):
            errores.append(error(PUNTO_TRAS_AUTOR, f"{antes} {fecha.group()}", f"{antes}. {fecha.group()}",
                                 "El autor termina con un punto, también el nombre de una organización."))
        if not r.texto[fecha.end():].lstrip().startswith("."):
            errores.append(error(PUNTO_TRAS_FECHA, fecha.group(), fecha.group() + ".",
                                 "Detrás de la fecha entre paréntesis va un punto."))
    numero = NUMERO_CON_ABREVIATURA.search(r.texto) if r.tipo == ARTICULO else None
    if numero:
        bien = f"{r.volumen}({numero.group('numero')})" if r.volumen else f"({numero.group('numero')})"
        errores.append(error(NUMERO_ABREVIADO, numero.group("escrito"), bien,
                             "El número de la revista va entre paréntesis justo detrás del volumen, sin «nº»."))
    falta = _lo_que_falta(r)
    if falta:
        errores.append(error(INCOMPLETA, " ".join(r.texto.split()), None, falta))
    return errores


def _autores_en_español(r):
    nombres = [f"{a.apellidos}, {a.iniciales}" if a.iniciales else a.apellidos for a in r.autores]
    if len(nombres) == 1:
        return nombres[0]
    return ", ".join(nombres[:-1]) + " y " + nombres[-1] + (f" ({r.rol})" if r.rol else "")


def _con_fecha(r, año):
    """«Lefèvre, A. (2009).», «Consejo Económico de Ejemplo. (2022).»: el autor, con su punto, y la fecha."""
    return f"{r.autores_escritos.rstrip('.')}. ({año})."


def _lo_que_falta(r):
    """Lo que le falta a una referencia leída, como frase en español; None si está completa."""
    if r.titulo is None:
        return "Falta el título."
    if r.tipo == ARTICULO and not r.volumen and not r.paginas and not r.articulo:
        return "Faltan el volumen y las páginas del artículo."
    if r.tipo == ARTICULO and r.volumen and not r.paginas and not r.articulo and not (r.doi or r.url):
        return "Faltan las páginas del artículo."
    if r.tipo == OBRA and not r.grupo and not r.fuente and not (r.doi or r.url):
        return "Falta la editorial (o el enlace, si es una obra de internet)."
    return None


# --- Orden alfabético -----------------------------------------------------------------

def _fuera_de_orden(referencias):
    """
    Las referencias fuera de su sitio en cada lista: las que no forman parte de la
    secuencia ordenada más larga (así, en «Molina, Collantes, Consejo…» la que está
    mal es Molina, no las demás). La clave de orden es la del paso 6.5.
    """
    errores = []
    for lista in sorted({r.lista for r in referencias}):
        de_la_lista = [r for r in referencias if r.lista == lista]
        claves = [clave_de_orden(r.texto) for r in de_la_lista]
        bien = _ordenadas_mas_larga(claves)
        ordenadas = sorted(range(len(de_la_lista)), key=lambda i: claves[i])
        for i, r in enumerate(de_la_lista):
            if i in bien:
                continue
            posicion = ordenadas.index(i)
            anterior = de_la_lista[ordenadas[posicion - 1]] if posicion > 0 else None
            siguiente = de_la_lista[ordenadas[posicion + 1]] if posicion + 1 < len(ordenadas) else None
            if anterior and siguiente:
                donde = f"después de «{_nombre(anterior)}» y antes de «{_nombre(siguiente)}»"
            elif siguiente:
                donde = f"la primera, antes de «{_nombre(siguiente)}»"
            else:
                donde = f"la última, después de «{_nombre(anterior)}»" if anterior else "en otro sitio"
            errores.append(ErrorDeForma(ORDEN, REFERENCIA, r.ubicacion, r.numero, _nombre(r), None,
                                        f"Las referencias van en orden alfabético por el apellido del primer "
                                        f"autor: esta debería ir {donde}.", r))
    return errores


def _ordenadas_mas_larga(claves):
    """Las posiciones de la secuencia más larga que ya está en orden (no tiene por qué ser seguida)."""
    if not claves:
        return set()
    largo, anterior = [1] * len(claves), [-1] * len(claves)
    for i in range(len(claves)):
        for j in range(i):
            if claves[j] <= claves[i] and largo[j] + 1 > largo[i]:
                largo[i], anterior[i] = largo[j] + 1, j
    i = max(range(len(claves)), key=lambda k: (largo[k], -k))
    posiciones = set()
    while i != -1:
        posiciones.add(i)
        i = anterior[i]
    return posiciones


def _nombre(r):
    if r.autores_escritos:
        return r.autores_escritos
    return textwrap.shorten(" ".join(r.texto.split()), 50, placeholder="…")


# --- Opción --forma -------------------------------------------------------------------

def como_texto(errores, ancho=None):
    """Los errores para la terminal, en dos apartados: citas y referencias."""
    ancho = max(ancho or shutil.get_terminal_size((100, 24)).columns - 1, 40)
    if not errores:
        return "No hay errores de forma en las citas ni en las referencias."
    lineas = []
    for en, titulo in ((CITA, "En las citas"), (REFERENCIA, "En las referencias")):
        de_aqui = [e for e in errores if e.en == en]
        lineas += ["" if lineas else None, f"{titulo} ({len(de_aqui)})"]
        if not de_aqui:
            lineas.append("  · ninguno")
        for e in de_aqui:
            lugar = f"párrafo {e.numero}" + (" (tabla)" if "tabla" in e.ubicacion else "") \
                if e.numero is not None else e.ubicacion.lower()
            lineas += textwrap.wrap(f"{lugar} · {TITULOS[e.tipo]}", width=ancho, initial_indent="  · ",
                                    subsequent_indent="    ")
            muestra = f"«{e.escrito}»" + (f" → «{e.correccion}»" if e.correccion else "")
            lineas += textwrap.wrap(muestra, width=ancho, initial_indent="    ", subsequent_indent="      ")
            lineas += textwrap.wrap(e.explicacion, width=ancho, initial_indent="    ↳ ", subsequent_indent="      ")
    return "\n".join(l for l in lineas if l is not None)
