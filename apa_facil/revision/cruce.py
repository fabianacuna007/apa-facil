"""
Paso 7.3: cruce de las citas del texto (paso 7.1) con la lista de referencias
(paso 7.2), según la web oficial de APA (Basic principles of citation: cada obra
citada debe estar en la lista y cada referencia debe citarse, con la misma
grafía de los autores y la misma fecha; Author–date citation system: cada cita
corresponde a una sola referencia; Personal communications: estas no llevan
referencia, y el detector del 7.1 no las cuenta como citas).

Solo compara: no cambia nada. Las citas se agrupan por obra (el mismo autor y el
mismo año, aunque estén escritos de forma distinta: «Gómez (s.f)» y «Gómez,
s.f.»), y cada obra se busca en la lista:

- coincide (los autores y el año son los de una sola referencia): no se informa;
- cita sin referencia;
- años distintos: los autores son los de una referencia, pero el año no (o la
  referencia no tiene año);
- autores escritos de forma distinta: el año coincide y los autores se parecen
  (tildes, un apellido de más o de menos, una errata, coautores que faltan);
- ambigua: puede ser de varias referencias.

Y al revés: las referencias que no cita ninguna obra, y las que el 7.2 no pudo
leer (no se pueden cruzar, así que no se dan por no citadas).

Que «s. f.» esté mal escrito, o «&» en lugar de «y», no importa aquí: es forma, y
la revisa el paso 7.4. Aquí solo importa a qué obra se refiere cada cita.
"""

import re
import shutil
import textwrap
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

from apa_facil.revision.citas import ABREVIATURA

EMPAREJADA = "emparejada"
SIN_REFERENCIA = "sin_referencia"
NO_CITADA = "no_citada"
AÑO_DISTINTO = "año_distinto"
AUTORES_DISTINTOS = "autores_distintos"
AMBIGUA = "ambigua"
NO_RECONOCIDA = "no_reconocida"

TITULOS = {
    SIN_REFERENCIA: "Citas sin referencia en la lista",
    NO_CITADA: "Referencias que no se citan en el texto",
    AÑO_DISTINTO: "Años que no coinciden",
    AUTORES_DISTINTOS: "Autores escritos de forma distinta",
    AMBIGUA: "Citas que pueden ser de varias referencias",
    NO_RECONOCIDA: "Referencias que no se han podido leer (no se pueden cruzar)",
}
ERRATA_DESDE = 5   # letras: en apellidos más cortos, una letra distinta es otro apellido


@dataclass
class Hallazgo:
    """Una obra citada (sus citas, en todas sus ubicaciones) y las referencias con que se ha cruzado."""
    tipo: str
    citas: list = field(default_factory=list)         # [Cita], en orden
    referencias: list = field(default_factory=list)   # [Referencia]
    detalle: Optional[str] = None                     # qué no coincide, en español


@dataclass
class Cruce:
    emparejadas: list = field(default_factory=list)
    sin_referencia: list = field(default_factory=list)
    no_citadas: list = field(default_factory=list)
    años_distintos: list = field(default_factory=list)
    autores_distintos: list = field(default_factory=list)
    ambiguas: list = field(default_factory=list)
    no_reconocidas: list = field(default_factory=list)

    def grupos(self):
        """(tipo, hallazgos) de lo que hay que revisar, en el orden en que se muestra."""
        return [(SIN_REFERENCIA, self.sin_referencia), (NO_CITADA, self.no_citadas),
                (AÑO_DISTINTO, self.años_distintos), (AUTORES_DISTINTOS, self.autores_distintos),
                (AMBIGUA, self.ambiguas), (NO_RECONOCIDA, self.no_reconocidas)]


def cruzar(citas, referencias):
    """Cruza las citas (paso 7.1) con las referencias (paso 7.2). Devuelve un Cruce."""
    cruce = Cruce()
    abreviaturas = {c.abreviatura: _sin_abreviatura(c.autores_escritos) for c in citas if c.abreviatura}
    leidas = [r for r in referencias if r.reconocida]
    citadas = set()
    for obra in _obras(citas, abreviaturas):
        hallazgo = _buscar(obra, leidas, abreviaturas)
        citadas.update(id(r) for r in hallazgo.referencias)
        {EMPAREJADA: cruce.emparejadas, SIN_REFERENCIA: cruce.sin_referencia, AÑO_DISTINTO: cruce.años_distintos,
         AUTORES_DISTINTOS: cruce.autores_distintos, AMBIGUA: cruce.ambiguas}[hallazgo.tipo].append(hallazgo)
    cruce.no_citadas = [Hallazgo(NO_CITADA, referencias=[r]) for r in leidas if id(r) not in citadas]
    cruce.no_reconocidas = [Hallazgo(NO_RECONOCIDA, referencias=[r]) for r in referencias if not r.reconocida]
    return cruce


# --- Las obras citadas ----------------------------------------------------------------

def _obras(citas, abreviaturas):
    """Las citas agrupadas por obra: los mismos autores (como están escritos) y el mismo año."""
    obras = {}
    for c in citas:
        clave = (tuple(" ".join(n.split()) for n in _nombres(c, abreviaturas)), c.et_al, _año(c.año))
        obras.setdefault(clave, []).append(c)
    return list(obras.values())


def _nombres(cita, abreviaturas):
    """Los apellidos de la cita, con las abreviaturas definidas en el texto puestas por el nombre entero."""
    escritos = _sin_abreviatura(cita.autores_escritos)
    if escritos in abreviaturas:
        return [abreviaturas[escritos]]
    return cita.autores


def _sin_abreviatura(texto):
    return " ".join(ABREVIATURA.sub(" ", texto).split())


def _año(año):
    """El año para comparar: «s. f.», «s.f» y «n.d.» son lo mismo (cómo se escriben lo revisa el 7.4)."""
    if año is None:
        return None
    texto = año.lower().replace(" ", "").replace(".", "")
    if texto in ("sf", "s/f", "nd"):
        return "sf"
    if texto in ("enprensa", "inpress"):
        return "prensa"
    return texto


def _sin_letra(año):
    return re.sub(r"(?<=\d)[a-z]$", "", año) if año else año


# --- Búsqueda de una obra en la lista -------------------------------------------------

def _buscar(obra, leidas, abreviaturas):
    cita = obra[0]
    año = _año(cita.año)
    comparadas = [(r, _comparar_autores(cita, r, abreviaturas)) for r in leidas]
    iguales = [r for r, comparacion in comparadas if comparacion == IGUALES]

    exactas = [r for r in iguales if _año(r.año) == año]
    if len(exactas) == 1:
        return Hallazgo(EMPAREJADA, obra, exactas)
    if exactas:
        return Hallazgo(AMBIGUA, obra, exactas, "hay varias referencias con los mismos autores y el mismo año; "
                                                "APA las distingue con una letra detrás del año (2017a, 2017b)")
    sin_letra = [r for r in iguales if _sin_letra(_año(r.año)) == año and año]
    if len(sin_letra) > 1:
        return Hallazgo(AMBIGUA, obra, sin_letra, "a la cita le falta la letra del año: en la lista hay "
                        + " y ".join(r.año for r in sin_letra))
    if iguales:
        return Hallazgo(AÑO_DISTINTO, obra, iguales, "; ".join(
            "la referencia no tiene año" if r.año is None else f"en la referencia, {r.año}" for r in iguales))

    parecidas = [(r, comparacion) for r, comparacion in comparadas if comparacion and _año(r.año) == año]
    if len(parecidas) == 1:
        referencia, detalle = parecidas[0]
        return Hallazgo(AUTORES_DISTINTOS, obra, [referencia], detalle)
    if parecidas:
        return Hallazgo(AMBIGUA, obra, [r for r, _ in parecidas],
                        "los autores se parecen a los de varias referencias del mismo año")

    otro_año = [r for r, comparacion in comparadas if comparacion]
    detalle = None
    if otro_año:
        detalle = "se parece a " + " y a ".join(f"«{referencia_corta(r)}»" for r in otro_año) \
                  + " de la lista, pero no coinciden ni los autores ni el año"
    return Hallazgo(SIN_REFERENCIA, obra, [], detalle)


IGUALES = "iguales"


def _comparar_autores(cita, ref, abreviaturas):
    """
    IGUALES si los autores de la cita son los de la referencia; si solo se parecen,
    un texto que dice en qué se diferencian; None si no tienen que ver.
    """
    if ref.grupo:
        return _comparar_grupo(cita, ref, abreviaturas)
    nombres = _nombres(cita, abreviaturas)
    de_la_referencia = [a.apellidos for a in ref.autores]
    if not nombres or not de_la_referencia:
        return None
    if cita.et_al:
        primero = _comparar_nombre(nombres[0], de_la_referencia[0])
        if primero is None:
            return None
        if len(de_la_referencia) < 3:
            return (f"la cita pone «et al.», pero la referencia tiene {_numero_de_autores(len(de_la_referencia))}: "
                    + _lista(de_la_referencia))
        return primero
    comparaciones = [_comparar_nombre(a, b) for a, b in zip(nombres, de_la_referencia)]
    if any(c is None for c in comparaciones):
        return None
    if len(nombres) < len(de_la_referencia):
        return (f"faltan autores en la cita: la referencia tiene {_numero_de_autores(len(de_la_referencia))} "
                f"({_lista(de_la_referencia)})")
    if len(nombres) > len(de_la_referencia):
        return (f"sobran autores en la cita: la referencia tiene {_numero_de_autores(len(de_la_referencia))} "
                f"({_lista(de_la_referencia)})")
    return next((c for c in comparaciones if c != IGUALES), IGUALES)


def _comparar_grupo(cita, ref, abreviaturas):
    """Una organización (o una obra sin autor, citada por el principio de su título)."""
    escritos = _sin_abreviatura(cita.autores_escritos)
    escritos = abreviaturas.get(escritos, escritos)
    if cita.et_al:
        return None
    if escritos == ref.autores_escritos:
        return IGUALES
    sin_comillas = escritos.strip("“”\"«»‘’'")
    if sin_comillas != escritos and _normal(ref.autores_escritos).startswith(_normal(sin_comillas)):
        return IGUALES   # una obra sin autor, citada por el principio de su título entre comillas (APA)
    if _normal(escritos) == _normal(ref.autores_escritos):
        return (f"«{escritos}» en la cita y «{ref.autores_escritos}» en la referencia: cambian las tildes, las "
                f"mayúsculas o los guiones")
    palabras, de_la_referencia = _normal(sin_comillas).split(), _normal(ref.autores_escritos).split()
    if len(palabras) >= 2 and de_la_referencia[:len(palabras)] == palabras:
        return "la cita solo tiene el principio del nombre (o del título)"
    return None


def _comparar_nombre(de_la_cita, de_la_referencia):
    """IGUALES, un texto con la diferencia si se parecen, o None."""
    a, b = " ".join(de_la_cita.split()), " ".join(de_la_referencia.split())
    if a == b:
        return IGUALES
    if _normal(a) == _normal(b):
        return f"«{a}» en la cita y «{b}» en la referencia: cambian las tildes, las mayúsculas o los guiones"
    palabras_a, palabras_b = _normal(a).split(), _normal(b).split()
    if palabras_a and palabras_b and (_contiene(palabras_b, palabras_a) or _contiene(palabras_a, palabras_b)):
        return f"«{a}» en la cita y «{b}» en la referencia: falta o sobra un apellido"
    if min(len(_normal(a)), len(_normal(b))) >= ERRATA_DESDE and _una_letra(_normal(a), _normal(b)):
        return f"«{a}» en la cita y «{b}» en la referencia: parece una errata"
    return None


def _normal(texto):
    """Sin tildes, en minúsculas y sin guiones ni apóstrofos: «Sáenz-Gil» → «saenz gil»."""
    sin_tildes = "".join(ch for ch in unicodedata.normalize("NFD", texto) if unicodedata.category(ch) != "Mn")
    return " ".join(re.sub(r"[-’'‐]", " ", sin_tildes.lower()).split())


def _contiene(largas, cortas):
    """¿Están las palabras `cortas` seguidas dentro de `largas` (y son menos)?"""
    return len(cortas) < len(largas) and any(largas[i:i + len(cortas)] == cortas
                                             for i in range(len(largas) - len(cortas) + 1))


def _una_letra(a, b):
    """¿Se diferencian en una sola letra (cambiada, de más o de menos)?"""
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    corta, larga = sorted((a, b), key=len)
    return any(larga[:i] + larga[i + 1:] == corta for i in range(len(larga)))


def _numero_de_autores(n):
    return "un solo autor" if n == 1 else f"{n} autores"


def _lista(nombres):
    return nombres[0] if len(nombres) == 1 else ", ".join(nombres[:-1]) + " y " + nombres[-1]


# --- Opción --cruce -------------------------------------------------------------------

def como_texto(cruce, ancho=None):
    """El cruce para la terminal: un apartado por cada grupo, con cada obra y dónde aparece."""
    ancho = max(ancho or shutil.get_terminal_size((100, 24)).columns - 1, 40)
    lineas = [f"Coinciden con su referencia: {len(cruce.emparejadas)} obra(s) citada(s)."]
    vacios = [TITULOS[tipo] for tipo, hallazgos in cruce.grupos() if not hallazgos]
    if len(vacios) == len(TITULOS):
        return "\n".join(lineas + ["No hay nada que revisar: todas las citas tienen su referencia y todas las "
                                   "referencias se citan."])
    for tipo, hallazgos in cruce.grupos():
        if not hallazgos:
            continue
        lineas += ["", f"{TITULOS[tipo]} ({len(hallazgos)})"]
        for h in hallazgos:
            lineas += textwrap.wrap(_describir(h), width=ancho, initial_indent="  · ", subsequent_indent="    ")
            if h.detalle:
                lineas += textwrap.wrap(h.detalle, width=ancho, initial_indent="    ↳ ", subsequent_indent="      ")
    lineas += [""] + textwrap.wrap("Sin nada que revisar: " + "; ".join(t[0].lower() + t[1:] for t in vacios) + ".",
                                   width=ancho, subsequent_indent="  ")
    return "\n".join(lineas)


def _describir(h):
    partes = []
    if h.citas:
        cita = h.citas[0]
        partes.append(f"{cita.autores_escritos} ({cita.año}) en el texto: {_donde(h.citas)}")
    for r in h.referencias:
        if h.tipo == NO_RECONOCIDA:
            partes.append(f"«{textwrap.shorten(' '.join(r.texto.split()), 70, placeholder='…')}», párrafo {r.numero}")
        else:
            partes.append(("en la lista: " if h.citas else "") + f"{referencia_corta(r)}, párrafo {r.numero}")
    return "; ".join(partes)


def referencia_corta(r):
    return f"{r.autores_escritos} ({r.año})" if r.año else f"{r.autores_escritos} (sin año)"


def _donde(citas):
    """«párrafos 12 y 58 (tabla)», «nota al pie 1»…, sin repetir."""
    lugares = []
    for c in citas:
        lugar = (f"{c.numero}" + (" (tabla)" if "tabla" in c.ubicacion else "")) if c.numero is not None \
            else c.ubicacion.lower()
        if lugar not in lugares:
            lugares.append(lugar)
    numeros = [l for l in lugares if l[0].isdigit()]
    otros = [l for l in lugares if not l[0].isdigit()]
    texto = []
    if numeros:
        texto.append(("párrafo " if len(numeros) == 1 else "párrafos ") + _lista(numeros))
    texto += otros
    return ", ".join(texto)
