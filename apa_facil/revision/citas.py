"""
Paso 7.1: detector de citas en el texto, según la web oficial de APA
(Author–date citation system, Parenthetical versus narrative in-text citations,
Citing multiple works, Quotations y Direct quotation of material with page
numbers).

Solo lee: no cambia nada del documento. Devuelve, para cada obra citada, dónde
está la cita y lo que tiene escrito (autores, año, páginas), **tal como lo
escribió el estudiante**. No juzga si está bien: eso lo hace el paso 7.4. Por eso
también reconoce las formas mal escritas que tiene que poder señalar después:
«(Martínez 2019)» (sin coma), «et. al», «s.f», «p.45», «&» en lugar de «y».

Dos formatos (APA):
- entre paréntesis: «(Koehler, 2016)», el autor y el año separados por una coma,
  varias obras separadas por «;» y, si el autor se repite, solo los años:
  «(Department of Veterans Affairs, n.d., 2017a, 2017b, 2019)»;
- narrativa: «Koehler (2016)», el nombre en la frase y el año entre paréntesis
  justo detrás.

Un paréntesis es una cita si tiene un año aislado (o «s. f.», «n.d.», «en
prensa») y, delante, un autor que empieza por mayúscula y no lleva números: así
no cuentan «(112 mujeres y 68 hombres)», «(PSQI)», «(N = 180)» ni «(2001-2024)».
Se busca en el cuerpo, el resumen, los apéndices (con sus tablas) y las notas al
pie; no en la portada, el índice ni la lista de referencias (paso 7.2).
"""

import re
import shutil
import textwrap
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from apa_facil import estructura
from apa_facil.documento import notas, texto_parrafo, w

ZONAS_CON_CITAS = ("cuerpo", "resumen", "anexos")

MAYUS = "A-ZÁÉÍÓÚÜÑÀÈÌÒÙÂÊÎÔÛÄËÏÖÇÅØÆŒŠŽ"
COMILLAS = "“\"«‘'"

# Un año: de 1500 a 2099, con letra si hace falta (2017a); o las obras sin fecha y en prensa.
AÑO = (r"(?:(?<![\w/\-–.])(?:1[5-9]|20)\d\d[a-z]?(?![\w\-–/])"
       r"|(?<![\w.])(?:s\.\s?f\b\.?|s/f\b|n\.\s?d\b\.?|en\s+prensa\b|in\s+press\b))")
AÑOS = rf"{AÑO}(?:\s*,\s*{AÑO})*"
# Páginas u otra ubicación: «p. 45», «p.45», «pp. 34–36», «pp. 67, 72», «p. S41», «párr. 3».
LOCALIZADOR = r"(?:pp?|págs?|párrs?|paras?)\.?\s*[A-Za-z]?\d[\w–—\-]*(?:\s*,\s*[A-Za-z]?\d[\w–—\-]*)*"

PARENTESIS = re.compile(r"\(([^()]*)\)")
# Una cita dentro de un paréntesis: [texto delante] autores[,] años[, páginas][, más texto].
CITA_ENTRE_PARENTESIS = re.compile(rf"""
    ^\s*(?P<prefijo>(?:[^\s{MAYUS}{COMILLAS}]\S*\s+)*?)
    (?P<autores>[{MAYUS}{COMILLAS}][^\d;]*?)
    (?P<coma>\s*,)?\s*
    (?P<años>{AÑOS})
    (?P<resto>[\s,].*)?$""", re.X | re.S)
# El paréntesis de una cita narrativa: solo años y, si acaso, las páginas: «(2016)», «(2017, p. 3)».
AÑOS_SOLOS = re.compile(rf"^\s*(?P<años>{AÑOS})\s*(?:,?\s*(?P<localizador>{LOCALIZADOR}))?\s*$", re.S)
BUSCAR_LOCALIZADOR = re.compile(rf"^[\s,]*(?P<localizador>{LOCALIZADOR})")
UN_AÑO = re.compile(AÑO)
ET_AL = re.compile(r"\bet\.?\s*al\b\.?")
ABREVIATURA = re.compile(r"\s*\[([^\]]+)\]\s*")
SEPARA_AUTORES = re.compile(r"\s*,\s*(?:y|e|&|and)\s+|\s*,\s*|\s+(?:y|e|and)\s+|\s*&\s*")
PALABRA = re.compile(r"et\.?\s*al\b\.?|[\w’'\-]+|\S")
MAXIMO_DE_PALABRAS = 12

# Palabras en minúscula que pueden ir entre las de un nombre: «Ortega y Blanco»,
# «Organización Mundial de la Salud», «Van der Berg».
ENLACES = {"y", "e", "&", "and", "de", "del", "la", "las", "los", "el", "da", "das", "do", "dos", "di",
           "van", "von", "der", "den", "le", "du", "of", "the", "for", "para", "sobre"}
Y = {"y", "e", "&", "and"}
# Palabras con mayúscula que empiezan una frase sin ser parte del nombre:
# «Según Ortega y Blanco (2018)» → Ortega y Blanco.
NO_SON_NOMBRES = {
    "según", "como", "para", "por", "en", "tal", "así", "ya", "pero", "aunque", "mientras", "tanto", "tras",
    "con", "de", "del", "desde", "sin", "sobre", "ante", "bajo", "entre", "hacia", "hasta", "también", "además",
    "asimismo", "el", "la", "los", "las", "lo", "un", "una", "unos", "unas", "este", "esta", "estos", "estas",
    "ese", "esa", "esos", "esas", "aquel", "aquella", "al", "y", "e", "o", "ni", "si", "cuando", "donde",
    "incluso", "solo", "sólo", "igualmente", "finalmente", "luego", "hoy", "ahora", "recientemente",
    "concretamente", "precisamente", "siguiendo", "citando", "véase", "ver", "cf", "according", "as", "in",
    "the", "a", "an", "see", "also", "both", "like", "following", "and", "for", "by",
}

ENTRE_PARENTESIS = "parentetica"
NARRATIVA = "narrativa"
SIN_PARENTESIS = "sin_parentesis"   # una celda de tabla que solo tiene la cita: «Gómez, s.f.»
NOMBRES_DE_TIPO = {ENTRE_PARENTESIS: "paréntesis", NARRATIVA: "narrativa", SIN_PARENTESIS: "sin paréntesis"}


@dataclass
class Cita:
    """Una obra citada en el texto, con todo tal como está escrito."""
    ubicacion: str                  # «Párrafo 7», «Párrafo 45 (dentro de una tabla)», «Nota al pie 1»
    numero: Optional[int]           # el número del párrafo (el del diagnóstico); None en las notas
    zona: str                       # la del clasificador, o «notas»
    tipo: str                       # ENTRE_PARENTESIS o NARRATIVA
    inicio: int                     # dónde empieza y acaba la cita en el texto del párrafo: todo el
    fin: int                        # paréntesis y, en las narrativas, también el nombre de delante
    fragmento: str                  # ese trozo del texto: «(Navarro y Gil, 2020; Castro et al., 2019)»
    autores_escritos: str           # «Navarro y Gil», «Pérez et. al», «Consejo Económico de Ejemplo»
    autores: list = field(default_factory=list)   # [«Navarro», «Gil»], sin «et al.» ni abreviatura
    año: str = ""                   # «2019», «2017a», «s. f.», «s.f», «n.d.», «en prensa»
    et_al: bool = False
    et_al_escrito: Optional[str] = None   # «et al.», «et. al», «et al»
    conector: Optional[str] = None        # «y», «&», «e», «and»
    abreviatura: Optional[str] = None     # «NIMH» en «National Institute of Mental Health [NIMH]»
    coma_antes_del_año: Optional[bool] = None   # solo entre paréntesis: «(Martínez 2019)» → False
    localizador: Optional[str] = None     # «p. 112», «p.45», «pp. 34–36»
    prefijo: Optional[str] = None         # el texto de delante en el paréntesis: «véase», «see also»
    grupo: int = 0                  # las citas del mismo paréntesis (o del mismo nombre) comparten grupo
    obras_en_el_grupo: int = 1


def detectar(doc, clasificacion=None):
    """
    Las citas del texto de un documento abierto con python-docx, en orden (una
    por obra citada). `clasificacion` es la de `estructura.clasificar`; si no se
    da, se clasifica el documento tal como está.
    """
    if clasificacion is None:
        clasificacion = estructura.clasificar(doc)
    zona_de = {c.elemento: c.zona for c in clasificacion}
    citas, grupos = [], iter(range(1, 1_000_000))
    zona = None
    for numero, p in enumerate(doc.element.body.iter(w("p")), start=1):
        # Los párrafos de las tablas y de los cuadros de texto no se clasifican:
        # van en la zona del último párrafo clasificado.
        zona = zona_de.get(p, zona)
        if zona not in ZONAS_CON_CITAS:
            continue
        celda = next(p.iterancestors(w("tc")), None) is not None
        ubicacion = f"Párrafo {numero}" + (" (dentro de una tabla)" if celda else "")
        citas += citas_del_texto(texto_parrafo(p), ubicacion, numero, zona, grupos, celda)
    for ubicacion, nota in notas(doc):
        for p in nota.iter(w("p")):
            citas += citas_del_texto(texto_parrafo(p), ubicacion, None, "notas", grupos)
    return citas


def citas_del_texto(texto, ubicacion="", numero=None, zona="cuerpo", grupos=None, celda=False):
    """
    Las citas de un texto (un párrafo), en el orden en que aparecen. En una celda
    de tabla (`celda`), también la que ocupa toda la celda sin paréntesis, como en
    una columna «Fuente»: «Gómez, s.f.».
    """
    grupos = grupos if grupos is not None else iter(range(1, 1_000_000))
    trozos = [_narrativas(texto, p) or _narrativa_con_abreviatura(texto, p) or _entre_parentesis(texto, p)
              for p in PARENTESIS.finditer(texto)]
    if celda and not any(trozos):
        trozos = [_sin_parentesis(texto)]
    citas = []
    for nuevas in filter(None, trozos):
        grupo = next(grupos)
        for cita in nuevas:
            cita.ubicacion, cita.numero, cita.zona = ubicacion, numero, zona
            cita.grupo, cita.obras_en_el_grupo = grupo, len(nuevas)
        citas += nuevas
    return citas


# --- Entre paréntesis -----------------------------------------------------------------

def _entre_parentesis(texto, parentesis):
    """Las obras de un paréntesis como «(Molina, 2017; Sáenz Gil, 2020)»; [] si no es una cita."""
    citas = []
    for parte in parentesis.group(1).split(";"):
        m = CITA_ENTRE_PARENTESIS.match(parte)
        if not m or not _parece_un_autor(m.group("autores")):
            continue
        resto = BUSCAR_LOCALIZADOR.match(m.group("resto") or "")
        for año in _años(m.group("años")):
            cita = _cita(texto, parentesis.start(), parentesis.end(), ENTRE_PARENTESIS, m.group("autores"), año)
            cita.coma_antes_del_año = m.group("coma") is not None
            cita.localizador = resto.group("localizador").strip() if resto else None
            cita.prefijo = m.group("prefijo").strip() or None
            citas.append(cita)
    return citas


def _sin_parentesis(texto):
    """
    Una celda de tabla que es solo una cita: «Pérez et. al, 2020». [] si no. Sin
    paréntesis hace falta la coma: «Población 2001» es un encabezado de columna.
    """
    m = CITA_ENTRE_PARENTESIS.match(texto)
    resto = (m.group("resto") or "").strip(" ,.") if m else ""
    if (not m or m.group("prefijo") or not m.group("coma") or not _parece_un_autor(m.group("autores"))
            or resto and not BUSCAR_LOCALIZADOR.fullmatch(resto)):
        return []
    citas = []
    for año in _años(m.group("años")):
        inicio = len(texto) - len(texto.lstrip())
        cita = _cita(texto, inicio, len(texto.rstrip()), SIN_PARENTESIS, m.group("autores"), año)
        cita.coma_antes_del_año = m.group("coma") is not None
        cita.localizador = resto or None
        citas.append(cita)
    return citas


def _parece_un_autor(autores):
    autores = autores.strip()
    palabras = autores.split()
    return (bool(palabras) and len(palabras) <= MAXIMO_DE_PALABRAS
            and not any(c in autores for c in "=%:+<>")
            and palabras[-1].lower() not in ENLACES - {"&"})


# --- Narrativas -----------------------------------------------------------------------

def _narrativas(texto, parentesis):
    """
    «Koehler (2016)», «Gutiérrez (2012, 2017)», «Molina (2017, p. 3)»: el paréntesis
    solo lleva años, y justo delante hay un nombre. [] si no.
    """
    m = AÑOS_SOLOS.match(parentesis.group(1))
    if not m:
        return []
    nombre = _nombre_de_delante(texto, parentesis.start())
    if nombre is None:
        return []
    inicio, autores = nombre
    citas = []
    for año in _años(m.group("años")):
        cita = _cita(texto, inicio, parentesis.end(), NARRATIVA, autores, año)
        cita.localizador = m.group("localizador")
        citas.append(cita)
    return citas


def _narrativa_con_abreviatura(texto, parentesis):
    """
    «National Institute of Mental Health (NIMH, 2020)»: la primera vez que se
    nombra una organización, con su abreviatura en el paréntesis (APA). Solo si
    la abreviatura son las iniciales del nombre de delante.
    """
    m = CITA_ENTRE_PARENTESIS.match(parentesis.group(1))
    if not m or m.group("prefijo") or not re.fullmatch(rf"[{MAYUS}]{{2,}}", m.group("autores").strip()):
        return []
    nombre = _nombre_de_delante(texto, parentesis.start())
    if nombre is None:
        return []
    inicio, autores = nombre
    iniciales = "".join(p[0] for p in autores.split() if p[0].isupper())
    if iniciales != m.group("autores").strip():
        return []
    citas = []
    for año in _años(m.group("años")):
        cita = _cita(texto, inicio, parentesis.end(), NARRATIVA, autores, año)
        cita.abreviatura = m.group("autores").strip()
        resto = BUSCAR_LOCALIZADOR.match(m.group("resto") or "")
        cita.localizador = resto.group("localizador").strip() if resto else None
        citas.append(cita)
    return citas


def _nombre_de_delante(texto, posicion):
    """
    (inicio, nombre) del nombre que va justo delante de `posicion` (separado solo
    por espacios): las palabras con mayúscula, «et al.» y los enlaces («y», «de»…),
    hasta la primera palabra en minúscula o el primer signo. Se quitan los enlaces
    y las palabras que empiezan frase que queden al principio («Según»). None si
    no queda ningún nombre.
    """
    desplazamiento = max(0, posicion - 200)
    antes = texto[desplazamiento:posicion]
    palabras = list(PALABRA.finditer(antes))
    if not palabras or antes[palabras[-1].end():].strip():
        return None
    nombre = []
    for palabra in reversed(palabras):
        valor = palabra.group()
        # La coma solo en «Pérez, Gómez y Ruiz»: con un «y» a su derecha. Así
        # «Por su parte, Herranz (2015)» o «En España, Pérez (2019)» no la cruzan.
        coma = valor == "," and any(p.group() in Y for p in nombre) and nombre[0].group() != ","
        if ET_AL.fullmatch(valor) or (valor[0].isupper() and _letras(valor)) or valor in ENLACES or coma:
            nombre.insert(0, palabra)
        else:
            break
    while nombre:
        primera = nombre[0].group()
        if primera in ENLACES or primera == "," or (primera.lower() in NO_SON_NOMBRES and len(nombre) > 1):
            nombre.pop(0)
        else:
            break
    if (not nombre or nombre[-1].group() in ENLACES
            or len(nombre) == 1 and nombre[0].group().lower() in NO_SON_NOMBRES):
        return None
    inicio = desplazamiento + nombre[0].start()
    return inicio, texto[inicio:desplazamiento + nombre[-1].end()]


def _letras(palabra):
    return all(c.isalpha() or c in "’'-" for c in palabra)


# --- Común ----------------------------------------------------------------------------

def _años(texto):
    """Cada año de una lista como «n.d., 2017a, 2017b», tal como está escrito."""
    return [" ".join(m.group().split()) for m in UN_AÑO.finditer(texto)]


def _cita(texto, inicio, fin, tipo, autores, año):
    autores = " ".join(autores.split())
    cita = Cita(ubicacion="", numero=None, zona="", tipo=tipo, inicio=inicio, fin=fin,
                fragmento=texto[inicio:fin], autores_escritos=autores, año=año)
    resto = autores
    abreviatura = ABREVIATURA.search(resto)
    if abreviatura:
        cita.abreviatura = abreviatura.group(1).strip()
        resto = (resto[:abreviatura.start()] + " " + resto[abreviatura.end():]).strip()
    et_al = ET_AL.search(resto)
    if et_al:
        cita.et_al, cita.et_al_escrito = True, et_al.group()
        resto = (resto[:et_al.start()] + resto[et_al.end():]).strip(" ,")
    conector = re.search(r"\s(y|e|and)\s|&", resto)
    if conector:
        cita.conector = conector.group(1) or "&"
    cita.autores = [a.strip(" ,") for a in SEPARA_AUTORES.split(resto) if a.strip(" ,")]
    return cita


# --- Opción --citas -------------------------------------------------------------------

def como_texto(citas, ancho=None):
    """Las citas en forma de tabla para la terminal, y un resumen al final."""
    ancho = ancho or shutil.get_terminal_size((100, 24)).columns - 1
    if not citas:
        return "No se ha encontrado ninguna cita en el texto."
    lugar = max(len("Párrafo"), *(len(_lugar(c)) for c in citas))
    autores = min(max(len("Autores"), *(len(c.autores_escritos) for c in citas)), 30)
    año = max(len("Año"), *(len(c.año) for c in citas))
    paginas = max(len("Páginas"), *(len(c.localizador or "") for c in citas))
    tipo = max(len(n) for n in NOMBRES_DE_TIPO.values())
    columnas = lugar + tipo + autores + año + paginas + 2 * 5
    texto = max(ancho - columnas, 20)

    lineas = [f"{'Párrafo':>{lugar}}  {'Tipo':<{tipo}}  {'Autores':<{autores}}  {'Año':<{año}}  "
              f"{'Páginas':<{paginas}}  Texto"]
    for c in citas:
        nombre = c.autores_escritos if len(c.autores_escritos) <= autores else c.autores_escritos[:autores - 1] + "…"
        fragmento = " ".join(c.fragmento.split())
        lineas.append(f"{_lugar(c):>{lugar}}  {NOMBRES_DE_TIPO[c.tipo]:<{tipo}}  {nombre:<{autores}}  "
                      f"{c.año:<{año}}  {c.localizador or '':<{paginas}}  "
                      f"{textwrap.shorten(fragmento, texto, placeholder='…')}")

    tipos = Counter(c.tipo for c in citas)
    obras = {(tuple(c.autores), c.año) for c in citas}
    partes = [f"{tipos[ENTRE_PARENTESIS]} entre paréntesis", _contar(tipos[NARRATIVA], "narrativa")]
    if tipos[SIN_PARENTESIS]:
        partes.append(f"{tipos[SIN_PARENTESIS]} sin paréntesis en tablas")
    lineas += ["", f"Resumen: {_contar(len(citas), 'cita')} ({', '.join(partes[:-1])} y {partes[-1]}) "
                   f"de {_contar(len(obras), 'obra distinta', 'obras distintas')}."]
    return "\n".join(lineas)


def _contar(n, singular, plural=None):
    return f"{n} {singular if n == 1 else plural or singular + 's'}"


def _lugar(c):
    if c.numero is not None:
        return f"{c.numero}" + (" (tabla)" if "tabla" in c.ubicacion else "")
    return c.ubicacion.replace("Nota al pie", "nota").replace("Nota al final", "nota final")
