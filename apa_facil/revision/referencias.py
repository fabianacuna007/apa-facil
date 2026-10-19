"""
Paso 7.2: lector de referencias, según la web oficial de APA (Elements of
reference list entries, Missing reference information, DOIs and URLs, y los
ejemplos de Journal article, Book/ebook y Webpage on a website references).

Solo lee: no cambia nada del documento. Una referencia de APA tiene cuatro
elementos, en este orden:

    Autor. (Fecha). Título. Fuente.

y, al final, el DOI o la URL. De cada entrada de la lista se guarda lo que está
escrito, **tal cual**: autores, año, título, fuente (revista o editorial),
volumen, número, páginas, lugar de edición (si lo hay, para el paso 7.4), DOI y
URL. No se juzga si está bien (eso es del 7.4) y **nunca se adivina**: lo que no
se encuentra se deja vacío, y una entrada que no se puede leer (por ejemplo, una
dirección web suelta) se marca como «no reconocida», con el motivo.

Cómo se separan las partes:
1. Se quita del final el DOI o la URL (y el «Recuperado de» de delante).
2. Se busca la fecha entre paréntesis: lo de delante es el autor, y lo de detrás,
   el título y la fuente. Si no la hay, el autor se reconoce por sus iniciales
   («Herranz, M.») y el año, si aparece suelto («Lefèvre, A. 2009.»).
3. El título y la fuente se separan por la cursiva, cuando la hay (APA pone en
   cursiva el título de las obras independientes y el nombre de la revista de un
   artículo), o si no, por el primer punto seguido de mayúscula.
4. La fuente de un artículo acaba en volumen, número y páginas: «34(2), 45-67».
"""

import re
import shutil
import textwrap
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from apa_facil import estructura
from apa_facil.documento import Estilos, w
from apa_facil.reglas.resumen import _activado

MAYUS = "A-ZÁÉÍÓÚÜÑÀÈÌÒÙÂÊÎÔÛÄËÏÖÇÅØÆŒŠŽ"

ALTA, MEDIA = "alta", "media"
ARTICULO, CAPITULO, OBRA = "articulo", "capitulo", "obra"
NOMBRES_DE_TIPO = {ARTICULO: "artículo de revista", CAPITULO: "capítulo de libro", OBRA: "obra independiente"}

# --- Patrones -------------------------------------------------------------------------

ENLACE = re.compile(r"(?:https?://|www\.)\S+|\bdoi:\s*\S+", re.I)
RECUPERADO = re.compile(r"(?:Recuperado(?:\s+el\s+[^,]{1,40}?,?)?\s+de|Retrieved(?:\s+[^,]{1,40}?,)?\s+from"
                        r"|Disponible\s+en|Consultado\s+en|Accessed\s+from)\s*:?\s*$", re.I)
AÑO = r"(?:1[5-9]|20)\d\d[a-z]?"
SIN_FECHA = r"s\.\s?f\.?|n\.\s?d\.?|en\s+prensa|in\s+press"
# La fecha entre paréntesis: «(2015)», «(2020, August 26)», «(2018, julio)», «(s. f.)», «(ca. 1999)».
FECHA = re.compile(rf"\((?P<fecha>(?:ca\.\s*)?{AÑO}(?:\s*,\s*[^()]{{1,30}})?|{SIN_FECHA})\)", re.I)
# El año suelto, como una parte de la referencia: «Lefèvre, A. 2009. Industrialisation…».
AÑO_SUELTO = re.compile(rf"(?:^|(?<=[.,]\s))(?P<año>{AÑO}|{SIN_FECHA})(?=\s*[.,]?\s|\s*[.,]?$)", re.I)
UN_AÑO = re.compile(rf"{AÑO}|{SIN_FECHA}", re.I)

# Un autor: «Ramírez Soto, M.», «Velasco Rodríguez, M. L.», «Santos-García, J.-P.».
INICIALES = rf"(?:[{MAYUS}]\.(?:\s?-\s?[{MAYUS}]\.)?\s?)+"
UN_AUTOR = re.compile(rf"\s*(?P<apellidos>[^,;&()\d.]+?),\s*(?P<iniciales>{INICIALES})")
SEPARADOR = re.compile(r"\s*(?:,\s*(?:\.\s?\.\s?\.\s*)?(?:(?:&|y|e|and)\s+)?|;\s*|(?:&|y|e|and)\s+)")
ROL = re.compile(r"\s*\((?P<rol>(?:Eds?|Coords?|Comps?|Dirs?|Trads?)\.)\)\s*\.?\s*$", re.I)
ROL_SUELTO = re.compile(r"\s*\((?P<rol>(?:Eds?|Coords?|Comps?|Dirs?|Trads?)\.)\)\s*\.?", re.I)

# Detrás del título: la edición, el volumen o el número de un informe: «(2nd ed.)», «(2.ª ed., Vol. 1)».
EDICION = re.compile(r"\s*\((?P<edicion>[^()]*(?:\bed\.|\bVol\.|\bvol\.|\bNo\.|\bnúm\.)[^()]*)\)\s*$")
# Abreviaturas que acaban en punto y no cierran el título.
ABREVIATURAS = {"ed.", "eds.", "vol.", "vols.", "pp.", "p.", "núm.", "no.", "n.º", "trad.", "coord.", "comp.",
                "dir.", "cap.", "art.", "ca.", "et", "al.", "st.", "sr.", "dr.", "dra."}
NUMERADOS = {"vol.", "vols.", "no.", "núm.", "n.º", "cap.", "parte", "part", "tomo"}

# El final de la fuente de un artículo: «34(2), 45-67», «8, 15-29», «13(3), Article e0193972», «(249), 30–33».
PAGINAS = r"(?:pp?\.\s*)?[A-Za-z]?\d+(?:\s*[-–—]\s*[A-Za-z]?\d+)?"
FINAL_DE_REVISTA = re.compile(rf"""
    ,\s*(?:
        (?P<volumen>\d+)\s*(?:\((?P<numero>[^()]+)\))?
      | \((?P<solo_numero>[^()]+)\)
      | (?:nº|n\.º|n°|núm\.|no\.)\s*(?P<numero_escrito>\d+)
    )
    (?:\s*,\s*(?:(?P<articulo>(?:Art[ií]culo|Article)\s+\S+?)|(?P<paginas>{PAGINAS})))?
    \s*\.?\s*$""", re.X | re.I)
CAPITULO_EN = re.compile(r"(?:^|(?<=[.,]\s))(?:En|In)\s+(?=\S)")
PAGINAS_DEL_CAPITULO = re.compile(r"\((?P<paginas>pp?\.\s*[^()]+)\)")
LUGAR_Y_EDITORIAL = re.compile(rf"^(?P<lugar>[{MAYUS}][^:,.;]{{1,40}}):\s*(?P<editorial>\S.*)$")

# Ciudades de edición frecuentes (APA 7 ya no pone el lugar: el paso 7.4 lo señalará).
CIUDADES = {
    "madrid", "barcelona", "valencia", "sevilla", "salamanca", "granada", "bilbao", "zaragoza", "málaga", "murcia",
    "oviedo", "santiago de compostela", "pamplona", "valladolid", "alicante", "córdoba", "cádiz", "la coruña",
    "a coruña", "santander", "méxico", "ciudad de méxico", "méxico d. f.", "méxico, d. f.", "buenos aires", "bogotá",
    "lima", "santiago", "santiago de chile", "caracas", "quito", "montevideo", "la habana", "san josé", "guadalajara",
    "medellín", "asunción", "la paz", "parís", "paris", "londres", "london", "nueva york", "new york", "roma",
    "lisboa", "berlín", "berlin", "oxford", "cambridge", "washington", "washington, dc", "washington, d.c.",
    "chicago", "boston", "ámsterdam", "amsterdam", "bruselas", "ginebra", "milán", "florencia", "múnich",
}


@dataclass
class Autor:
    """Un autor, tal como está escrito. Las organizaciones no tienen iniciales."""
    apellidos: str              # «Ramírez Soto», o el nombre de la organización
    iniciales: Optional[str]    # «M.», «M. L.»; None en una organización


@dataclass
class Referencia:
    """Una entrada de la lista de referencias, con cada parte tal como está escrita."""
    ubicacion: str                  # «Párrafo 20»
    numero: int                     # el número del párrafo (el del diagnóstico)
    texto: str                      # la entrada entera
    lista: int = 1                  # 1, 2… si hay varias listas (Bibliografía y Webgrafía)
    rotulo: Optional[str] = None    # el rótulo de su lista, tal cual: «BIBLIOGRAFÍA», «Webgrafía:»
    reconocida: bool = True
    motivo: Optional[str] = None    # por qué no se ha podido leer (en las no reconocidas)
    confianza: str = ALTA           # alta: las partes en el orden de APA; media: algo fuera de su sitio
    dudas: list = field(default_factory=list)   # qué se ha encontrado fuera de su sitio, en español
    tipo: Optional[str] = None      # ARTICULO, CAPITULO u OBRA (libro, informe, página web…)

    autores_escritos: Optional[str] = None   # «Ortega, L. y Blanco, R.»
    autores: list = field(default_factory=list)   # [Autor]
    grupo: bool = False             # autor sin iniciales: una organización (o el título de una obra sin autor)
    conector: Optional[str] = None  # «y», «&», «;» entre autores
    rol: Optional[str] = None       # «Eds.», «Coord.»…

    fecha_escrita: Optional[str] = None   # «2015», «2020, August 26», «s. f.»
    año: Optional[str] = None       # «2015», «2017a», «s. f.», «n.d.», «en prensa»
    año_entre_parentesis: Optional[bool] = None
    año_tras_el_autor: Optional[bool] = None

    titulo: Optional[str] = None
    titulo_en_cursiva: Optional[bool] = None   # None si solo una parte va en cursiva
    edicion: Optional[str] = None   # «2nd ed.», «2.ª ed., Vol. 1»

    fuente: Optional[str] = None    # la revista, la editorial o el sitio web (sin el lugar)
    fuente_en_cursiva: Optional[bool] = None
    volumen: Optional[str] = None
    numero_de_revista: Optional[str] = None   # el número (issue): «2» en «34(2)», «7» en «nº 7»
    paginas: Optional[str] = None   # «45-67», «pp. 9–43»
    articulo: Optional[str] = None  # «Article e0193972»
    lugar: Optional[str] = None     # la ciudad de edición, si la hay: «Madrid», «Paris»

    doi: Optional[str] = None       # tal como está: «https://doi.org/10…», «doi:10.0000/…»
    url: Optional[str] = None
    recuperado_de: Optional[str] = None   # «Recuperado de», «Retrieved from»…


class _Texto:
    """Un texto y, para cada letra, si se ve en cursiva. Se recorta sin perder la cursiva."""

    def __init__(self, texto, cursiva=None):
        self.texto = texto
        self.cursiva = list(cursiva) if cursiva is not None else [False] * len(texto)

    def __getitem__(self, corte):
        return _Texto(self.texto[corte], self.cursiva[corte])

    def __add__(self, otro):
        return _Texto(self.texto + otro.texto, self.cursiva + otro.cursiva)

    def __str__(self):
        return self.texto

    def recortado(self, quitar=" "):
        inicio, fin = 0, len(self.texto)
        while inicio < fin and self.texto[inicio] in quitar:
            inicio += 1
        while fin > inicio and self.texto[fin - 1] in quitar:
            fin -= 1
        return self[inicio:fin]

    def en_cursiva(self):
        """True si todas sus letras van en cursiva, False si ninguna, None si solo algunas."""
        letras = [c for ch, c in zip(self.texto, self.cursiva) if ch.isalnum()]
        if not letras:
            return None
        return True if all(letras) else False if not any(letras) else None


# --- Lectura del documento ------------------------------------------------------------

def leer(doc, clasificacion=None):
    """
    Las referencias de un documento abierto con python-docx, en orden: los
    párrafos `referencia` de la zona de referencias del clasificador, con el
    rótulo de su lista. `clasificacion` es la de `estructura.clasificar`; si no
    se da, se clasifica el documento tal como está.
    """
    if clasificacion is None:
        clasificacion = estructura.clasificar(doc)
    estilos = Estilos(doc)
    referencias, lista, rotulo = [], 0, None
    for c in clasificacion:
        if c.tipo == "titulo_referencias":
            lista, rotulo = lista + 1, " ".join(c.texto.split())
        elif c.tipo == "referencia" and c.zona == "referencias":
            texto, cursiva = _texto_con_cursiva(estilos, c.elemento)
            referencia = leer_texto(texto, cursiva, f"Párrafo {c.numero}", c.numero)
            referencia.lista, referencia.rotulo = max(lista, 1), rotulo
            referencias.append(referencia)
    return referencias


def _texto_con_cursiva(estilos, p):
    """El texto del párrafo (como `documento.texto_parrafo`) y si cada letra se ve en cursiva."""
    texto, cursiva, vistos = [], [], {}
    for el in p.iter(w("t"), w("tab"), w("br"), w("cr"), w("noBreakHyphen")):
        if next(el.iterancestors(w("p")), None) is not p:
            continue   # pertenece a un párrafo de dentro (cuadro de texto)
        r = next(el.iterancestors(w("r")), None)
        nombre = etree_nombre(el)
        if nombre == "t":
            trozo = el.text or ""
        elif nombre == "tab" and r is not None and el.getparent() is r:
            trozo = "\t"
        elif nombre in ("br", "cr"):
            trozo = "\n"
        elif nombre == "noBreakHyphen":
            trozo = "-"
        else:
            continue
        if r is not None and r not in vistos:
            vistos[r] = _activado(estilos, r, p, "i")
        texto.append(trozo)
        cursiva += [vistos.get(r, False)] * len(trozo)
    return "".join(texto), cursiva


def etree_nombre(el):
    return el.tag.rsplit("}", 1)[-1]


# --- Lectura de una entrada -----------------------------------------------------------

def leer_texto(texto, cursiva=None, ubicacion="", numero=0):
    """Lee una entrada de la lista de referencias. `cursiva`: para cada letra, si va en cursiva."""
    ref = Referencia(ubicacion=ubicacion, numero=numero, texto=texto)
    cuerpo = _Texto(texto, cursiva).recortado(" \t\n")

    # 1. El DOI o la URL del final, y el «Recuperado de» de delante.
    enlaces = list(ENLACE.finditer(cuerpo.texto))
    if enlaces and not cuerpo.texto[enlaces[-1].end():].strip(" .;"):
        enlace = enlaces[-1]
        valor = enlace.group().rstrip(".,;")
        if "doi.org/" in valor.lower() or valor.lower().startswith("doi:"):
            ref.doi = valor
        else:
            ref.url = valor
        cuerpo = cuerpo[:enlace.start()].recortado(" \t\n")
        recuperado = RECUPERADO.search(cuerpo.texto)
        if recuperado:
            ref.recuperado_de = " ".join(recuperado.group().rstrip(" :").split())
            cuerpo = cuerpo[:recuperado.start()].recortado(" \t\n")
    if not any(ch.isalpha() for ch in cuerpo.texto):
        return _no_reconocida(ref, "Solo tiene una dirección web: faltan el autor, la fecha y el título."
                              if ref.url or ref.doi else "No tiene texto que se pueda leer.")

    # 2. La fecha y el autor.
    fecha = FECHA.search(cuerpo.texto)
    if fecha:
        ref.fecha_escrita = " ".join(fecha.group("fecha").split())
        ref.año = _año(ref.fecha_escrita)
        ref.año_entre_parentesis = True
        autor = cuerpo[:fecha.start()].recortado(" \t\n")
        resto = cuerpo[fecha.end():]
        if not _leer_autores(ref, autor.texto, completo=True):
            _grupo(ref, autor.texto)
        ref.año_tras_el_autor = True
    else:
        fin = _leer_autores(ref, cuerpo.texto, completo=False)
        if fin is None:
            return _no_reconocida(ref, "No se distingue el autor del título: no tiene la fecha entre paréntesis "
                                       "ni un autor con sus iniciales («Apellido, A.»).")
        resto = cuerpo[fin:]
        suelto = AÑO_SUELTO.search(resto.texto)
        if suelto:
            ref.fecha_escrita = ref.año = " ".join(suelto.group("año").split())
            ref.año_entre_parentesis = False
            ref.año_tras_el_autor = not resto.texto[:suelto.start()].strip(" .,")
            resto = resto[:suelto.start()] + resto[suelto.end():]
            ref.dudas.append("el año no va entre paréntesis" + ("" if ref.año_tras_el_autor
                                                                else " ni detrás del autor"))
        else:
            ref.dudas.append("no tiene fecha")

    # 3. El título y la fuente.
    resto = resto.recortado(" .\t\n")
    if not resto.texto:
        ref.dudas.append("no tiene título")
    else:
        _titulo_y_fuente(ref, resto)
    ref.confianza = MEDIA if ref.dudas else ALTA
    return ref


def _no_reconocida(ref, motivo):
    ref.reconocida, ref.motivo, ref.confianza = False, motivo, None
    return ref


def _año(fecha):
    m = UN_AÑO.search(fecha)
    return " ".join(m.group().split()) if m else None


def _leer_autores(ref, texto, completo):
    """
    Lee al principio de `texto` los autores con iniciales («Castro, P., Llorente,
    M. y Benítez, J.»). Con `completo`, tienen que ocupar todo el texto (salvo el
    papel, «(Eds.)», y el punto final). Devuelve dónde acaban, o None si no hay.
    """
    rol = ROL.search(texto) if completo else None
    zona = texto[:rol.start()] if rol else texto
    autores, posicion, conectores = [], 0, []
    while True:
        m = UN_AUTOR.match(zona, posicion)
        if not m:
            break
        autores.append(Autor(" ".join(m.group("apellidos").split()), m.group("iniciales").strip()))
        posicion = m.end()
        separador = SEPARADOR.match(zona, posicion)
        if not separador or not separador.group().strip() or not UN_AUTOR.match(zona, separador.end()):
            break
        conectores.append(separador.group().strip(" ,.") or ",")
        posicion = separador.end()
    if not autores:
        return None
    if completo and zona[posicion:].strip(" ."):
        return None
    escritos = texto[:posicion]
    if not completo:   # sin fecha: el papel puede ir justo detrás de los autores
        rol = ROL_SUELTO.match(texto, posicion)
        posicion = rol.end() if rol else posicion
    ref.autores = autores
    ref.autores_escritos = " ".join(escritos.split()).rstrip(",")
    ref.rol = rol.group("rol") if rol else None
    ref.conector = next((c for c in conectores if c in ("y", "&", "e", "and", ";")), None)
    return posicion


def _grupo(ref, texto):
    """El autor sin iniciales: una organización (o el título, en una obra sin autor)."""
    rol = ROL.search(texto)
    if rol:
        ref.rol, texto = rol.group("rol"), texto[:rol.start()]
    nombre = " ".join(texto.split()).rstrip(" .")
    if nombre:
        ref.autores, ref.autores_escritos, ref.grupo = [Autor(nombre, None)], nombre, True
    else:
        ref.dudas.append("no tiene autor")


def _titulo_y_fuente(ref, resto):
    capitulo = CAPITULO_EN.search(resto.texto)
    cursivas = [i for i, (ch, c) in enumerate(zip(resto.texto, resto.cursiva)) if c and ch.isalnum()]
    if capitulo and capitulo.start() > 0:
        titulo, fuente = resto[:capitulo.start()], resto[capitulo.start():]
    elif cursivas and resto.cursiva[0]:
        # El título en cursiva (obra independiente): hasta el primer punto después de la cursiva.
        fin = _fin_de_frase(resto.texto, desde=_fin_de_la_cursiva(resto))
        titulo, fuente = (resto[:fin], resto[fin:]) if fin is not None else (resto, None)
    elif cursivas:
        # La cursiva empieza más tarde: es la revista (o la editorial); el título va delante.
        inicio = cursivas[0]
        while inicio > 0 and resto.cursiva[inicio - 1] and not resto.texto[inicio - 1].isspace():
            inicio -= 1
        titulo, fuente = resto[:inicio], resto[inicio:]
    else:
        fin = _fin_de_frase(resto.texto)
        titulo, fuente = (resto[:fin], resto[fin:]) if fin is not None else (resto, None)
    if fuente is None or not fuente.texto.strip(" ."):
        # Sin punto entre el título y la revista: «Autoestima digital, Revista Ejemplo
        # Joven, nº 7». Si acaba como una revista, se separan por la coma de delante.
        texto = titulo.texto.rstrip(" .")
        revista = FINAL_DE_REVISTA.search(texto)
        if revista and "," in texto[:revista.start()]:
            coma = texto.rindex(",", 0, revista.start())
            titulo, fuente = titulo[:coma], titulo[coma + 1:]
            ref.dudas.append("el título y la revista solo están separados por una coma")

    titulo = titulo.recortado(" .\t\n")
    edicion = EDICION.search(titulo.texto)
    if edicion:
        ref.edicion = " ".join(edicion.group("edicion").split())
        titulo = titulo[:edicion.start()].recortado(" .")
    ref.titulo = " ".join(titulo.texto.split()) or None
    ref.titulo_en_cursiva = titulo.en_cursiva()
    fuente = fuente.recortado(" .\t\n") if fuente is not None else None
    _fuente(ref, fuente, capitulo is not None and capitulo.start() > 0)


def _fin_de_la_cursiva(texto):
    i = 0
    while i < len(texto.texto) and (texto.cursiva[i] or not texto.texto[i].isalnum()):
        i += 1
    while i > 0 and not texto.cursiva[i - 1] and not texto.texto[i - 1].isalnum():
        i -= 1   # vuelve hasta la última letra en cursiva (sin contar «. » ni espacios)
    return i


def _fin_de_frase(texto, desde=0):
    """
    Dónde acaba el título: el primer punto (o «?», «!») seguido de espacio y de
    otra palabra, sin contar las abreviaturas («ed.», «Vol.») ni las iniciales
    («In A. Pérez»). Si el título abre con «¿», su «?» solo cuenta si no hay otro
    punto antes del final (en español, «¿Lugares que no importan? La despoblación…»).
    None si no hay.
    """
    candidatos = []
    for m in re.finditer(r"[.?!:]?[.?!](?=\s+\S)", texto[desde:]):
        posicion = desde + m.end()
        signo = texto[posicion - 1]
        palabras = texto[:posicion].lower().split()
        palabra = palabras[-1] if palabras else ""
        anterior = palabras[-2] if len(palabras) > 1 else ""
        if signo == "." and (palabra in ABREVIATURAS or re.fullmatch(rf"[{MAYUS.lower()}a-z]\.", palabra)
                             or re.fullmatch(r"\d+\.", palabra) and anterior in NUMERADOS):
            continue   # «Vol. 1. Building…»: ni «Vol.» ni «1.» cierran el título
        candidatos.append((posicion, signo))
    puntos = [p for p, s in candidatos if s == "."]
    for posicion, signo in candidatos:
        if signo in "?!" and ("¿" if signo == "?" else "¡") in texto[:posicion] and puntos:
            continue
        return posicion
    return None


def _fuente(ref, fuente, es_capitulo):
    if fuente is None or not fuente.texto:
        ref.tipo = OBRA
        return
    ref.fuente_en_cursiva = fuente.en_cursiva()
    texto = " ".join(fuente.texto.split())
    if es_capitulo:
        ref.tipo, ref.fuente = CAPITULO, texto
        paginas = PAGINAS_DEL_CAPITULO.search(texto)
        ref.paginas = " ".join(paginas.group("paginas").split()) if paginas else None
        return
    revista = FINAL_DE_REVISTA.search(fuente.texto)
    if revista and revista.start() > 0:
        ref.tipo = ARTICULO
        ref.volumen = revista.group("volumen")
        ref.numero_de_revista = (revista.group("numero") or revista.group("solo_numero")
                                 or revista.group("numero_escrito"))
        ref.paginas, ref.articulo = revista.group("paginas"), revista.group("articulo")
        nombre = fuente[:revista.start()].recortado(" \t\n")
        ref.fuente = " ".join(nombre.texto.split()) or None
        ref.fuente_en_cursiva = nombre.en_cursiva()   # el nombre de la revista (sin volumen ni páginas)
        return
    ref.tipo = OBRA
    lugar = LUGAR_Y_EDITORIAL.match(texto)
    if lugar and len(lugar.group("lugar").split()) <= 4:
        ref.lugar, ref.fuente = lugar.group("lugar").strip(), lugar.group("editorial").strip() or None
    elif texto.lower() in CIUDADES:
        ref.lugar, ref.fuente = texto, None
    elif "," in texto and texto.rsplit(",", 1)[1].strip().lower() in CIUDADES:
        editorial, ciudad = texto.rsplit(",", 1)
        ref.lugar, ref.fuente = ciudad.strip(), editorial.strip()
    else:
        ref.fuente = texto


# --- Opción --referencias -------------------------------------------------------------

def como_texto(referencias, ancho=None):
    """Cada referencia en un bloque con sus partes, y un resumen al final."""
    ancho = ancho or shutil.get_terminal_size((100, 24)).columns - 1
    if not referencias:
        return "No se ha encontrado ninguna lista de referencias."
    numero = max(2, len(str(referencias[-1].numero)))
    sangria = " " * (numero + 2)
    lineas, lista = [], None
    for r in referencias:
        if r.lista != lista:
            lista = r.lista
            rotulo = f"── Lista {lista}" + (f" («{r.rotulo}»)" if r.rotulo else "") + " "
            lineas.append(rotulo + "─" * max(ancho - len(rotulo), 3))
        if not r.reconocida:
            lineas.append(f"{r.numero:>{numero}}  No reconocida. {r.motivo}")
            lineas += _campo(sangria, "Texto", " ".join(r.texto.split()), ancho)
            continue
        lineas.append(f"{r.numero:>{numero}}  {NOMBRES_DE_TIPO[r.tipo].capitalize()} · confianza {r.confianza}")
        autores = r.autores_escritos and (r.autores_escritos + (f" ({r.rol})" if r.rol else "")
                                          + (" (organización o título)" if r.grupo else ""))
        año = r.fecha_escrita and (r.fecha_escrita + ("" if r.año_entre_parentesis else " (sin paréntesis)"))
        titulo = r.titulo and (r.titulo + (f" ({r.edicion})" if r.edicion else "")
                               + (" (en cursiva)" if r.titulo_en_cursiva else ""))
        fuente = r.fuente and (r.fuente + (" (en cursiva)" if r.fuente_en_cursiva else ""))
        revista = ", ".join(parte for parte in (
            r.volumen and f"volumen {r.volumen}", r.numero_de_revista and f"número {r.numero_de_revista}",
            r.paginas and f"páginas {r.paginas}", r.articulo) if parte)
        for nombre, valor in (("Autores", autores), ("Año", año), ("Título", titulo), ("Fuente", fuente),
                              ("Revista", revista), ("Lugar", r.lugar), ("DOI", r.doi),
                              ("URL", r.url), ("Antes del enlace", r.recuperado_de)):
            if valor:
                lineas += _campo(sangria, nombre, valor, ancho)
        if r.dudas:
            lineas += _campo(sangria, "Fuera de su sitio", "; ".join(r.dudas), ancho)

    reconocidas = [r for r in referencias if r.reconocida]
    confianzas = Counter(r.confianza for r in reconocidas)
    lineas += ["", f"Resumen: {len(referencias)} referencia(s): {len(reconocidas)} reconocida(s) "
                   f"(confianza alta {confianzas[ALTA]}, media {confianzas[MEDIA]}) y "
                   f"{len(referencias) - len(reconocidas)} no reconocida(s)."]
    return "\n".join(lineas)


def _campo(sangria, nombre, valor, ancho):
    return textwrap.wrap(f"{nombre}: {valor}", width=max(ancho, 40), initial_indent=sangria,
                         subsequent_indent=sangria + "  ")
