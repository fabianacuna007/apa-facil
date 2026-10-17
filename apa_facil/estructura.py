"""
Clasificador de párrafos (paso 3.1): decide qué es cada párrafo del cuerpo
(portada, título, cuerpo, lista, referencia…) y con qué confianza, aunque el
estudiante no haya usado los estilos de Word.

Funciona en tres pasos:
  1. Rasgos: lo que se puede medir en cada párrafo (texto, estilo de Word,
     negrita, tamaño, sangrías, imágenes, saltos, campos de índice…) y qué
     tiene al lado (una tabla, una imagen).
  2. Portada: los párrafos del principio, hasta el primer salto, el primer
     título o el primer párrafo largo.
  3. Recorrido en orden: cada párrafo recibe su tipo y la zona del trabajo
     en la que está (portada, índice, resumen, cuerpo, referencias, anexos).

La confianza dice de dónde sale la decisión:
  - alta: una señal explícita de Word (estilo, lista, campo, imagen) o un
    rótulo exacto ("Referencias", "Anexo I…");
  - media: se deduce del formato o de la posición (negrita corta, viñeta
    escrita a mano, párrafo pegado a una tabla);
  - baja: hay señales contradictorias o se decide por descarte.

Solo se clasifican los párrafos del cuerpo: no los de dentro de tablas ni de
cuadros de texto. El clasificador no cambia nada del documento.
"""

import re
from dataclasses import dataclass
from typing import Optional

from apa_facil.documento import MC, WP, Estilos, texto_parrafo, w
from apa_facil.reglas.alineacion import _salto_antes, _salto_en

TIPOS = ("portada", "indice", "titulo", "cuerpo", "lista", "leyenda_tabla", "leyenda_figura",
         "cita_bloque", "titulo_referencias", "referencia", "anexo", "imagen", "vacio")
ZONAS = ("portada", "indice", "resumen", "cuerpo", "referencias", "anexos")
ALTA, MEDIA, BAJA = "alta", "media", "baja"

PALABRAS_DE_TITULO = 15        # un título casi nunca llega a 15 palabras
PALABRAS_DE_PARRAFO_LARGO = 25  # un renglón de portada casi nunca llega a 20
PALABRAS_DE_CITA_EN_BLOQUE = 40  # APA: 40 palabras o más
NIVEL_MAXIMO = 5

ROTULOS_DE_REFERENCIAS = {"referencias", "referencia", "referencias bibliográficas",
                          "referencias bibliograficas", "bibliografía", "bibliografia", "webgrafía",
                          "webgrafia", "fuentes", "fuentes consultadas", "bibliografía y webgrafía",
                          "bibliografia y webgrafia", "obras citadas", "lista de referencias",
                          "references", "bibliography"}
ROTULOS_DE_RESUMEN = {"resumen", "abstract"}
ROTULOS_DE_INDICE = {"índice", "indice", "índice general", "indice general", "contenido", "contenidos",
                     "tabla de contenido", "tabla de contenidos", "sumario", "table of contents"}
ANEXO = re.compile(r"^(anexo|anexos|apéndice|apendice|apéndices|apendices)\b", re.IGNORECASE)
LISTA_A_MANO = re.compile(r"^\s*([•·▪◦‣●○■□\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z]\))\s")
LEYENDA_DE_TABLA = re.compile(r"^(tabla|cuadro)\s+[A-Z]?\d", re.IGNORECASE)
LEYENDA_DE_FIGURA = re.compile(r"^(figura|gráfico|grafico|imagen|ilustración|ilustracion|mapa|"
                               r"fotografía|fotografia|diagrama|esquema)\s+[A-Z]?\d", re.IGNORECASE)
NOTA = re.compile(r"^(nota\s*\.|fuente\s*:|fuentes\s*:)", re.IGNORECASE)
FINAL_DE_FRASE = re.compile(r"[.…:][\"'»”)]*$")
NUMERACION = re.compile(r"^(\d+(?:\.\d+)*)\.?\s")   # "1.", "1.1.", "2.3.1"
URL = re.compile(r"https?://|www\.|doi\.org|\bdoi:", re.IGNORECASE)
AUTOR_Y_FECHA = re.compile(r"^[^\W\d][^,(]{0,80},.*(\(\s*(\d{4}|s\.\s?f\.)|\b\d{4}\b)", re.IGNORECASE)
CON_FECHA = re.compile(r"\(\s*(\d{4}[a-z]?|s\.\s?f\.)\s*[),]|\b(1[5-9]|20)\d{2}\b")
ESTILO_DE_TITULO_NUMERADO = re.compile(r"^(heading|título|titulo)\s*(\d)$")
ESTILO_DE_INDICE = re.compile(r"^(toc|tdc)\s*\d$|^(toc heading|título de tdc|titulo de tdc)$")
ESTILOS_DE_LEYENDA = ("caption", "descripción", "descripcion", "epígrafe", "epigrafe")
ESTILOS_DE_TITULO_DEL_TRABAJO = ("title", "subtitle", "título", "titulo", "subtítulo", "subtitulo")


@dataclass
class Parrafo:
    """Un párrafo clasificado."""
    posicion: int                 # 0, 1, 2… en el orden del documento
    elemento: object              # el w:p del XML
    texto: str                    # en una sola línea, sin espacios repetidos
    tipo: str                     # uno de TIPOS
    confianza: str                # alta, media o baja
    motivo: str                   # por qué, en español
    zona: str                     # uno de ZONAS
    nivel: Optional[int] = None   # solo en los títulos: de 1 a 5


# --- 1. Rasgos ------------------------------------------------------------------------

@dataclass
class _Rasgos:
    texto: str
    palabras: int
    estilo: str              # nombre del estilo del párrafo, tal como lo muestra Word
    estilos: list            # nombres de toda la cadena de estilos, en minúsculas
    nivel_de_word: Optional[int]
    lista_de_word: bool
    negrita: bool
    mayusculas: bool
    tamano: Optional[float]  # el más frecuente, en puntos
    alineacion: str
    izquierda: int
    francesa: int
    imagen: bool
    salto_antes: bool
    salto_despues: bool
    de_indice: bool          # campo o control de índice de Word, o estilo de índice
    antes: Optional[str] = None    # lo que hay justo antes: "tabla", "imagen" o "texto"
    despues: Optional[str] = None


def _capa(estilos, p, etiqueta):
    """El primer elemento `etiqueta` en las capas de formato del párrafo (directo, estilos, por defecto)."""
    for capa in estilos._capas_parrafo(p):
        el = capa.find(w(etiqueta)) if capa is not None else None
        if el is not None:
            return el
    return None


def _entero(valor):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def _nivel_de_word(estilos, p, nombres):
    """El nivel de título que Word le da al párrafo (nivel de esquema o estilo "Título N"), o None."""
    nivel = _capa(estilos, p, "outlineLvl")
    if nivel is not None and nivel.get(w("val")) not in (None, "9"):
        return min(_entero(nivel.get(w("val"))) + 1, NIVEL_MAXIMO)
    for nombre in nombres:
        encontrado = ESTILO_DE_TITULO_NUMERADO.match(nombre)
        if encontrado:
            return min(int(encontrado.group(2)), NIVEL_MAXIMO)
    return None


def _es_lista_de_word(estilos, p):
    numeracion = _capa(estilos, p, "numPr")
    if numeracion is None:
        return False
    num_id = numeracion.find(w("numId"))
    return num_id is None or num_id.get(w("val")) != "0"


def _fragmentos_con_texto(p):
    return [r for r in p.iter(w("r"))
            if next(r.iterancestors(w("p")), None) is p
            and "".join(t.text or "" for t in r.iter(w("t"))).strip()]


def _todo_en_negrita(estilos, p, fragmentos):
    if not fragmentos:
        return False
    for r in fragmentos:
        negrita = next((capa.find(w("b")) for capa in estilos._capas_texto(r, p)
                        if capa is not None and capa.find(w("b")) is not None), None)
        if negrita is None or negrita.get(w("val")) in ("0", "false", "off"):
            return False
    return True


def _tamano(estilos, p, fragmentos):
    """El tamaño de letra que más se ve en el párrafo (contando letras), o None si no tiene texto."""
    letras = {}
    for r in fragmentos:
        largo = len("".join(t.text or "" for t in r.iter(w("t"))))
        tamano = estilos.tamano(r, p)
        letras[tamano] = letras.get(tamano, 0) + largo
    return max(letras, key=letras.get) if letras else None


def _tiene_imagen(p):
    for el in p.iter(f"{{{WP}}}inline", f"{{{WP}}}anchor", w("pict")):
        if next(el.iterancestors(f"{{{MC}}}Fallback"), None) is not None:
            continue   # copia para versiones antiguas de Word
        if next(el.iterancestors(w("p")), None) is p:
            return True
    return False


def _parrafos_de_indice(parrafos):
    """
    Los párrafos que forman parte de un índice de Word: los de dentro de un
    campo TOC (que puede ocupar varios párrafos) o de un control de contenido
    de índice.
    """
    resultado, pila = set(), []   # pila: el código de cada campo abierto ("TOC", "PAGEREF"…)
    for p in parrafos:
        dentro = "TOC" in pila
        for el in p.iter(w("fldChar"), w("instrText"), w("fldSimple")):
            if el.tag == w("fldChar"):
                tipo = el.get(w("fldCharType"))
                if tipo == "begin":
                    pila.append("")
                elif tipo == "end" and pila:
                    pila.pop()
            elif el.tag == w("instrText"):
                palabras = (el.text or "").split()
                if pila and not pila[-1] and palabras:   # el código puede venir partido en trozos
                    pila[-1] = palabras[0].upper()
            elif (el.get(w("instr")) or "").split()[:1] == ["TOC"]:
                dentro = True
            dentro = dentro or "TOC" in pila
        for sdt in p.iterancestors(w("sdt")):
            galeria = sdt.find(f".//{w('docPartGallery')}")
            alias = sdt.find(f"{w('sdtPr')}/{w('alias')}")
            if (galeria is not None and "contents" in (galeria.get(w("val")) or "").lower()) or \
               (alias is not None and _normalizar(alias.get(w("val")) or "") in ROTULOS_DE_INDICE):
                dentro = True
        if dentro:
            resultado.add(p)
    return resultado


def _bloques(doc):
    """Los párrafos del cuerpo y las tablas de primer nivel, en orden."""
    bloques = []
    for el in doc.element.body.iter(w("p"), w("tbl")):
        if next(el.iterancestors(w("tc"), w("txbxContent")), None) is not None:
            continue
        bloques.append(el)
    return bloques


def _rasgos(estilos, p, indice):
    texto = " ".join(texto_parrafo(p).split())
    nombres = [estilos.nombre(e.get(w("styleId"))).lower() for e in estilos.cadena(estilos.estilo_de(p))]
    fragmentos = _fragmentos_con_texto(p)
    letras = [c for c in texto if c.isalpha()]
    return _Rasgos(
        texto=texto,
        palabras=len(texto.split()),
        estilo=estilos.nombre(estilos.estilo_de(p)),
        estilos=nombres,
        nivel_de_word=_nivel_de_word(estilos, p, nombres),
        lista_de_word=_es_lista_de_word(estilos, p),
        negrita=_todo_en_negrita(estilos, p, fragmentos),
        mayusculas=bool(letras) and all(c.isupper() for c in letras),
        tamano=_tamano(estilos, p, fragmentos),
        alineacion=estilos.parrafo(p, "jc", "val") or "left",
        izquierda=max(_entero(estilos.parrafo(p, "ind", "left")), _entero(estilos.parrafo(p, "ind", "start"))),
        francesa=_entero(estilos.parrafo(p, "ind", "hanging")),
        imagen=_tiene_imagen(p),
        salto_antes=_salto_antes(estilos, p),
        salto_despues=_salto_en(p),
        de_indice=p in indice or any(ESTILO_DE_INDICE.match(n) for n in nombres),
    )


def _vecinos(bloques, rasgos):
    """Anota en cada párrafo qué tiene justo antes y justo después, sin contar los párrafos vacíos."""
    def que_es(el):
        if el.tag == w("tbl"):
            return "tabla"
        r = rasgos[el]
        if r.imagen and not r.texto:
            return "imagen"
        return "texto" if r.texto else None

    utiles = [(el, que_es(el)) for el in bloques]
    utiles = [(el, cosa) for el, cosa in utiles if cosa]
    for i, (el, _) in enumerate(utiles):
        if el.tag == w("p"):
            rasgos[el].antes = utiles[i - 1][1] if i > 0 else None
            rasgos[el].despues = utiles[i + 1][1] if i + 1 < len(utiles) else None


# --- 2. Portada -----------------------------------------------------------------------

def _normalizar(texto):
    return " ".join(texto.split()).rstrip(":.").lower()


def _es_rotulo(r):
    n = _normalizar(r.texto)
    return (n in ROTULOS_DE_REFERENCIAS or n in ROTULOS_DE_RESUMEN or n in ROTULOS_DE_INDICE
            or bool(ANEXO.match(r.texto) and r.palabras < PALABRAS_DE_TITULO))


def _portada(parrafos, rasgos):
    """
    Los párrafos del principio que forman la portada, y si acaba en un salto
    (más seguro) o se ha deducido por dónde empieza el texto.
    Se para en el primer título de Word, en un rótulo (Resumen, Índice…) o en
    el primer párrafo largo; si se para en un párrafo largo y justo antes hay
    algo que parece un título ("Introducción"), ese título ya no es portada.
    """
    portada = []
    for p in parrafos:
        r = rasgos[p]
        if r.salto_antes and portada:
            return portada, True
        if r.nivel_de_word or r.de_indice or _es_rotulo(r):
            return portada, False
        if r.palabras >= PALABRAS_DE_PARRAFO_LARGO:
            return _sin_titulo_final(portada, rasgos, r), False
        portada.append(p)
        if r.salto_despues:
            return portada, True
    return portada, False


def _sin_titulo_final(portada, rasgos, largo):
    con_texto = [p for p in portada if rasgos[p].texto]
    if not con_texto:
        return portada
    ultimo = rasgos[con_texto[-1]]
    mas_grande = ultimo.tamano and largo.tamano and ultimo.tamano > largo.tamano
    if _parece_titulo(ultimo) and (ultimo.negrita or ultimo.mayusculas or mas_grande):
        return portada[:portada.index(con_texto[-1])]
    return portada


# --- 3. Clasificación -----------------------------------------------------------------

def _parece_titulo(r):
    """Corto, sin enlaces, y que no acaba como una frase o va entero en negrita o mayúsculas."""
    if r.palabras >= PALABRAS_DE_TITULO or URL.search(r.texto):
        return False
    return not FINAL_DE_FRASE.search(r.texto) or r.mayusculas or r.negrita


def _parece_referencia(r):
    return bool(CON_FECHA.search(r.texto) or URL.search(r.texto))


def _nivel_por_numeracion(texto):
    numeracion = NUMERACION.match(texto)
    if not numeracion:
        return 1
    return min(len(numeracion.group(1).split(".")), NIVEL_MAXIMO)


def _extracto(texto, largo=20):
    return texto if len(texto) <= largo else texto[:largo].rstrip() + "…"


def _forma_de_titulo(r):
    rasgos = [nombre for nombre, tiene in (("en negrita", r.negrita), ("en mayúsculas", r.mayusculas),
                                           ("sin punto final", not FINAL_DE_FRASE.search(r.texto)))
              if tiene]
    return "corto, " + " y ".join(rasgos) if rasgos else "corto"


class _Clasificador:
    def __init__(self, doc, estilos):
        self.estilos = estilos or Estilos(doc)
        bloques = _bloques(doc)
        self.parrafos = [b for b in bloques if b.tag == w("p")]
        indice = _parrafos_de_indice(self.parrafos)
        self.rasgos = {p: _rasgos(self.estilos, p, indice) for p in self.parrafos}
        _vecinos(bloques, self.rasgos)
        self.portada, self.portada_con_salto = _portada(self.parrafos, self.rasgos)
        self.resultado = []

    def clasificar(self):
        zona = "cuerpo"
        portada = set(self.portada)
        for posicion, p in enumerate(self.parrafos):
            r = self.rasgos[p]
            if p in portada:
                tipo, confianza, motivo, nivel = self._en_portada(r)
                zona_del_parrafo = "portada"
            else:
                if r.salto_antes and zona in ("indice", "resumen"):
                    zona = "cuerpo"
                zona, (tipo, confianza, motivo, nivel) = self._decidir(r, zona)
                zona_del_parrafo = zona
                if r.salto_despues and zona in ("indice", "resumen"):
                    zona = "cuerpo"   # el índice y el resumen acaban en el salto de página
            self.resultado.append(Parrafo(posicion, p, r.texto, tipo, confianza, motivo,
                                          zona_del_parrafo, nivel))
        return self.resultado

    # Portada ---------------------------------------------------------------------------

    def _en_portada(self, r):
        hasta = ("antes del primer salto de página" if self.portada_con_salto
                 else "antes del primer párrafo de texto")
        if not r.texto:
            return self._vacio_o_imagen(r) + (None,)
        return "portada", ALTA if self.portada_con_salto else MEDIA, f"está al principio, {hasta}", None

    def _vacio_o_imagen(self, r):
        if r.imagen:
            return "imagen", ALTA, "solo contiene una imagen, un gráfico o una forma"
        if r.salto_despues or r.salto_antes:
            return "vacio", ALTA, "no tiene texto (lleva un salto de página o de sección)"
        return "vacio", ALTA, "no tiene texto"

    # Resto del documento ---------------------------------------------------------------

    def _decidir(self, r, zona):
        """Devuelve la zona (que puede cambiar con un rótulo o un título) y (tipo, confianza, motivo, nivel)."""
        n = _normalizar(r.texto)

        if not r.texto and not r.de_indice:
            return zona, self._vacio_o_imagen(r) + (None,)
        if r.de_indice:
            motivo = ("forma parte del índice automático de Word" if not any(ESTILO_DE_INDICE.match(e) for e in r.estilos)
                      else "tiene un estilo de índice de Word")
            return "indice", ("indice", ALTA, motivo, None)
        if n in ROTULOS_DE_INDICE and r.palabras < PALABRAS_DE_TITULO:
            return "indice", ("indice", ALTA, f"rótulo del índice («{r.texto}»)", None)
        if n in ROTULOS_DE_REFERENCIAS:
            return "referencias", ("titulo_referencias", ALTA,
                                   f"rótulo de la lista de referencias («{r.texto}»)", None)
        if ANEXO.match(r.texto) and r.palabras < PALABRAS_DE_TITULO and \
                (r.nivel_de_word or _parece_titulo(r) or r.alineacion == "center"):
            return "anexos", ("anexo", ALTA, f"rótulo de un anexo o apéndice («{_extracto(r.texto)}»)", None)
        if n in ROTULOS_DE_RESUMEN:
            return "resumen", ("titulo", ALTA, f"rótulo del resumen («{r.texto}»)", 1)
        if r.nivel_de_word:
            if zona != "anexos":
                zona = "cuerpo"
            de_estilo = any(ESTILO_DE_TITULO_NUMERADO.match(e) for e in r.estilos)
            motivo = f"estilo «{r.estilo}» de Word" if de_estilo else "nivel de esquema de Word"
            return zona, ("titulo", ALTA, motivo, r.nivel_de_word)

        if zona == "indice":
            return zona, ("indice", MEDIA, "está entre el rótulo del índice y el siguiente salto de página", None)
        if zona == "referencias":
            return self._en_referencias(r)

        decision = self._en_el_texto(r)
        if decision[0] == "titulo" and zona == "resumen":
            zona = "cuerpo"   # el resumen acaba en el siguiente título
        return zona, decision

    def _en_referencias(self, r):
        if r.francesa > 0:
            return "referencias", ("referencia", ALTA, "está en la lista de referencias, con sangría francesa", None)
        if AUTOR_Y_FECHA.match(r.texto):
            return "referencias", ("referencia", ALTA, "está en la lista de referencias y empieza por autor y año",
                                   None)
        if _parece_referencia(r):
            return "referencias", ("referencia", MEDIA, "está en la lista de referencias y tiene un año o un enlace",
                                   None)
        if _parece_titulo(r) and (r.negrita or r.mayusculas) and r.palabras <= 5:
            return "cuerpo", ("titulo", MEDIA, f"{_forma_de_titulo(r)}: acaba la lista de referencias",
                              _nivel_por_numeracion(r.texto))
        return "referencias", ("referencia", BAJA, "está en la lista de referencias, pero no tiene año ni enlace",
                               None)

    def _en_el_texto(self, r):
        # Leyendas: "Tabla 1", "Figura 2", "Nota.", o el estilo de leyenda de Word.
        leyenda = self._leyenda(r)
        if leyenda:
            return leyenda
        if r.lista_de_word:
            return "lista", ALTA, "lista de Word (con viñetas o numerada)", None
        if _parece_titulo(r) and r.despues == "tabla":
            return "leyenda_tabla", MEDIA, f"{_forma_de_titulo(r)}, justo encima de una tabla", None
        if _parece_titulo(r) and (r.antes == "imagen" or r.despues == "imagen"):
            return "leyenda_figura", MEDIA, f"{_forma_de_titulo(r)}, junto a una imagen", None
        if _parece_titulo(r) and (r.negrita or r.mayusculas):
            # "1. Cambios en…" en negrita es un título, no una lista
            return "titulo", MEDIA, self._motivo_de_titulo(r), _nivel_por_numeracion(r.texto)
        if LISTA_A_MANO.match(r.texto):
            return "lista", MEDIA, "empieza con una viñeta o un número escrito a mano", None
        if _parece_titulo(r):
            return "titulo", MEDIA, self._motivo_de_titulo(r), _nivel_por_numeracion(r.texto)
        if r.izquierda > 0 and r.francesa == 0:
            if r.palabras >= PALABRAS_DE_CITA_EN_BLOQUE:
                return "cita_bloque", ALTA, f"tiene sangría a la izquierda y {r.palabras} palabras", None
            return "cita_bloque", MEDIA, (f"tiene sangría a la izquierda, pero solo {r.palabras} palabras "
                                          f"(una cita en bloque tiene {PALABRAS_DE_CITA_EN_BLOQUE} o más)"), None
        if r.francesa > 0 and AUTOR_Y_FECHA.match(r.texto):
            return "referencia", MEDIA, "tiene sangría francesa y empieza por autor y año, fuera de la lista", None
        if FINAL_DE_FRASE.search(r.texto):
            return "cuerpo", ALTA, f"texto de {r.palabras} palabras que acaba como una frase", None
        return "cuerpo", MEDIA, f"texto de {r.palabras} palabras", None

    def _motivo_de_titulo(self, r):
        numeracion = NUMERACION.match(r.texto)
        nivel = (f"nivel {_nivel_por_numeracion(r.texto)} por la numeración «{numeracion.group(1)}»" if numeracion
                 else "nivel 1 provisional")
        return f"{_forma_de_titulo(r)} ({nivel}; se afina en el paso 3.2)"

    def _leyenda(self, r):
        de_word = any(c in e for e in r.estilos for c in ESTILOS_DE_LEYENDA)
        if LEYENDA_DE_TABLA.match(r.texto):
            tipo, de_que = "leyenda_tabla", "tabla"
        elif LEYENDA_DE_FIGURA.match(r.texto):
            tipo, de_que = "leyenda_figura", "imagen"
        elif NOTA.match(r.texto) or de_word:
            tipo, de_que = self._leyenda_de_que(r)
        else:
            return None
        junto = de_que in (r.antes, r.despues)
        etiqueta = r.texto.split()[0].rstrip(":.")
        motivo = "estilo de leyenda de Word" if de_word else f"empieza por «{etiqueta}»"
        if junto:
            motivo += f" y está junto a una {de_que}"
        return tipo, ALTA if junto or de_word else MEDIA, motivo, None

    def _leyenda_de_que(self, r):
        """Una nota ("Nota.", "Fuente:") es de una tabla o de una figura según lo que tenga cerca."""
        if "tabla" in (r.antes, r.despues):
            return "leyenda_tabla", "tabla"
        if "imagen" in (r.antes, r.despues):
            return "leyenda_figura", "imagen"
        anterior = next((c.tipo for c in reversed(self.resultado) if c.tipo != "vacio"), None)
        if anterior == "leyenda_figura":
            return "leyenda_figura", "imagen"
        return "leyenda_tabla", "tabla"


def clasificar(doc, estilos=None):
    """
    Clasifica los párrafos del cuerpo de un documento abierto con python-docx.
    Devuelve una lista de Parrafo en el orden del documento.
    """
    return _Clasificador(doc, estilos).clasificar()
