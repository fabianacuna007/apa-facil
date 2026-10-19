"""
Clasificador de párrafos (pasos 3.1 y 3.2): decide qué es cada párrafo del
cuerpo (portada, título, cuerpo, lista, referencia…) y con qué confianza,
aunque el estudiante no haya usado los estilos de Word.

Funciona en cuatro pasos:
  1. Rasgos: lo que se puede medir en cada párrafo (texto, estilo de Word,
     negrita, tamaño, sangrías, imágenes, saltos, campos de índice…) y qué
     tiene al lado (una tabla, una imagen, el párrafo anterior y el siguiente).
  2. Portada: los párrafos del principio, hasta el primer salto, el primer
     título o el primer párrafo largo.
  3. Recorrido en orden: cada párrafo recibe su tipo y la zona del trabajo
     en la que está (portada, índice, resumen, cuerpo, referencias, anexos).
     Los títulos que no usan los estilos de Word ("títulos falsos") se
     reconocen sumando indicios: negrita, mayúsculas, letra más grande que el
     texto, numeración de apartado, sin punto final, seguidos de texto…
  4. Niveles de los títulos falsos: por la numeración ("1.1." es nivel 2) o,
     si no la tienen, por su formato comparado con el de los demás títulos.
     Un título falso no puede bajar más de un nivel respecto al anterior.

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
import textwrap
from collections import Counter
from dataclasses import dataclass, field
from typing import Optional

from apa_facil.documento import (LEYENDA_CON_NUMERO, MC, ROTULO_DE_APENDICE, ROTULOS_DE_INDICE, ROTULOS_DE_REFERENCIAS,
                                 WP, Estilos,
                                 texto_parrafo, w)
from apa_facil.reglas.alineacion import _salto_antes, _salto_en

TIPOS = ("portada", "indice", "titulo", "cuerpo", "lista", "leyenda_tabla", "leyenda_figura",
         "cita_bloque", "titulo_referencias", "referencia", "anexo", "imagen", "vacio")
ZONAS = ("portada", "indice", "resumen", "cuerpo", "referencias", "anexos")
ALTA, MEDIA, BAJA = "alta", "media", "baja"

PALABRAS_DE_TITULO = 15        # un título casi nunca llega a 15 palabras
PALABRAS_DE_PARRAFO_LARGO = 25  # un renglón de portada casi nunca llega a 20
PALABRAS_DE_CITA_EN_BLOQUE = 40  # APA: 40 palabras o más
NIVEL_MAXIMO = 5

ROTULOS_DE_RESUMEN = {"resumen", "abstract"}
ANEXO = re.compile(r"^(anexo|anexos|apéndice|apendice|apéndices|apendices)\b", re.IGNORECASE)
LISTA_A_MANO = re.compile(r"^\s*([•·▪◦‣●○■□\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z]\))\s")
LEYENDA_DE_TABLA = re.compile(r"^(tabla|cuadro)\s+[A-Z]?\d", re.IGNORECASE)
LEYENDA_DE_FIGURA = re.compile(r"^(figura|gráfico|grafico|imagen|ilustración|ilustracion|mapa|"
                               r"fotografía|fotografia|diagrama|esquema)\s+[A-Z]?\d", re.IGNORECASE)
NOTA = re.compile(r"^(nota\s*\.|fuente\s*:|fuentes\s*:)", re.IGNORECASE)
FINAL_DE_FRASE = re.compile(r"[.…:][\"'»”)]*$")
ACABA_EN_PUNTO = re.compile(r"[.…][\"'»”)]*$")
VINETA = re.compile(r"^\s*[•·▪◦‣●○■□\-–—*]\s")
# Numeración de apartado: "1.", "1.1.", "2.3.1", "1)", "I.", "II.", "A.", "a)"
NUMERACION_ARABIGA = re.compile(r"^(\d{1,2}(?:\.\d{1,2})*)[.)]?\s")
NUMERACION_ROMANA = re.compile(r"^([IVX]+)[.)]\s")
NUMERACION_CON_LETRA = re.compile(r"^([A-Za-z])[.)]\s")
URL = re.compile(r"https?://|www\.|doi\.org|\bdoi:", re.IGNORECASE)
ETIQUETA = re.compile(r"^[^\W\d_]+( [^\W\d_]+)?:\s")   # «Autora: …», «Trabajo de: …» (datos de la portada)
AUTOR_Y_FECHA = re.compile(r"^[^\W\d][^,(]{0,80},.*(\(\s*(\d{4}|s\.\s?f\.)|\b\d{4}\b)", re.IGNORECASE)
CON_FECHA = re.compile(r"\(\s*(\d{4}[a-z]?|s\.\s?f\.)\s*[),]|\b(1[5-9]|20)\d{2}\b")
ESTILO_DE_TITULO_NUMERADO = re.compile(r"^(heading|título|titulo)\s*(\d)$")
ESTILO_DE_INDICE = re.compile(r"^(toc|tdc)\s*\d$|^(toc heading|título de tdc|titulo de tdc)$")
ESTILOS_DE_LEYENDA = ("caption", "descripción", "descripcion", "epígrafe", "epigrafe")
ESTILOS_DE_TITULO_DEL_TRABAJO = ("title", "subtitle", "título", "titulo", "subtítulo", "subtitulo")


@dataclass
class Parrafo:
    """Un párrafo clasificado."""
    posicion: int                 # 0, 1, 2… entre los párrafos clasificados
    elemento: object              # el w:p del XML
    texto: str                    # en una sola línea, sin espacios repetidos
    tipo: str                     # uno de TIPOS
    confianza: str                # alta, media o baja
    motivo: str                   # por qué, en español
    zona: str                     # uno de ZONAS
    nivel: Optional[int] = None   # solo en los títulos: de 1 a 5
    # El número con que se nombra el párrafo en el diagnóstico, el verificador y la
    # radiografía: desde 1 y contando también los párrafos de las tablas.
    numero: int = 0
    # En la portada: ¿es el título del trabajo (o su subtítulo)? (paso 4.4)
    titulo_del_trabajo: bool = False


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
    cursiva: bool
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
    # Los párrafos con texto de justo antes y justo después (None si hay una tabla o una imagen).
    anterior: Optional["_Rasgos"] = field(default=None, repr=False)
    siguiente: Optional["_Rasgos"] = field(default=None, repr=False)


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


def _todo_con(estilos, p, fragmentos, etiqueta):
    """¿Va todo el texto del párrafo en negrita (etiqueta "b") o en cursiva ("i")?"""
    if not fragmentos:
        return False
    for r in fragmentos:
        valor = next((capa.find(w(etiqueta)) for capa in estilos._capas_texto(r, p)
                      if capa is not None and capa.find(w(etiqueta)) is not None), None)
        if valor is None or valor.get(w("val")) in ("0", "false", "off"):
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
        negrita=_todo_con(estilos, p, fragmentos, "b"),
        cursiva=_todo_con(estilos, p, fragmentos, "i"),
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

    def de_texto(i):
        if 0 <= i < len(utiles) and utiles[i][1] == "texto":
            return rasgos[utiles[i][0]]
        return None

    utiles = [(el, que_es(el)) for el in bloques]
    utiles = [(el, cosa) for el, cosa in utiles if cosa]
    for i, (el, _) in enumerate(utiles):
        if el.tag == w("p"):
            rasgos[el].antes = utiles[i - 1][1] if i > 0 else None
            rasgos[el].despues = utiles[i + 1][1] if i + 1 < len(utiles) else None
            rasgos[el].anterior, rasgos[el].siguiente = de_texto(i - 1), de_texto(i + 1)


def _tamano_del_texto(rasgos):
    """El tamaño de letra del texto normal: el de los párrafos largos, contando letras."""
    letras = Counter()
    for r in rasgos:
        if r.tamano and r.palabras >= PALABRAS_DE_PARRAFO_LARGO:
            letras[r.tamano] += len(r.texto)
    if not letras:
        letras.update({r.tamano: len(r.texto) for r in rasgos if r.tamano})
    return letras.most_common(1)[0][0] if letras else None


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


def _numeracion(texto):
    """
    La numeración de apartado con que empieza el texto, como (clase, número, nivel):
    ("arábiga", "1.1", 2), ("romana", "II", 1), ("letra", "a", None), o None si no tiene.
    """
    for clase, patron in (("arábiga", NUMERACION_ARABIGA), ("romana", NUMERACION_ROMANA),
                          ("letra", NUMERACION_CON_LETRA)):
        encontrada = patron.match(texto)
        if encontrada:
            numero = encontrada.group(1)
            nivel = {"arábiga": min(len(numero.split(".")), NIVEL_MAXIMO), "romana": 1}.get(clase)
            return clase, numero, nivel
    return None


def _marca(texto):
    """Con qué empieza una línea de lista o de apartado: viñeta, o clase y profundidad de la numeración."""
    if VINETA.match(texto):
        return ("viñeta",)
    numeracion = _numeracion(texto)
    return (numeracion[0], numeracion[2]) if numeracion else None


def _en_serie(r):
    """¿Es una línea más de una lista escrita a mano? (el párrafo de antes o el de después empieza igual)"""
    marca = _marca(r.texto)
    return marca is not None and any(vecino is not None and _marca(vecino.texto) == marca
                                     for vecino in (r.anterior, r.siguiente))


def _seguido_de_texto(r):
    """¿Va justo antes de un párrafo de texto normal (y no de otra línea corta o de una lista)?"""
    s = r.siguiente
    return (s is not None and not s.nivel_de_word and not s.lista_de_word and not LISTA_A_MANO.match(s.texto)
            and (s.palabras >= PALABRAS_DE_TITULO or bool(ACABA_EN_PUNTO.search(s.texto))))


def _formato(r):
    """El aspecto de un título: (tamaño, centrado, mayúsculas, negrita, cursiva)."""
    return (r.tamano or 0, r.alineacion == "center", r.mayusculas, r.negrita, r.cursiva)


def _destacado(formato):
    """Para ordenar formatos del más destacado al menos: más grande, centrado, mayúsculas, negrita, sin cursiva."""
    tamano, centrado, mayusculas, negrita, cursiva = formato
    return (-tamano, not centrado, not mayusculas, not negrita, cursiva)


def _describir(formato):
    tamano, centrado, mayusculas, negrita, cursiva = formato
    partes = [_pt(tamano)] if tamano else []
    partes += [nombre for nombre, tiene in (("centrado", centrado), ("en mayúsculas", mayusculas),
                                            ("en negrita", negrita), ("en cursiva", cursiva)) if tiene]
    return ", ".join(partes) or "sin formato especial"


def _pt(tamano):
    return f"{tamano:g} pt".replace(".", ",")


def _lista_de_palabras(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


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
        portada = set(self.portada)
        self.tamano_del_texto = _tamano_del_texto([self.rasgos[p] for p in self.parrafos if p not in portada])
        self.numeros = {p: i for i, p in enumerate(doc.element.body.iter(w("p")), start=1)}
        self.resultado = []
        self.falsos = {}   # posición de cada título falso → indicios que lo delatan

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
                if tipo == "titulo" and nivel is None:
                    self.falsos[posicion], motivo = motivo, ""   # el motivo se escribe con el nivel
                if r.salto_despues and zona in ("indice", "resumen"):
                    zona = "cuerpo"   # el índice y el resumen acaban en el salto de página
            self.resultado.append(Parrafo(posicion, p, r.texto, tipo, confianza, motivo,
                                          zona_del_parrafo, nivel, self.numeros[p]))
        self._niveles_de_titulos_falsos()
        titulo = set(self._titulo_del_trabajo())
        for c in self.resultado:
            c.titulo_del_trabajo = c.elemento in titulo
        return self.resultado

    # Título del trabajo (paso 4.4) -------------------------------------------------------

    def _titulo_del_trabajo(self):
        """
        Los párrafos de la portada que forman el título del trabajo: el que usa el
        estilo «Title» de Word o, si no, el de letra más grande (si solo hay uno así);
        y, justo debajo, el subtítulo, si va en negrita y con letra más grande que el
        texto. Se decide sobre el original: después de la regla de fuente, todos los
        párrafos tienen el mismo tamaño.
        """
        con_texto = [p for p in self.portada if self.rasgos[p].texto]
        de_estilo = [p for p in con_texto if self.rasgos[p].estilos[:1] in (["title"], ["título"], ["titulo"])]
        if de_estilo:
            titulo = de_estilo[0]
        else:
            tamanos = [self.rasgos[p].tamano or 0 for p in con_texto]
            if not tamanos or tamanos.count(max(tamanos)) != 1:
                return []
            titulo = con_texto[tamanos.index(max(tamanos))]
        partes = [titulo]
        siguiente = self.parrafos.index(titulo) + 1
        if siguiente < len(self.parrafos) and self.parrafos[siguiente] in con_texto:
            r = self.rasgos[self.parrafos[siguiente]]
            if (r.negrita and r.tamano and self.tamano_del_texto and r.tamano > self.tamano_del_texto
                    and r.palabras < PALABRAS_DE_TITULO and not ETIQUETA.match(r.texto)):
                partes.append(self.parrafos[siguiente])
        return partes

    # Niveles de los títulos falsos -----------------------------------------------------

    def _niveles_de_titulos_falsos(self):
        """
        1. Por la numeración: "1." nivel 1, "1.1." nivel 2… ("I." es nivel 1 y, si el
           documento usa números romanos, "1." pasa a nivel 2); "a)" va un nivel
           por debajo del título anterior.
        2. Sin numeración, por el formato: los títulos de Word y los numerados dicen
           qué formato tiene cada nivel; un formato que no sale en ellos se coloca
           según lo destacado que es (tamaño, centrado, mayúsculas, negrita, cursiva).
        3. Un título falso no baja más de un nivel respecto al título anterior.
        """
        titulos = [c for c in self.resultado if c.tipo == "titulo"]
        romanos = any((_numeracion(c.texto) or ("",))[0] == "romana"
                      for c in titulos if c.posicion in self.falsos)

        def por_numeracion(c):
            numeracion = _numeracion(c.texto)
            if not numeracion or numeracion[0] == "letra":
                return None
            clase, _, nivel = numeracion
            return min(nivel + 1, NIVEL_MAXIMO) if clase == "arábiga" and romanos else nivel

        # Qué formato tiene cada nivel según los títulos de Word y los numerados.
        votos = {}
        for c in titulos:
            r = self.rasgos[c.elemento]
            nivel = por_numeracion(c) if c.posicion in self.falsos else r.nivel_de_word
            if nivel:
                votos.setdefault(_formato(r), Counter())[nivel] += 1
        anclados = {formato: cuenta.most_common(1)[0][0] for formato, cuenta in votos.items()}

        # Los formatos sin nivel conocido se colocan entre los demás, del más destacado al menos.
        formatos = set(anclados) | {_formato(self.rasgos[c.elemento]) for c in titulos if c.posicion in self.falsos}
        por_formato, actual = {}, 0
        for formato in sorted(formatos, key=_destacado):
            actual = anclados.get(formato, min(actual + 1, NIVEL_MAXIMO))
            por_formato[formato] = actual

        anterior = sin_letra = None   # el nivel del título anterior, y del último que no es "a)", "b)"…
        for c in titulos:
            if c.posicion in self.falsos:
                c.nivel, explicacion = self._nivel_de_titulo_falso(c, sin_letra, por_numeracion(c),
                                                                   por_formato, anclados)
                if anterior is not None and c.nivel > anterior + 1:
                    explicacion += (f"; se sube a nivel {anterior + 1} porque el título anterior es de nivel "
                                    f"{anterior} y no se puede saltar un nivel")
                    c.nivel, c.confianza = anterior + 1, BAJA
                c.motivo = (f"título sin estilo de Word: {_lista_de_palabras(self.falsos[c.posicion])}. "
                            f"{explicacion[0].upper()}{explicacion[1:]}.")
            anterior = c.nivel
            if c.posicion not in self.falsos or (_numeracion(c.texto) or ("",))[0] != "letra":
                sin_letra = c.nivel

    def _nivel_de_titulo_falso(self, c, sin_letra, por_numeracion, por_formato, anclados):
        numeracion = _numeracion(c.texto)
        if por_numeracion:
            return por_numeracion, f"nivel {por_numeracion} por la numeración «{numeracion[1]}»"
        if numeracion:   # "a)", "b)": un nivel por debajo del último título sin letra
            nivel = min((sin_letra or 0) + 1, NIVEL_MAXIMO)
            return nivel, f"nivel {nivel}: la letra «{numeracion[1]}» va por debajo del título anterior"
        formato = _formato(self.rasgos[c.elemento])
        nivel = por_formato[formato]
        if formato in anclados:
            return nivel, f"nivel {nivel}: mismo formato que los demás títulos de nivel {nivel}"
        if len(por_formato) == 1:
            return nivel, "nivel 1: es el único formato de título del documento"
        if nivel == 1:
            return nivel, f"nivel 1: su formato ({_describir(formato)}) es el más destacado de los títulos"
        return nivel, (f"nivel {nivel}: su formato ({_describir(formato)}) es menos destacado que el de los "
                       f"títulos de nivel {nivel - 1}")

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
        rotulo = ROTULO_DE_APENDICE.match(r.anterior.texto) if r.anterior and zona == "anexos" else None
        if rotulo and not rotulo["titulo"] and not ANEXO.match(r.texto) and r.palabras < PALABRAS_DE_TITULO:
            # El título que va debajo de «Apéndice A» (como lo deja el paso 6.7).
            return "anexos", ("anexo", ALTA, f"título del apéndice: va justo debajo de «{r.anterior.texto}»", None)
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
            indicios = [n for n, t in (("en negrita", r.negrita), ("en mayúsculas", r.mayusculas)) if t]
            return "cuerpo", ("titulo", MEDIA, indicios + ["acaba la lista de referencias"], None)
        return "referencias", ("referencia", BAJA, "está en la lista de referencias, pero no tiene año ni enlace",
                               None)

    def _en_el_texto(self, r):
        # Leyendas: "Tabla 1", "Figura 2", "Nota.", o el estilo de leyenda de Word.
        leyenda = self._leyenda(r)
        if leyenda:
            return leyenda
        if r.lista_de_word:
            return "lista", ALTA, "lista de Word (con viñetas o numerada)", None
        numero = LEYENDA_CON_NUMERO.match(r.anterior.texto) if r.anterior else None
        if numero and not numero["titulo"] and r.despues in ("tabla", "imagen"):   # como lo dejan los pasos 6.1 y 6.2
            de_que, encima_de = ("tabla", "la tabla") if r.despues == "tabla" else ("figura", "la imagen")
            return (f"leyenda_{de_que}", ALTA, f"título de {de_que}: debajo de «{r.anterior.texto}» y encima de "
                    f"{encima_de}", None)
        if _parece_titulo(r) and r.despues == "tabla":
            return "leyenda_tabla", MEDIA, f"{_forma_de_titulo(r)}, justo encima de una tabla", None
        if _parece_titulo(r) and (r.antes == "imagen" or r.despues == "imagen"):
            return "leyenda_figura", MEDIA, f"{_forma_de_titulo(r)}, junto a una imagen", None
        titulo = self._titulo_falso(r)
        if titulo:
            confianza, indicios = titulo
            return "titulo", confianza, indicios, None   # el nivel se decide al final
        if LISTA_A_MANO.match(r.texto):
            motivo = "empieza con una viñeta o un número escrito a mano"
            if _en_serie(r):
                motivo += ", igual que la línea de al lado"
            return "lista", MEDIA, motivo, None
        if r.izquierda > 0 and r.francesa == 0:
            if r.palabras >= PALABRAS_DE_CITA_EN_BLOQUE:
                return "cita_bloque", ALTA, f"tiene sangría a la izquierda y {r.palabras} palabras", None
            return "cita_bloque", MEDIA, (f"tiene sangría a la izquierda, pero solo {r.palabras} palabras "
                                          f"(una cita en bloque tiene {PALABRAS_DE_CITA_EN_BLOQUE} o más)"), None
        if r.francesa > 0:
            # En un trabajo, la sangría francesa es casi siempre de una referencia: mejor no tocarla.
            if AUTOR_Y_FECHA.match(r.texto):
                return "referencia", MEDIA, "tiene sangría francesa y empieza por autor y año, fuera de la lista", None
            return "referencia", BAJA, "tiene sangría francesa, pero está fuera de la lista y sin autor y año", None
        if FINAL_DE_FRASE.search(r.texto):
            return "cuerpo", ALTA, f"texto de {r.palabras} palabras que acaba como una frase", None
        return "cuerpo", MEDIA, f"texto de {r.palabras} palabras", None

    def _titulo_falso(self, r):
        """
        ¿Es un título aunque no use los estilos de Word? Se suman indicios.
        Devuelve (confianza, indicios) o None si no parece un título.
        Indicios fuertes: negrita, mayúsculas, letra más grande que el texto y
        numeración de apartado seguida de texto. Débiles: sin punto final,
        cursiva, centrado y seguido de un párrafo de texto.
        """
        if r.palabras >= PALABRAS_DE_TITULO or URL.search(r.texto) or VINETA.match(r.texto):
            return None
        if _en_serie(r):
            return None   # una línea más de una lista numerada a mano
        seguido = _seguido_de_texto(r)
        numeracion = _numeracion(r.texto)
        # "1. Marco teórico" seguido directamente de "1.1. Antecedentes" también es un apartado.
        del_siguiente = _numeracion(r.siguiente.texto) if r.siguiente else None
        subapartado = bool(numeracion and del_siguiente and numeracion[0] == del_siguiente[0] == "arábiga"
                           and del_siguiente[2] == numeracion[2] + 1)
        mas_grande = bool(r.tamano and self.tamano_del_texto and r.tamano > self.tamano_del_texto)
        fuertes = [nombre for nombre, tiene in (
            ("en negrita", r.negrita),
            ("en mayúsculas", r.mayusculas),
            (f"con letra más grande que el texto ({_pt(r.tamano or 0)} frente a {_pt(self.tamano_del_texto or 0)})",
             mas_grande),
            ("numerado como un apartado", numeracion is not None and (seguido or subapartado))) if tiene]
        debiles = [nombre for nombre, tiene in (
            ("sin punto final", not FINAL_DE_FRASE.search(r.texto)),
            ("en cursiva", r.cursiva),
            ("centrado", r.alineacion == "center"),
            ("seguido de un párrafo de texto", seguido)) if tiene]
        indicios = fuertes + debiles

        if ACABA_EN_PUNTO.search(r.texto):
            # Una frase corta en negrita o en mayúsculas: puede ser un título, pero es dudoso.
            return (BAJA, indicios + ["aunque acaba en punto"]) if r.negrita or r.mayusculas else None
        if fuertes:
            return ALTA if len(indicios) >= 3 else MEDIA, indicios
        if "sin punto final" in debiles and (seguido or r.alineacion == "center"):
            return BAJA, indicios
        return None

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


class Clasificacion(list):
    """
    La lista de párrafos clasificados (Parrafo) y, en `tablas`, el número de párrafo
    del primero de cada tabla del cuerpo en el documento clasificado (paso 6.1: la
    integridad nombra los párrafos por su número en el original, y las tablas no
    tienen Parrafo; después de borrar párrafos ya no se puede contar).
    """

    def __init__(self, parrafos=(), tablas=None):
        super().__init__(parrafos)
        self.tablas = dict(tablas or {})


def clasificar(doc, estilos=None):
    """
    Clasifica los párrafos del cuerpo de un documento abierto con python-docx.
    Devuelve una Clasificacion: una lista de Parrafo en el orden del documento.
    """
    clasificador = _Clasificador(doc, estilos)
    parrafos = clasificador.clasificar()
    tablas = {}
    for tabla in (b for b in _bloques(doc) if b.tag == w("tbl")):
        primero = tabla.find(f".//{w('p')}")
        if primero is not None:
            tablas[tabla] = clasificador.numeros[primero]
    return Clasificacion(parrafos, tablas)


# --- Modo diagnóstico (paso 3.3) ------------------------------------------------------

NOMBRES_DE_TIPO = {"portada": "portada", "indice": "índice", "titulo": "título", "cuerpo": "cuerpo",
                   "lista": "lista", "leyenda_tabla": "leyenda de tabla", "leyenda_figura": "leyenda de figura",
                   "cita_bloque": "cita en bloque", "titulo_referencias": "rótulo de referencias",
                   "referencia": "referencia", "anexo": "rótulo de anexo", "imagen": "imagen", "vacio": "vacío"}
NOMBRES_DE_ZONA = {"portada": "Portada", "indice": "Índice", "resumen": "Resumen", "cuerpo": "Cuerpo",
                   "referencias": "Referencias", "anexos": "Anexos"}
ANCHO_DEL_TIPO = max(len(nombre) for nombre in NOMBRES_DE_TIPO.values())
ANCHO_DE_LA_CONFIANZA = len("Confianza")
ANCHO_MINIMO_DEL_TEXTO = 20


def nombre_de_tipo(c):
    """"título (nivel 2)", "leyenda de tabla"…"""
    nombre = NOMBRES_DE_TIPO[c.tipo]
    return f"{nombre} (nivel {c.nivel})" if c.nivel else nombre


def como_texto(clasificados, ancho=100):
    """
    La clasificación en forma de tabla para la terminal: número de párrafo,
    tipo, confianza y el principio del texto, con una línea cada vez que
    empieza una zona. Los párrafos de confianza media o baja llevan debajo el
    motivo (son los dudosos, los que hay que revisar). Al final, los recuentos.
    """
    if not clasificados:
        return "El documento no tiene párrafos de texto."
    numero = max(2, len(str(clasificados[-1].numero)))
    columnas = numero + 2 + ANCHO_DEL_TIPO + 2 + ANCHO_DE_LA_CONFIANZA + 2
    texto = max(ancho - columnas, ANCHO_MINIMO_DEL_TEXTO)
    ancho = columnas + texto

    lineas = [f"{'Nº':>{numero}}  {'Tipo':<{ANCHO_DEL_TIPO}}  {'Confianza':<{ANCHO_DE_LA_CONFIANZA}}  Texto"]
    zona = None
    for c in clasificados:
        if c.zona != zona:
            zona = c.zona
            rotulo = f"── {NOMBRES_DE_ZONA[zona]} "
            lineas.append(rotulo + "─" * (ancho - len(rotulo)))
        extracto = c.texto or "(sin texto)"
        if len(extracto) > texto:
            extracto = extracto[:texto - 1].rstrip() + "…"
        lineas.append(f"{c.numero:>{numero}}  {nombre_de_tipo(c):<{ANCHO_DEL_TIPO}}  "
                      f"{c.confianza:<{ANCHO_DE_LA_CONFIANZA}}  {extracto}")
        if c.confianza != ALTA:
            lineas.append(textwrap.fill(c.motivo, width=ancho, initial_indent=" " * (numero + 2) + "↳ ",
                                        subsequent_indent=" " * (numero + 4)))

    confianzas = Counter(c.confianza for c in clasificados)
    tipos = Counter(NOMBRES_DE_TIPO[c.tipo] for c in clasificados)
    lineas += ["", f"Resumen: {len(clasificados)} párrafos · confianza alta {confianzas[ALTA]}, "
                   f"media {confianzas[MEDIA]}, baja {confianzas[BAJA]}",
               *_en_lineas("Tipos: ", [f"{nombre} {n}" for nombre, n in tipos.most_common()], ancho)]
    return "\n".join(lineas)


def _en_lineas(prefijo, partes, ancho, separador=" · "):
    """Une las partes en líneas de `ancho` como mucho, sin partir ninguna parte."""
    lineas, actual = [], prefijo + partes[0]
    for parte in partes[1:]:
        if len(actual) + len(separador) + len(parte) > ancho:
            lineas.append(actual + separador.rstrip())
            actual = " " * len(prefijo) + parte
        else:
            actual += separador + parte
    return lineas + [actual]
