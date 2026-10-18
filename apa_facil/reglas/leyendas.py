"""
Leyendas de tablas y figuras: lo que comparten los pasos 6.1 y 6.2.

Según la web oficial de APA (Table setup, Figure setup y Paragraph alignment and
indentation: «tables and figures have the same overall setup»):
- encima, el número («Tabla 1», «Figura 1») en negrita y, en la línea siguiente,
  el título en cursiva, con mayúscula inicial en las palabras principales; los dos
  a la izquierda, sin sangría y a doble espacio (el doble espacio ya lo pone el
  paso 1.3);
- debajo, si la hay, la nota: «Nota.» en cursiva y el texto en letra normal, a la
  izquierda y sin sangría;
- numeradas en el orden en que se mencionan;
- arriba o abajo de la página, no en medio, y con una línea en blanco entre la
  tabla o figura y el texto: solo se avisa.

Lo que cambia el texto, anotado en el informe y comprobado por la integridad
(decisiones del propietario): la leyenda que está debajo sube encima; la palabra
del rótulo pasa a ser la de APA («Cuadro 1» → «Tabla 1», «Gráfico 1» → «Figura 1»);
el número y el título que van en la misma línea se separan, y desaparece lo que
había entre ellos («: », «. », « – »); el título lleva las mayúsculas de APA, como
los títulos del paso 4.6 (--titulos-tal-cual las deja); si ninguna tabla (o
figura) tiene número, se añade «Tabla 1», «Tabla 2»… a las que tienen título (si
unas lo tienen y otras no, solo se avisa, para no descuadrar la numeración); y a
la nota que empieza por «Fuente:» se le pone «Nota. » delante.

Cada regla (tablas.py, figuras.py) dice qué es una tabla o una figura y llama a
`buscar` y a `tratar`. En un apéndice no se añade el número: APA las numera con
la letra del apéndice (paso 6.7).
"""

import copy
import re
from dataclasses import dataclass, field
from typing import Optional

from docx.oxml import OxmlElement
from lxml import etree

from apa_facil.documento import (LEYENDA_CON_NUMERO, NOTA_SIN_ROTULO, PALABRAS_DE_ROTULO, ROTULO_DE_NOTA, Estilos,
                                 texto_parrafo, w)
from apa_facil.estructura import ESTILOS_DE_LEYENDA, NOTA, _bloques
from apa_facil.reglas import titulos
from apa_facil.reglas.fuente import _runs_propios
from apa_facil.reglas.resumen import _activado, _partir, _poner

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
ZONAS_SIN_LEYENDAS_APA = ("portada", "indice")
# Lo que impide separar el número del título sin dudas: cambios marcados, saltos,
# imágenes, notas, comentarios, símbolos.
NO_SE_DIVIDE_SI_LLEVA = ("ins", "del", "delText", "moveFrom", "moveTo", "br", "cr", "drawing", "pict",
                         "footnoteReference", "endnoteReference", "commentReference", "sectPr", "sym",
                         "noBreakHyphen", "txbxContent")
# El título lleva dentro lo que APA pone en una nota: «(elaboración propia)», «(Fuente: INE)».
NOTA_EN_EL_TITULO = re.compile(r"\((elaboraci[oó]n propia|fuentes?\s*:[^)]*)\)", re.IGNORECASE)


@dataclass(frozen=True)
class Clase:
    """Tablas o figuras: cómo se nombran."""
    singular: str          # "tabla"
    plural: str            # "tablas"
    rotulo: str            # "Tabla": la palabra de APA
    tipo_de_leyenda: str   # "leyenda_tabla": el tipo que les da el clasificador

    @property
    def titulo(self):      # «Tablas: …» al principio de los mensajes del informe
        return self.plural.capitalize()

    @property
    def numero_o_letra(self):   # «Tabla 1» o, en un apéndice, «Tabla A1»
        return re.compile(rf"{self.rotulo} [A-Z]?[1-9]\d{{0,2}}")


TABLA = Clase("tabla", "tablas", "Tabla", "leyenda_tabla")
FIGURA = Clase("figura", "figuras", "Figura", "leyenda_figura")


@dataclass
class Objeto:
    """Una tabla o una figura y sus leyendas."""
    elemento: object                # el w:tbl, o el w:p de la imagen
    zona: str
    destino: Optional[int]          # número de su primer párrafo en el original (para la integridad)
    encima: list = field(default_factory=list)   # Parrafo de las leyendas de encima, en orden
    debajo: list = field(default_factory=list)
    numero: object = None           # el Parrafo con «Tabla N» (o _NumeroNuevo)
    titulo: object = None           # el w:p del título (puede ser nuevo)
    notas: list = field(default_factory=list)


@dataclass
class _NumeroNuevo:
    """El párrafo «Tabla N» añadido (hace las veces de un Parrafo clasificado)."""
    elemento: object


def _extracto(texto, largo=40):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _plural(n, singular, plural=None):
    return f"{n} {singular if n == 1 else plural or singular + 's'}"


def _lista(textos, maximo=3):
    return ", ".join(f"«{_extracto(t)}»" for t in textos[:maximo]) + ("…" if len(textos) > maximo else "")


def _texto(c):
    return " ".join(texto_parrafo(c.elemento).split())


def _n(cuantos):
    """La «n» del plural de un verbo: «estaba» / «estaban»."""
    return "n" if cuantos > 1 else ""


# --- Qué tablas o figuras hay y cuáles son sus leyendas --------------------------------------

def _sigue_un_objeto(bloques, j, de, es_objeto):
    """¿Va el bloque j justo antes de otra tabla o figura (solo con párrafos vacíos en medio)?"""
    for b in bloques[j + 1:]:
        if es_objeto(b):
            return True
        c = de.get(b)
        if c is None or c.tipo != "vacio":
            return False
    return False


def _leyendas(bloques, k, paso, de, usados, maximo, clase, es_objeto):
    """Las leyendas seguidas encima (paso -1) o debajo (+1) del bloque k, sin contar los vacíos."""
    encontradas, j = [], k + paso
    while 0 <= j < len(bloques) and len(encontradas) < maximo:
        c = de.get(bloques[j])
        if c is None or c.elemento in usados or c.tipo not in ("vacio", clase.tipo_de_leyenda):
            break
        if c.tipo == clase.tipo_de_leyenda:
            # Una leyenda con número que va justo antes de otra tabla o figura es de esa otra.
            if paso > 0 and LEYENDA_CON_NUMERO.match(_texto(c)) and _sigue_un_objeto(bloques, j, de, es_objeto):
                break
            encontradas.append(c)
        j += paso
    return encontradas if paso > 0 else encontradas[::-1]


def buscar(doc, clasificados, clase, es_objeto, destino_de):
    """
    Las tablas o figuras del texto (`es_objeto(bloque)`), con sus leyendas.
    `destino_de(bloque)`: el número de su primer párrafo en el original.
    No se miran las de la portada ni las del índice.
    """
    de = {c.elemento: c for c in clasificados if c.elemento.getparent() is not None}
    bloques = _bloques(doc)
    resultado, usados = [], set()
    for k, b in enumerate(bloques):
        if not es_objeto(b):
            continue
        vecinos = [de[x] for x in reversed(bloques[:k]) if x in de and x is not b] + \
                  [de[x] for x in bloques[k + 1:] if x in de]
        zona = de[b].zona if b in de else vecinos[0].zona if vecinos else "cuerpo"
        if zona in ZONAS_SIN_LEYENDAS_APA:
            continue
        o = Objeto(b, zona, destino_de(b))
        o.encima = _leyendas(bloques, k, -1, de, usados, 2, clase, es_objeto)
        usados.update(c.elemento for c in o.encima)
        o.debajo = _leyendas(bloques, k, +1, de, usados, 3, clase, es_objeto)
        usados.update(c.elemento for c in o.debajo)
        resultado.append(o)
    return resultado


# --- Separar, mover y añadir ---------------------------------------------------------------

def _largo(r):
    return sum(len(el.text or "") if el.tag == w("t") else 1 for el in r if el.tag in (w("t"), w("tab")))


def _posiciones(p):
    """Cada fragmento del párrafo con la posición (en su texto) en que empieza."""
    resultado, posicion = [], 0
    for r in _runs_propios(p):
        resultado.append((r, posicion))
        posicion += _largo(r)
    return resultado


def _se_puede_cambiar(p, corte):
    """
    ¿Se puede partir el párrafo en `corte`, o cambiar su texto, sin estropear un campo
    de Word (p. ej. la numeración automática) ni unos cambios marcados?
    """
    if any(next(p.iter(w(e)), None) is not None for e in NO_SE_DIVIDE_SI_LLEVA):
        return False
    abiertos = []
    for r, posicion in _posiciones(p):
        for marca in r.iter(w("fldChar")):
            if marca.get(w("fldCharType")) == "begin":
                abiertos.append(posicion)
            elif marca.get(w("fldCharType")) == "end" and abiertos and (abiertos.pop() < corte) != (posicion < corte):
                return False
    for simple in p.iter(w("fldSimple")):
        dentro = [posicion for r, posicion in _posiciones(p) if next(r.iterancestors(w("fldSimple")), None) is simple]
        if dentro and (min(dentro) < corte) != (max(dentro) < corte):
            return False
    return not abiertos


def _quedarse_con(p, desde, hasta, corte):
    """
    Deja en el párrafo solo el texto entre las posiciones `desde` y `hasta`. Los
    fragmentos sin texto (marcas de campo) se quedan en la parte en la que están
    respecto a `corte`. Se quitan los fragmentos que se quedan vacíos.
    """
    primera = desde == 0
    for r, inicio in _posiciones(p):
        largo = _largo(r)
        if not largo:
            if (inicio < corte) != primera:
                r.getparent().remove(r)
            continue
        posicion = inicio
        for el in [el for el in r if el.tag in (w("t"), w("tab"))]:
            if el.tag == w("tab"):
                if not desde <= posicion < hasta:
                    r.remove(el)
                posicion += 1
                continue
            texto = el.text or ""
            el.text = texto[max(0, desde - posicion):max(0, hasta - posicion)]
            el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            posicion += len(texto)
        if not _largo(r):
            r.getparent().remove(r)
    for contenedor in list(p.iter(w("hyperlink"), w("fldSimple"), w("smartTag"))):
        if next(contenedor.iter(w("r")), None) is None:
            contenedor.getparent().remove(contenedor)


def _dividir(p, fin, desde):
    """Parte la leyenda en dos párrafos: el número (hasta `fin`) y el título (desde `desde`). Devuelve el del título."""
    titulo = copy.deepcopy(p)
    for marca in list(titulo.iter(w("bookmarkStart"), w("bookmarkEnd"))):   # no se repiten
        marca.getparent().remove(marca)
    for atributo in [a for a in titulo.attrib if a.startswith(f"{{{W14}}}")]:   # identificadores de párrafo
        del titulo.attrib[atributo]
    _quedarse_con(p, 0, fin, desde)
    _quedarse_con(titulo, desde, len(texto_parrafo(titulo)), desde)
    p.addnext(titulo)
    return titulo


def _cambiar_palabra(p, largo, nueva):
    """Cambia las `largo` primeras letras del párrafo por `nueva`, en el primer fragmento que las contiene."""
    for r, inicio in _posiciones(p):
        if inicio == 0 and _largo(r) >= largo:
            t = r.find(w("t"))
            if t is not None and len(t.text or "") >= largo:
                t.text = nueva + t.text[largo:]
                return True
    return False


def _parrafo_nuevo(modelo, texto):
    """Un párrafo con `texto`, con el formato de párrafo de `modelo` y el de su primer fragmento."""
    nuevo = OxmlElement("w:p")
    ppr = modelo.find(w("pPr"))
    if ppr is not None:
        nuevo.append(copy.deepcopy(ppr))
    nuevo.append(_fragmento_como(modelo, texto))
    return nuevo


def _fragmento_como(p, texto):
    """Un fragmento nuevo con `texto` y el formato del primer fragmento con texto de `p`."""
    r = OxmlElement("w:r")
    modelo = next((x for x in _runs_propios(p) if _largo(x)), None)
    if modelo is not None and modelo.find(w("rPr")) is not None:
        r.append(copy.deepcopy(modelo.find(w("rPr"))))
    t = OxmlElement("w:t")
    t.text = texto
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def _anadir_rotulo_de_nota(p):
    """Pone «Nota. » al principio del párrafo, en un fragmento nuevo."""
    r = _fragmento_como(p, ROTULO_DE_NOTA)
    primero = next(iter(_runs_propios(p)), None)
    if primero is not None:
        primero.addprevious(r)
    else:
        p.append(r)


# --- Formato --------------------------------------------------------------------------------

def _xml(el):
    return etree.tostring(el)


def formato_de_parrafo(p, conservar_con_el_siguiente, sin_sangria=True):
    """A la izquierda, sin sangría y, si se pide, pegado a lo que sigue. Devuelve si cambia algo."""
    antes = copy.deepcopy(p.find(w("pPr")))
    ppr = p.get_or_add_pPr()
    if sin_sangria:
        ppr.get_or_add_jc().set(w("val"), "left")
        ind = ppr.get_or_add_ind()
        for atributo in ("hanging", "hangingChars", "firstLineChars", "leftChars", "startChars", "start"):
            ind.attrib.pop(w(atributo), None)
        ind.set(w("left"), "0")
        ind.set(w("firstLine"), "0")
    if conservar_con_el_siguiente:
        ppr.get_or_add_keepNext().attrib.pop(w("val"), None)
    return antes is None or _xml(antes) != _xml(ppr)


def _formato_de_texto(doc, p, cursiva_hasta, negrita):
    """
    Negrita (`negrita`) o no en todo el párrafo, y cursiva solo hasta la posición
    `cursiva_hasta` (0: nada; None: todo). Devuelve si cambia algo.
    """
    estilos = Estilos(doc)
    if cursiva_hasta:
        for r, inicio in _posiciones(p):
            if inicio < cursiva_hasta < inicio + _largo(r):
                _partir(r, cursiva_hasta - inicio)
                break
    cambiado = False
    for r, inicio in _posiciones(p):
        if not "".join(t.text or "" for t in r.iter(w("t"))).strip():
            continue
        cursiva = cursiva_hasta is None or inicio < cursiva_hasta
        for etiqueta, quiere in (("i", cursiva), ("b", negrita)):
            if _activado(estilos, r, p, etiqueta) != quiere:
                _poner(r, etiqueta, quiere)
                cambiado = True
    return cambiado


def _sin_color_de_estilo(doc, p):
    """Quita el color del estilo de leyenda de Word («Descripción»), que por defecto es gris azulado."""
    estilos = Estilos(doc)
    for estilo in estilos.cadena(estilos.estilo_de(p)):
        nombre = estilos.nombre(estilo.get(w("styleId"))).lower()
        if any(c in nombre for c in ESTILOS_DE_LEYENDA):
            for color in list(estilo.iter(w("color"))):
                color.getparent().remove(color)


# --- Las leyendas ----------------------------------------------------------------------------

@dataclass
class _Hecho:
    """Lo que se ha cambiado y lo que hay que avisar, para el informe."""
    subidas: list = field(default_factory=list)
    rotulos: list = field(default_factory=list)       # (antes, después)
    separadas: list = field(default_factory=list)     # (leyenda, número, título antes, título después)
    numeradas: list = field(default_factory=list)     # (número, título)
    con_nota: list = field(default_factory=list)
    mayusculas: list = field(default_factory=list)    # (antes, después)
    dudosas: list = field(default_factory=list)
    con_formato: int = 0
    hay_notas: bool = False
    sin_mayusculas: list = field(default_factory=list)
    sin_titulo: list = field(default_factory=list)
    sin_mover: list = field(default_factory=list)
    sin_dividir: list = field(default_factory=list)
    sin_rotulo: list = field(default_factory=list)
    notas_encima: list = field(default_factory=list)
    nota_en_el_titulo: list = field(default_factory=list)


def _papeles(o, hecho):
    """Cuál de sus leyendas es el número, cuál el título y cuáles las notas. Devuelve el Parrafo del título."""
    for c in o.encima + o.debajo:
        if NOTA.match(_texto(c)):
            (o.notas if c in o.debajo else hecho.notas_encima).append(c)
    o.numero = next((c for c in o.encima + o.debajo if LEYENDA_CON_NUMERO.match(_texto(c))), None)
    lista = o.encima if o.numero in o.encima or o.numero is None else o.debajo
    resto = [c for c in lista if c is not o.numero and c not in o.notas and c not in hecho.notas_encima]
    if o.numero is None:
        titulo = resto[-1] if resto and lista is o.encima else None
    else:
        titulo = next((c for c in lista[lista.index(o.numero) + 1:] if c in resto), None)
    o.titulo = titulo.elemento if titulo is not None else None
    return titulo


def _mayusculas(p, clasificados, hecho):
    hecho.dudosas += titulos._poner_mayusculas(p, lambda x: titulos.mayusculas_apa(x, clasificados)) or []


def _leyendas_de(o, clase, informe, clasificados, titulos_tal_cual, hecho):
    titulo = _papeles(o, hecho)

    # 1. La leyenda que está debajo sube encima (con su título, si va en otro párrafo).
    if o.numero is not None and o.numero in o.debajo:
        if o.destino is None:
            hecho.sin_mover.append(_texto(o.numero))
        else:
            for c in [o.numero] + ([titulo] if titulo is not None else []):
                o.elemento.addprevious(c.elemento)
                informe.cambio_previsto(c.numero, "mover_leyenda", desde=(o.destino,))
            hecho.subidas.append(_texto(o.numero))

    if o.numero is not None:
        p = o.numero.elemento
        texto = texto_parrafo(p)
        leyenda = LEYENDA_CON_NUMERO.match(texto)
        # 2. La palabra de APA: «Cuadro 1» → «Tabla 1», «Gráfico 1» → «Figura 1».
        palabra = leyenda["palabra"]
        if palabra != clase.rotulo and PALABRAS_DE_ROTULO[palabra.lower()] == clase.rotulo:
            if _se_puede_cambiar(p, len(palabra)) and _cambiar_palabra(p, len(palabra), clase.rotulo):
                informe.cambio_previsto(o.numero.numero, "cambiar_rotulo", quitado=palabra, nuevo=clase.rotulo)
                hecho.rotulos.append((leyenda["numero"], clase.rotulo + leyenda["numero"][len(palabra):]))
                texto = texto_parrafo(p)
                leyenda = LEYENDA_CON_NUMERO.match(texto)
            else:
                hecho.sin_rotulo.append(leyenda["numero"])
        # 3. El número y el título en la misma línea se separan.
        if leyenda["titulo"]:
            if _se_puede_cambiar(p, leyenda.end("separador")) and o.titulo is None:
                o.titulo = _dividir(p, leyenda.end("numero"), leyenda.end("separador"))
                if not titulos_tal_cual:
                    _mayusculas(o.titulo, clasificados, hecho)
                informe.cambio_previsto(o.numero.numero, "dividir_leyenda", quitado=leyenda["separador"],
                                        nuevo=texto_parrafo(o.titulo))
                hecho.separadas.append((texto, leyenda["numero"], leyenda["titulo"], texto_parrafo(o.titulo)))
            else:
                hecho.sin_dividir.append(texto)
        elif titulo is None:
            hecho.sin_titulo.append(_texto(o.numero))

    # 4. Las mayúsculas de APA en el título que ya iba en su propio párrafo.
    if titulo is not None:
        antes = texto_parrafo(titulo.elemento)
        if titulos_tal_cual:
            if titulos.mayusculas_apa(antes, clasificados)[0] != antes:
                hecho.sin_mayusculas.append(antes)
        else:
            _mayusculas(titulo.elemento, clasificados, hecho)
            if texto_parrafo(titulo.elemento) != antes:
                informe.cambio_previsto(titulo.numero, "mayusculas_apa", nuevo=texto_parrafo(titulo.elemento))
                hecho.mayusculas.append((antes, texto_parrafo(titulo.elemento)))
    if o.titulo is not None and NOTA_EN_EL_TITULO.search(texto_parrafo(o.titulo)):
        hecho.nota_en_el_titulo.append(NOTA_EN_EL_TITULO.search(texto_parrafo(o.titulo)).group(0))

    # 5. «Nota.» delante de la nota que empieza por «Fuente:».
    for c in o.notas:
        if NOTA_SIN_ROTULO.match(texto_parrafo(c.elemento)):
            _anadir_rotulo_de_nota(c.elemento)
            informe.cambio_previsto(c.numero, "añadir_nota", nuevo=ROTULO_DE_NOTA)
            hecho.con_nota.append(texto_parrafo(c.elemento))


def _numerar(objetos, clase, informe, hecho):
    """El número de las que no lo tienen, si ninguna lo tiene (en un apéndice, no: paso 6.7)."""
    sin_numero = [o for o in objetos if o.numero is None and o.titulo is not None]
    if not sin_numero or any(o.numero for o in objetos):
        return sin_numero
    siguiente = 1
    for o in objetos:
        if o.numero is None and o.titulo is not None and o.zona != "anexos":
            titulo = next(c for c in o.encima if c.elemento is o.titulo)
            rotulo = f"{clase.rotulo} {siguiente}"
            nuevo = _parrafo_nuevo(o.titulo, rotulo)
            o.titulo.addprevious(nuevo)
            informe.cambio_previsto(titulo.numero, "añadir_numero", nuevo=rotulo)
            hecho.numeradas.append((rotulo, texto_parrafo(o.titulo)))
            o.numero = _NumeroNuevo(nuevo)
            siguiente += 1
    return sin_numero


def tratar(doc, opciones, informe, clasificados, objetos, clase):
    """
    Las leyendas de las tablas o figuras `objetos`: las sube, las separa, les pone
    número, «Nota.» y su formato, y anota en el informe lo que ha hecho y lo que
    hay que revisar. Devuelve cuántas han cambiado de formato.
    """
    hecho = _Hecho()
    for o in objetos:
        _leyendas_de(o, clase, informe, clasificados, opciones.get("titulos_tal_cual"), hecho)
    sin_numero = _numerar(objetos, clase, informe, hecho)

    for o in objetos:
        cambiado = False
        if o.numero is not None:
            _sin_color_de_estilo(doc, o.numero.elemento)
            cambiado |= formato_de_parrafo(o.numero.elemento, True)
            cambiado |= _formato_de_texto(doc, o.numero.elemento, 0, True)
        if o.titulo is not None:
            _sin_color_de_estilo(doc, o.titulo)
            cambiado |= formato_de_parrafo(o.titulo, True)
            cambiado |= _formato_de_texto(doc, o.titulo, None, False)
        for c in o.notas:
            hecho.hay_notas = True
            rotulo = NOTA.match(texto_parrafo(c.elemento))
            _sin_color_de_estilo(doc, c.elemento)
            cambiado |= formato_de_parrafo(c.elemento, False)
            cambiado |= _formato_de_texto(doc, c.elemento, rotulo.end(), False)
        hecho.con_formato += cambiado

    _informar(informe, clase, hecho)
    _avisar(doc, informe, clase, objetos, hecho, sin_numero)
    return hecho.con_formato


def _informar(informe, clase, hecho):
    c, s = clase.titulo, clase.singular
    if hecho.subidas:
        n = _n(len(hecho.subidas))
        informe.cambio(f"{c}: {'la leyenda' if not n else 'las leyendas'} {_lista(hecho.subidas)} estaba{n} debajo "
                       f"de la {s} y ahora va{n} encima, como pide APA.")
    if hecho.rotulos:
        informe.cambio(f"{c}: APA llama «{clase.rotulo}» a todas las {clase.plural}: "
                       + ", ".join(f"«{a}» → «{d}»" for a, d in hecho.rotulos[:3])
                       + ("…" if len(hecho.rotulos) > 3 else "") + ". Si el texto las nombra de otra forma "
                       f"(«{hecho.rotulos[0][0].lower()}»), cámbialo tú.")
    if hecho.separadas:
        texto, numero, _, titulo = hecho.separadas[0]
        con_mayusculas = any(antes != despues for _, _, antes, despues in hecho.separadas)
        informe.cambio(f"{c}: en {_plural(len(hecho.separadas), 'leyenda')} se ha separado el número del título, "
                       f"que ahora va en la línea siguiente («{_extracto(texto)}» → «{numero}» y, debajo, "
                       f"«{_extracto(titulo)}»), como pide APA."
                       + (" El título lleva además mayúscula inicial en las palabras principales (si prefieres "
                          "dejarlo como estaba, usa la opción --titulos-tal-cual)." if con_mayusculas else ""))
    if hecho.numeradas:
        n = _n(len(hecho.numeradas))
        informe.cambio(f"{c}: se ha añadido el número a {_plural(len(hecho.numeradas), s)} que no lo tenía{n} "
                       f"(«{hecho.numeradas[0][0]}» encima de «{_extracto(hecho.numeradas[0][1])}»). Comprueba que el "
                       f"texto se refiere a {'ella' if not n else 'ellas'} con ese número: APA numera las {clase.plural} "
                       "en el orden en que se mencionan.")
    if hecho.con_nota:
        n = _n(len(hecho.con_nota))
        informe.cambio(f"{c}: {'la nota' if not n else 'las notas'} de debajo de la {s} empieza{n} ahora por "
                       f"«Nota.», como pide APA ({_lista(hecho.con_nota)}).")
    if hecho.mayusculas:
        n = _n(len(hecho.mayusculas))
        informe.cambio(f"{c}: {_plural(len(hecho.mayusculas), 'título')} de {s} lleva{n} ahora mayúscula inicial en "
                       "las palabras principales, como pide APA ("
                       + ", ".join(f"«{_extracto(a)}» → «{_extracto(d)}»" for a, d in hecho.mayusculas[:3])
                       + "). Si prefieres dejarlos como estaban, usa la opción --titulos-tal-cual.")
    titulos.avisar_dudosas(informe, hecho.dudosas, f"los títulos de las {clase.plural}")
    if hecho.con_formato:
        informe.cambio(f"{c}: el número y el título de {_plural(hecho.con_formato, s)} van ahora a la izquierda y "
                       "sin sangría, el número en negrita y el título en cursiva"
                       + ("; en las notas, solo «Nota.» va en cursiva." if hecho.hay_notas else "."))


def _sin_linea_en_blanco(doc, objetos, clase):
    """Las tablas o figuras con texto pegado encima o debajo (sin un párrafo vacío en medio)."""
    bloques = _bloques(doc)
    posicion = {b: i for i, b in enumerate(bloques)}

    def con_texto(i):
        return 0 <= i < len(bloques) and bloques[i].tag == w("p") and texto_parrafo(bloques[i]).strip() != ""

    pegadas = []
    for n, o in enumerate(objetos, start=1):
        arriba = [x for x in (o.numero.elemento if o.numero else None, o.titulo) if x is not None]
        primero = arriba[0] if arriba else o.elemento
        ultimo = o.notas[-1].elemento if o.notas else o.elemento
        if con_texto(posicion[primero] - 1) or con_texto(posicion[ultimo] + 1):
            pegadas.append(_texto(o.numero).split(":")[0].split(".")[0] if o.numero else f"la {clase.singular} {n}")
    return pegadas


def _avisar(doc, informe, clase, objetos, hecho, sin_numero):
    c, s, p = clase.titulo, clase.singular, clase.plural
    if hecho.sin_mayusculas:
        n = _n(len(hecho.sin_mayusculas))
        informe.aviso(f"{_plural(len(hecho.sin_mayusculas), 'título')} de {s} no lleva{n} las mayúsculas de APA "
                      f"({_lista(hecho.sin_mayusculas)}). Se ha dejado como estaba porque has usado la opción "
                      "--titulos-tal-cual.")
    originales = any(o.numero for o in objetos if not isinstance(o.numero, _NumeroNuevo))
    if sin_numero and originales:
        n = _n(len(sin_numero))
        informe.aviso(f"{_plural(len(sin_numero), s)} no lleva{n} número "
                      f"({_lista([texto_parrafo(o.titulo) for o in sin_numero])}). Como otras sí lo llevan, no se ha "
                      f"añadido, para no descuadrar la numeración: ponlo tú encima del título («{clase.rotulo} 2»…).")
    anexos = [o for o in sin_numero if o.zona == "anexos" and o.numero is None]
    if anexos and not originales:
        informe.aviso(f"{_plural(len(anexos), s)} de un apéndice no lleva{_n(len(anexos))} número: APA las numera "
                      f"con la letra del apéndice («{clase.rotulo} A1»). Ponlo tú.")
    sin_nada = [n for n, o in enumerate(objetos, start=1) if o.numero is None and o.titulo is None]
    if sin_nada:
        cuales = f"La {s} {sin_nada[0]}" if len(sin_nada) == 1 else f"Las {p} {', '.join(map(str, sin_nada))}"
        informe.aviso(f"{cuales} no lleva{_n(len(sin_nada))} número ni título. APA pone encima de cada {s} "
                      f"«{clase.rotulo} N» en negrita y, debajo, su título en cursiva: escríbelos tú.")
    if hecho.sin_titulo:
        informe.aviso(f"{_lista(hecho.sin_titulo)} no tiene{_n(len(hecho.sin_titulo))} título: APA lo pone en la "
                      "línea siguiente al número, en cursiva. Escríbelo tú.")
    if hecho.sin_dividir:
        informe.aviso(f"En {_lista(hecho.sin_dividir)} el número y el título van en la misma línea, pero la leyenda "
                      "lleva un campo de Word o cambios marcados y no se ha podido separar: pon tú el título en la "
                      "línea siguiente.")
    if hecho.sin_rotulo:
        informe.aviso(f"APA llama «{clase.rotulo}» a todas las {p} ({_lista(hecho.sin_rotulo)}), pero la leyenda "
                      "lleva un campo de Word o cambios marcados y no se ha podido cambiar: cámbialo tú.")
    if hecho.sin_mover:
        informe.aviso(f"La leyenda {_lista(hecho.sin_mover)} está debajo de la {s} y no se ha podido subir: APA la "
                      "pone encima. Muévela tú.")
    if hecho.notas_encima:
        informe.aviso(f"La nota {_lista([_texto(x) for x in hecho.notas_encima])} está encima de la {s}: APA pone "
                      "las notas debajo. Muévela tú.")
    if hecho.nota_en_el_titulo:
        informe.aviso(f"El título de {_plural(len(hecho.nota_en_el_titulo), s)} lleva "
                      f"{_lista(hecho.nota_en_el_titulo)}: APA lo pone en una nota debajo de la {s} («Nota. "
                      "Elaboración propia.»). Se ha dejado en el título: muévelo tú si quieres.")
    numeros = [int(m.group(1)) for o in objetos if o.numero is not None and o.zona != "anexos"
               for m in [re.search(r"\s(\d+)$", LEYENDA_CON_NUMERO.match(texto_parrafo(o.numero.elemento))["numero"])]
               if m]
    if numeros and numeros != list(range(1, len(numeros) + 1)):
        informe.aviso(f"La numeración de las {p} no va seguida ({', '.join(map(str, numeros))}): APA las numera "
                      "1, 2, 3… en el orden en que se mencionan en el texto. Revísala tú.")
    pegadas = _sin_linea_en_blanco(doc, objetos, clase)
    if pegadas:
        informe.aviso(f"{_lista(pegadas)} no tiene{_n(len(pegadas))} una línea en blanco entre la {s} y el texto: "
                      "APA la recomienda para separarlas. Añádela tú con un párrafo vacío encima del número y "
                      f"debajo de la {s} (o de su nota).")
    recordar_posicion(informe, clase)


RECORDATORIO = ("Comprueba en Word que cada {} queda arriba o abajo de la página, y no en medio del texto, como "
                "recomienda APA (también puede ir en una página para ella sola). La herramienta no puede saberlo: "
                "depende de cómo reparte Word las páginas.")


def recordar_posicion(informe, clase):
    """Un solo recordatorio para tablas y figuras: si ya está el de las tablas, se amplía."""
    for entrada in informe.entradas:
        if entrada.tipo == "aviso" and entrada.mensaje == RECORDATORIO.format(TABLA.singular) and clase is FIGURA:
            entrada.mensaje = RECORDATORIO.format("tabla y cada figura")
            return
    informe.aviso(RECORDATORIO.format(clase.singular))
