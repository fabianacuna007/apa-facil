"""
Paso 7.7 (decisión del propietario, 2026-10-19): corrige en la copia los errores
de forma de las citas y de las referencias que el paso 7.4 sabe corregir sin
inventar nada. Por defecto; con la opción --citas-tal-cual solo se avisa.

Se corrige, con las reglas de la web oficial de APA (Parenthetical versus
narrative citations, Author–date citation system, Elements of reference list
entries) y la versión en español de la especificación del ROADMAP:
- en las citas: la coma antes del año, «et al.», «s. f.» y «en prensa», «&» por
  «y», el espacio de «p. 45» y «p.»/«pp.»;
- en las referencias: los autores separados por «;» o con «&» (comas y «y»), el
  DOI antiguo (→ https://doi.org/…), la ciudad de la editorial (se quita),
  «Recuperado de» sin fecha de consulta (se quita), el año sin paréntesis justo
  detrás del autor, el punto después del autor y de la fecha, y «nº 7» (→ «12(7)»).

No se corrige, y sigue siendo un aviso, lo que obligaría a inventar o a decidir
por el estudiante: las referencias incompletas, el año que falta, las que no se
han podido leer, el cruce de citas y referencias (7.3) y lo que está en las notas
al pie (python-docx no las carga como XML). Tampoco un error cuyo texto ya no está
tal cual en la copia, ni uno que pisa a otro.

Cada corrección cambia solo lo que difiere entre lo escrito y lo correcto, con el
formato del primer fragmento que toca: la cursiva del título de una revista, por
ejemplo, se conserva. Va en el informe como cambio y como cambio previsto
«corregir_forma», que la comprobación de integridad acepta solo si no quita más
de unas pocas palabras y solo añade las de APA (y, s. f., en prensa,
https://doi.org/, p., pp.: documento.correccion_de_forma_permitida).
"""

import re
from collections import defaultdict
from dataclasses import dataclass

from apa_facil.documento import (_fragmentos_del_parrafo, _piezas, aislar,
                                 correccion_de_forma_permitida, texto_parrafo, w)
from apa_facil.revision import forma

EN_CITAS = (forma.COMA, forma.ET_AL, forma.SIN_FECHA, forma.AMPERSAND, forma.ESPACIO_EN_PAGINA, forma.P_O_PP)
EN_REFERENCIAS = (forma.AMPERSAND, forma.PUNTO_Y_COMA, forma.DOI_ANTIGUO, forma.CIUDAD, forma.RECUPERADO_DE,
                  forma.AÑO_SIN_PARENTESIS, forma.PUNTO_TRAS_AUTOR, forma.PUNTO_TRAS_FECHA,
                  forma.NUMERO_ABREVIADO)

# Lo que se dice en el informe de cada corrección.
QUE_SE_HA_HECHO = {
    (forma.CITA, forma.COMA): "se ha añadido la coma entre el autor y el año",
    (forma.CITA, forma.ET_AL): "se ha escrito «et al.» como pide APA («et» sin punto y «al.» con punto)",
    (forma.CITA, forma.SIN_FECHA): "se ha escrito la fecha como pide APA en español («s. f.», «en prensa»)",
    (forma.CITA, forma.AMPERSAND): "se ha cambiado «&» por «y», como pide APA en español",
    (forma.CITA, forma.ESPACIO_EN_PAGINA): "se ha añadido el espacio entre «p.» y el número de página",
    (forma.CITA, forma.P_O_PP): "se ha puesto «p.» para una página y «pp.» para varias",
    (forma.REFERENCIA, forma.AMPERSAND): "se ha unido el último autor con «y», como pide APA en español",
    (forma.REFERENCIA, forma.PUNTO_Y_COMA): "se han separado los autores con comas, y el último con «y»",
    (forma.REFERENCIA, forma.DOI_ANTIGUO): "se ha escrito el DOI como enlace (https://doi.org/…)",
    (forma.REFERENCIA, forma.CIUDAD): "se ha quitado la ciudad de la editorial, que APA 7 ya no pone",
    (forma.REFERENCIA, forma.RECUPERADO_DE): "se ha quitado «Recuperado de» delante del enlace",
    (forma.REFERENCIA, forma.AÑO_SIN_PARENTESIS): "se ha puesto el año entre paréntesis, seguido de un punto",
    (forma.REFERENCIA, forma.PUNTO_TRAS_AUTOR): "se ha añadido el punto después del autor",
    (forma.REFERENCIA, forma.PUNTO_TRAS_FECHA): "se ha añadido el punto después de la fecha",
    (forma.REFERENCIA, forma.NUMERO_ABREVIADO):
        "se ha puesto el número de la revista entre paréntesis justo detrás del volumen, sin «nº»",
}
EJEMPLOS = 3

@dataclass
class _Cambio:
    inicio: int          # en el texto actual del párrafo
    fin: int
    nuevo: str
    errores: list        # los errores que corrige (en una cita pueden ser varios)
    muestra: tuple       # (antes, después), para el informe


def corregir(revision, informe):
    """
    Corrige en el documento (en su sitio: los párrafos son los mismos objetos que
    en el original) los errores de forma que se pueden corregir, y los anota en el
    informe. Devuelve los id() de los errores corregidos, que ya no se avisan.
    """
    corregidos, hechos = set(), defaultdict(list)   # (en, tipo) → [(antes, después)]
    for elemento, cambios in _cambios(revision).items():
        # De derecha a izquierda: así los cambios no mueven el sitio de los que quedan por hacer.
        for cambio in sorted(_sin_solapes(cambios), key=lambda c: -c.inicio):
            texto = texto_parrafo(elemento)
            inicio, fin, nuevo = _ampliado(texto, cambio.inicio, cambio.fin, cambio.nuevo)
            quitado = texto[inicio:fin]
            if not correccion_de_forma_permitida(quitado, nuevo) or not _sustituir(elemento, inicio, fin, quitado, nuevo):
                continue
            informe.cambio_previsto(cambio.errores[0].numero, "corregir_forma", quitado=quitado, nuevo=nuevo,
                                    datos=texto[:inicio].count(quitado))
            for e in cambio.errores:
                corregidos.add(id(e))
                hechos[(e.en, e.tipo)].append(cambio.muestra)

    for (en, tipo), muestras in hechos.items():
        muestras = list(dict.fromkeys(muestras))
        if en == forma.CITA:
            quien, donde = "Citas", "cita" if len(muestras) == 1 else "citas"
        else:
            quien, donde = "Referencias", "referencia" if len(muestras) == 1 else "referencias"
        ejemplos = "; ".join(f"«{_corto(a)}» → «{_corto(d)}»" if d.strip() else f"«{_corto(a)}», quitado"
                             for a, d in muestras[:EJEMPLOS])
        mas = f" y {len(muestras) - EJEMPLOS} más" if len(muestras) > EJEMPLOS else ""
        informe.cambio(f"{quien}: {QUE_SE_HA_HECHO[(en, tipo)]} en {len(muestras)} {donde}: {ejemplos}{mas}.")
    return corregidos


def _corto(texto, largo=70):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1].rstrip() + "…"


# --- Qué cambiar ----------------------------------------------------------------------

def _cambios(revision):
    """Para cada párrafo (el w:p), los cambios que hay que hacer en su texto actual."""
    cambios = defaultdict(list)
    citas = defaultdict(list)   # (párrafo, inicio de la cita) → sus errores
    for e in revision.errores:
        if e.numero is None or (e.correccion is None and e.tipo != forma.CIUDAD):
            continue   # en una nota al pie, o sin corrección propuesta (la ciudad sola se quita sin más)
        elemento = revision.lugares.elementos.get(e.numero)
        if elemento is None:
            continue
        if e.en == forma.CITA and e.tipo in EN_CITAS:
            citas[(elemento, e.origen.inicio)].append(e)
        elif e.en == forma.REFERENCIA and e.tipo in EN_REFERENCIAS:
            cambio = _en_la_referencia(e, texto_parrafo(elemento))
            if cambio:
                cambios[elemento].append(cambio)
    for (elemento, _), errores in citas.items():
        cambio = _en_la_cita(errores, texto_parrafo(elemento))
        if cambio:
            cambios[elemento].append(cambio)
    return cambios


def _en_la_cita(errores, texto):
    """
    Todos los errores de una cita, juntos: cada corrección del 7.4 se reduce a lo
    que cambia, y se aplican todas si no se pisan («(Pérez et. al 2020)» →
    «(Pérez et al., 2020)»).
    """
    cita = errores[0].origen
    fragmento = cita.fragmento
    # Las reglas pueden haber movido la cita (la limpieza quita el tabulador del principio):
    # se busca la aparición más cercana a donde estaba.
    sitios = [m.start() for m in re.finditer(re.escape(fragmento), texto)]
    if not sitios:
        return None   # ya no está tal cual: no se toca
    desde = min(sitios, key=lambda s: abs(s - cita.inicio))
    trozos = sorted((_diferencia(fragmento, e.correccion) for e in errores), key=lambda t: t[0])
    resultado, hasta = [], 0
    for inicio, fin, nuevo in trozos:
        if inicio < hasta:
            return None   # dos correcciones se pisan: mejor avisar
        resultado += [fragmento[hasta:inicio], nuevo]
        hasta = fin
    corregido = "".join(resultado) + fragmento[hasta:]
    inicio, fin, nuevo = _diferencia(fragmento, corregido)
    return _Cambio(desde + inicio, desde + fin, nuevo, errores, (fragmento, corregido))


def _en_la_referencia(e, texto):
    """El cambio de un error de una referencia en el texto actual de su párrafo, o None si no se encuentra."""
    sitio = _sitio_en_la_referencia(e, texto)
    if sitio is None:
        return None
    (desde, hasta), correcto = sitio
    inicio, fin, nuevo = _diferencia(texto[desde:hasta], correcto)
    return _Cambio(desde + inicio, desde + fin, nuevo, [e], (texto[desde:hasta], correcto))


def _sitio_en_la_referencia(e, texto):
    """((desde, hasta), lo correcto): el trozo del texto actual que se cambia, y por qué."""
    r = e.origen
    if e.tipo in (forma.AMPERSAND, forma.PUNTO_Y_COMA):
        sitio = _unico(texto, re.escape(e.escrito))
        return (sitio, e.correccion) if sitio else None
    if e.tipo == forma.DOI_ANTIGUO:
        sitio = _ultimo(texto, re.escape(e.escrito))
        return (sitio, e.correccion) if sitio else None
    if e.tipo == forma.CIUDAD and r.fuente:
        # «Madrid: Editorial Síntesis» o «Editorial Síntesis, Madrid» → «Editorial Síntesis»
        sitio = _unico(texto, re.escape(r.lugar) + r"\s*[:,]\s*" + re.escape(r.fuente))
        if sitio is None:
            sitio = _unico(texto, re.escape(r.fuente) + r"\s*,\s*" + re.escape(r.lugar) + r"(?=\.?\s*$)")
        return (sitio, r.fuente) if sitio else None
    if e.tipo == forma.CIUDAD:   # sin editorial (la organización autora): «…territorial. Madrid.» → «…territorial.»
        sitio = _unico(texto, r"(?<=\.)\s+" + re.escape(r.lugar) + r"\.?(?=\s*$)")
        return (sitio, "") if sitio else None
    if e.tipo == forma.RECUPERADO_DE and (r.url or r.doi):
        enlace = r.url or r.doi
        sitio = _ultimo(texto, re.escape(r.recuperado_de) + r"\s+" + re.escape(enlace))
        return (sitio, enlace) if sitio else None
    if e.tipo == forma.AÑO_SIN_PARENTESIS and r.año_tras_el_autor and r.autores_escritos:
        # «Lefèvre, A. 2009.» → «Lefèvre, A. (2009).»
        patron = re.escape(r.autores_escritos.rstrip(".")) + r"\.?\s*,?\s*" + re.escape(r.año) + r"[.,]?"
        sitio = _unico(texto, patron)
        if sitio is None or texto[:sitio[0]].strip():
            return None   # tiene que ser el principio de la referencia
        return sitio, e.correccion
    if e.tipo == forma.PUNTO_TRAS_AUTOR:   # «Consejo Económico (2022)» → «Consejo Económico. (2022)»
        fecha = forma.FECHA.search(texto)
        if fecha is None:
            return None
        antes = texto[:fecha.start()].rstrip()
        if not antes or antes.endswith("."):
            return None
        return (0, fecha.end()), antes + "." + texto[len(antes):fecha.end()]
    if e.tipo == forma.PUNTO_TRAS_FECHA:   # «(2019) Título» o «(2019), Título» → «(2019). Título»
        fecha = forma.FECHA.search(texto)
        if fecha is None or texto[fecha.end():].lstrip().startswith("."):
            return None
        signo = re.match(r"[,;:]?", texto[fecha.end():]).group()
        return (fecha.start(), fecha.end() + len(signo)), fecha.group() + "."
    if e.tipo == forma.NUMERO_ABREVIADO and r.volumen:   # «12, nº 7» → «12(7)»
        sitio = _unico(texto, re.escape(r.volumen) + r"\s*,\s*" + re.escape(e.escrito))
        return (sitio, e.correccion) if sitio else None
    return None


def _unico(texto, patron):
    """Dónde está el patrón en el texto, si está una sola vez."""
    encontrados = list(re.finditer(patron, texto))
    return encontrados[0].span() if len(encontrados) == 1 else None


def _ultimo(texto, patron):
    encontrados = list(re.finditer(patron, texto))
    return encontrados[-1].span() if encontrados else None


def _diferencia(antes, despues):
    """Lo que cambia de `antes` a `despues`: (inicio, fin, nuevo), sin lo que tienen igual al principio y al final."""
    prefijo = 0
    while prefijo < min(len(antes), len(despues)) and antes[prefijo] == despues[prefijo]:
        prefijo += 1
    sufijo = 0
    while (sufijo < min(len(antes), len(despues)) - prefijo
           and antes[len(antes) - 1 - sufijo] == despues[len(despues) - 1 - sufijo]):
        sufijo += 1
    return prefijo, len(antes) - sufijo, despues[prefijo:len(despues) - sufijo]


def _sin_solapes(cambios):
    """Los cambios que no se pisan (si dos se pisan, o añaden en el mismo sitio, se queda el primero)."""
    elegidos = []
    for cambio in sorted(cambios, key=lambda c: (c.inicio, c.fin)):
        if elegidos and (cambio.inicio < elegidos[-1].fin or cambio.inicio == elegidos[-1].inicio):
            continue
        elegidos.append(cambio)
    return elegidos


def _ampliado(texto, inicio, fin, nuevo):
    """Un añadido se hace junto a la letra de delante (o la de detrás), para que tenga su formato."""
    if inicio < fin:
        return inicio, fin, nuevo
    if inicio > 0:
        return inicio - 1, fin, texto[inicio - 1] + nuevo
    if fin < len(texto):
        return inicio, fin + 1, nuevo + texto[fin]
    return inicio, fin, nuevo


# --- Cómo se cambia -------------------------------------------------------------------

def _sustituir(p, inicio, fin, quitado, nuevo):
    """
    Cambia texto[inicio:fin] (que es `quitado`, no vacío) por `nuevo` en el
    párrafo, con el formato del primer fragmento que toca. Solo si lo que se toca
    es texto (ni tabuladores, ni saltos, ni notas, ni imágenes). Devuelve si se ha hecho.
    """
    texto = texto_parrafo(p)
    if inicio >= fin or texto[inicio:fin] != quitado:
        return False
    trozo = aislar(p, inicio, fin)
    if trozo is None:
        return False
    fragmentos = _fragmentos_del_parrafo(p)
    tocados = fragmentos[fragmentos.index(trozo[0]):fragmentos.index(trozo[1]) + 1]
    if any(_etiqueta(el) != "t" for r in tocados for el in _piezas(r)):
        return False
    for r in tocados:
        for el in _piezas(r):
            r.remove(el)
    if nuevo:
        t = tocados[0].makeelement(w("t"), {})
        t.text = nuevo
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        tocados[0].append(t)
    for r in tocados:
        if not _piezas(r):
            r.getparent().remove(r)
    return texto_parrafo(p) == texto[:inicio] + nuevo + texto[fin:]


def _etiqueta(el):
    return el.tag.rsplit("}", 1)[-1]
