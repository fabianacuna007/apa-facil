"""
Paso 7.5: la revisión de citas y referencias, en el informe.

Junta los pasos 7.1 a 7.4 (detector de citas, lector de referencias, cruce y
errores de forma) y los convierte en avisos para el estudiante. No cambia nada
del documento. Cada aviso es un tipo de problema («falta la coma antes del
año») y lleva sus casos: dónde está cada uno, lo escrito y cómo se corrige.

Dónde está, en palabras, porque Word no numera los párrafos (y en la copia los
números cambian: se borran vacíos, se sustituye la portada, se reordenan las
referencias):
- en el texto: el apartado y el principio del párrafo;
- en una tabla: su orden en el trabajo («la 1.ª tabla») y su apartado;
- en una nota: «la nota al pie 1»;
- en la lista de referencias: el principio de la referencia.
El número de párrafo del original va además en el JSON (para la web y el 7.6).

La revisión se hace sobre el original, antes de las reglas (`analizar`), y se
pasa al informe al final (`avisar`), detrás de los avisos de formato. No se
repite lo que ya avisa el paso 6.5: el orden alfabético, las referencias que son
solo una dirección web y las obras con los mismos autores y año sin letra.
"""

from dataclasses import dataclass, field

from apa_facil.documento import texto_parrafo, w
from apa_facil.reglas.orden import SOLO_UNA_URL
from apa_facil.revision import citas as detector
from apa_facil.revision import cruce as cruzador
from apa_facil.revision import forma as revisor
from apa_facil.revision import referencias as lector

PALABRAS_DEL_PRINCIPIO = 6
LETRAS_DE_LA_REFERENCIA = 45


@dataclass
class Revision:
    """Lo que se ha encontrado en el original, listo para pasarlo al informe."""
    citas: list = field(default_factory=list)
    referencias: list = field(default_factory=list)
    cruce: object = None
    errores: list = field(default_factory=list)
    lugares: object = None
    # Paso 7.6: (mensaje del aviso, caso, Ancla) de cada caso, para poner sus comentarios en la copia.
    anclas: list = field(default_factory=list)
    # Paso 7.7: los id() de los errores de forma ya corregidos en la copia, que no se avisan.
    corregidos: set = field(default_factory=set)


@dataclass
class Ancla:
    """Dónde poner el comentario de un caso: su párrafo (el objeto, que sigue siendo el mismo en la copia) y qué texto marcar."""
    elemento: object            # el w:p; None en las notas al pie (no llevan comentario)
    buscar: str = None          # el texto que se marca; None: toda la referencia
    cerca: int = None           # dónde estaba en el original (para elegir si aparece varias veces)
    ultima: bool = False        # si aparece varias veces y no hay `cerca`, la última (la ciudad, el DOI…)
    referencia: bool = False    # en una referencia: si no se encuentra el texto, se marca la referencia entera


def analizar(doc, clasificacion):
    """Revisa las citas y las referencias del original (antes de aplicar las reglas)."""
    citas = detector.detectar(doc, clasificacion)
    referencias = lector.leer(doc, clasificacion)
    return Revision(citas, referencias, cruzador.cruzar(citas, referencias), revisor.revisar(citas, referencias),
                    _Lugares(doc, clasificacion))


def revisar(doc, clasificacion, informe):
    """Analiza y avisa de una vez (para cuando no hay reglas en medio)."""
    avisar(analizar(doc, clasificacion), informe)


# --- Dónde está cada cosa -------------------------------------------------------------

class _Lugares:
    """Para cada párrafo del original: su apartado y cómo empieza, o en qué tabla está."""

    def __init__(self, doc, clasificacion):
        clasificados = {c.elemento: c for c in clasificacion}
        self.parrafos = {}   # número → (apartado, principio, nº de tabla o None)
        self.elementos = {}  # número → el w:p (paso 7.6)
        tablas, apartado = {}, None
        for numero, p in enumerate(doc.element.body.iter(w("p")), start=1):
            c = clasificados.get(p)
            if c is not None and c.zona != "portada" and c.tipo in ("titulo", "titulo_referencias", "anexo"):
                apartado = " ".join(c.texto.split()).rstrip(":")
            tabla = next(p.iterancestors(w("tbl")), None)
            if tabla is not None:
                for exterior in tabla.iterancestors(w("tbl")):
                    tabla = exterior   # una tabla dentro de otra cuenta como la de fuera
                tablas.setdefault(tabla, len(tablas) + 1)
            self.elementos[numero] = p
            self.parrafos[numero] = (apartado, _principio(texto_parrafo(p)),
                                     tablas[tabla] if tabla is not None else None)

    def de_cita(self, cita):
        if cita.numero is None:   # en una nota
            return f"En la {cita.ubicacion[0].lower()}{cita.ubicacion[1:]}"
        apartado, principio, tabla = self.parrafos.get(cita.numero, (None, "", None))
        donde_apartado = f"apartado «{apartado}»" if apartado else None
        if tabla:
            return f"En la {tabla}.ª tabla" + (f" ({donde_apartado})" if donde_apartado else "")
        parrafo = f"en el párrafo que empieza «{principio}»"
        return f"En «{apartado}», {parrafo}" if apartado else parrafo[0].upper() + parrafo[1:]

    @staticmethod
    def de_referencia(referencia):
        texto = " ".join(referencia.texto.split())
        corto = texto if len(texto) <= LETRAS_DE_LA_REFERENCIA else texto[:LETRAS_DE_LA_REFERENCIA - 1].rstrip() + "…"
        return f"En la lista de referencias: «{corto}»"


def _principio(texto):
    palabras = texto.split()
    if len(palabras) <= PALABRAS_DEL_PRINCIPIO:
        return " ".join(palabras)
    return " ".join(palabras[:PALABRAS_DEL_PRINCIPIO]).rstrip(",;:") + "…"


def _caso(donde, extracto, correccion=None, explicacion=None, parrafo=None):
    return {"donde": donde, "extracto": " ".join(extracto.split()), "correccion": correccion,
            "explicacion": explicacion, "parrafo": parrafo}


# --- Los avisos -----------------------------------------------------------------------

def avisar(revision, informe):
    """Pasa la revisión al informe: primero el cruce, luego la forma de las citas y la de las referencias."""
    lugares = revision.lugares
    cruce = revision.cruce
    años_de_las_citas = {}   # referencia sin año → años con que la cita el texto (para «falta el año»)
    for h in cruce.años_distintos:
        for r in h.referencias:
            if r.año is None:
                años_de_las_citas.setdefault(id(r), []).append(h.citas[0].año)

    def casos_de_citas(hallazgos, con_detalle=False):
        return [(_caso(lugares.de_cita(c), c.fragmento, explicacion=h.detalle if con_detalle else None,
                       parrafo=c.numero), _ancla_de_cita(c, lugares))
                for h in hallazgos for c in h.citas]

    def aviso(mensaje, pares):
        informe.aviso(mensaje, [caso for caso, _ in pares])
        revision.anclas += [(mensaje, caso, ancla) for caso, ancla in pares]

    if cruce.sin_referencia:
        aviso(f"{_obras(len(cruce.sin_referencia), 'citada en el texto no está', 'citadas en el texto no están')} "
                      "en la lista de referencias. APA pide que cada obra citada tenga su referencia: añádela, o revisa "
                      "si el autor y el año coinciden con los de alguna de la lista.",
                      casos_de_citas(cruce.sin_referencia, con_detalle=True))
    if cruce.no_citadas:
        aviso(f"{_obras(len(cruce.no_citadas), 'de la lista de referencias no se cita', 'de la lista de referencias no se citan')} "
                      "en el texto. APA pide que cada referencia se cite en el texto (o en una tabla, figura, nota o "
                      "apéndice): si no has usado esa obra, quítala de la lista.",
                      [(_caso(lugares.de_referencia(r), cruzador.referencia_corta(r), parrafo=r.numero),
                        Ancla(lugares.elementos.get(r.numero), referencia=True))
                       for h in cruce.no_citadas for r in h.referencias])
    años = [h for h in cruce.años_distintos if any(r.año is not None for r in h.referencias)]
    if años:   # las referencias sin año van en «falta el año», con el año de la cita como pista
        aviso("El año de la cita no coincide con el de su referencia: APA pide que sea el mismo en el texto y "
                      "en la lista.", casos_de_citas(años, con_detalle=True))
    if cruce.autores_distintos:
        aviso("Los autores de la cita no están escritos como en su referencia: APA pide que coincidan "
                      "(apellidos, tildes y número de autores).", casos_de_citas(cruce.autores_distintos, True))
    ambiguas = [h for h in cruce.ambiguas if not h.detalle.startswith("hay varias referencias con los mismos")]
    if ambiguas:   # las de los mismos autores y año sin letra ya las avisa el paso 6.5
        aviso("La cita puede corresponder a varias referencias de la lista: APA pide que cada cita lleve a una "
                      "sola.", casos_de_citas(ambiguas, con_detalle=True))

    errores = [e for e in revision.errores if _hay_que_avisar(e) and id(e) not in revision.corregidos]
    for en in (revisor.CITA, revisor.REFERENCIA):
        tipos = list(dict.fromkeys(e.tipo for e in errores if e.en == en))
        for tipo in tipos:
            del_tipo = [e for e in errores if e.en == en and e.tipo == tipo]
            # La explicación va en el mensaje si es la misma para todos; si no, en cada caso.
            explicaciones = {e.explicacion for e in del_tipo}
            comun = explicaciones.pop() if len(explicaciones) == 1 else None
            pares = [(_caso_de_error(e, lugares, años_de_las_citas, comun), _ancla_de_error(e, lugares))
                     for e in del_tipo]
            donde = "en las citas del texto" if en == revisor.CITA else "en la lista de referencias"
            aviso(f"{revisor.TITULOS[tipo]} {donde}." + (f" {comun}" if comun else ""), pares)


def _hay_que_avisar(error):
    if error.tipo == revisor.ORDEN:
        return False   # lo hace el paso 6.5 (y por defecto reordena la copia)
    if error.tipo == revisor.NO_RECONOCIDA and SOLO_UNA_URL.match(error.origen.texto):
        return False   # el 6.5 ya avisa de las referencias que son solo una dirección web
    return True


def _caso_de_error(error, lugares, años_de_las_citas, comun):
    donde = lugares.de_cita(error.origen) if error.en == revisor.CITA else lugares.de_referencia(error.origen)
    correccion = error.correccion
    explicacion = None if error.explicacion == comun else error.explicacion
    citado = años_de_las_citas.get(id(error.origen)) if error.tipo == revisor.FALTA_AÑO else None
    if citado:   # la pista sale del propio texto: el año con que se cita la obra
        correccion = f"{error.origen.autores_escritos.rstrip('.')}. ({citado[0]})."
        explicacion = f"en el texto la citas como de {_y(list(dict.fromkeys(citado)))}"
    extracto = error.escrito
    if error.tipo in (revisor.INCOMPLETA, revisor.NO_RECONOCIDA):   # lo escrito es la referencia entera
        donde, extracto = "En la lista de referencias", error.origen.texto
    return _caso(donde, extracto, correccion, explicacion, error.numero)


def _ancla_de_cita(cita, lugares):
    elemento = lugares.elementos.get(cita.numero) if cita.numero is not None else None
    return Ancla(elemento, cita.fragmento, cita.inicio)


# Lo que va al final de la referencia: si el texto se repite, se marca la última vez.
AL_FINAL = (revisor.CIUDAD, revisor.DOI_ANTIGUO, revisor.RECUPERADO_DE)


def _ancla_de_error(error, lugares):
    if error.en == revisor.CITA:
        return _ancla_de_cita(error.origen, lugares)
    buscar = None if error.tipo in (revisor.INCOMPLETA, revisor.NO_RECONOCIDA) else error.escrito
    return Ancla(lugares.elementos.get(error.numero), buscar, ultima=error.tipo in AL_FINAL, referencia=True)


def _y(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


def _obras(n, singular, plural):
    return f"Una obra {singular}" if n == 1 else f"{n} obras {plural}"
