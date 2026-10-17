"""
Verificador: comprueba si un documento Word cumple las reglas APA de la
fase 1 y lista cada incumplimiento con su ubicación (sección, número de
párrafo y un extracto del texto). No modifica nada.

Cómo comprueba:
- Lee el XML del documento y calcula el formato efectivo de cada párrafo y de
  cada fragmento de texto (valores por defecto → estilos → formato directo).
  No pregunta a Word, que da "fuente mezclada" en rangos grandes aunque no lo
  estén (ver "Problemas abiertos" en ROADMAP.md).
- Respeta las mismas excepciones que las reglas (portada, código, párrafos con
  tabulador…), usando los mismos detectores. Lo que entra por excepción no es
  un incumplimiento: se lista aparte, como "excepción admitida".
- Qué párrafos deben llevar sangría lo decide el clasificador de párrafos
  (apa_facil/estructura.py), igual que en la regla. El verificador solo tiene
  la copia, así que la clasifica a ella; el programa, en cambio, clasificó el
  original. Pueden diferir en títulos que solo se distinguían por el tamaño de
  letra, que la regla de fuente iguala: en los 5 documentos de prueba coincide
  qué párrafos son de texto, y si alguna vez no, aparecería como incumplimiento.
  La fase 4 lo resuelve al dar estilos de Word a los títulos.
- Limitación: si el clasificador se equivoca (por ejemplo, toma un párrafo de
  texto por un título), el verificador no lo ve. Para eso están las pruebas del
  clasificador con la clasificación esperada de los documentos de prueba.

Uso:
    python herramientas/verificar.py archivo.docx [otro.docx ...]

Termina con código 0 si todos los documentos cumplen y 1 si alguno no.
"""

import sys
from dataclasses import dataclass, field
from pathlib import Path

# Para encontrar el paquete apa_facil al ejecutar este archivo directamente.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apa_facil import documento  # noqa: E402
from apa_facil.__main__ import configurar_terminal  # noqa: E402
from apa_facil.documento import ErrorDocumento, Estilos, cm, en_ecuacion, texto_parrafo, w  # noqa: E402
from apa_facil.reglas import alineacion, fuente, interlineado, margenes, paginacion, sangria  # noqa: E402

MARGENES = "1.1 Márgenes"
FUENTE = "1.2 Fuente"
INTERLINEADO = "1.3 Interlineado"
ALINEACION = "1.4 Alineación"
SANGRIA = "1.5 Sangría"
PAGINACION = "1.6 Número de página"
REGLAS = [MARGENES, FUENTE, INTERLINEADO, ALINEACION, SANGRIA, PAGINACION]

MAXIMO_POR_REGLA = 8   # incumplimientos que se muestran de cada regla; del resto, solo cuántos son


@dataclass
class Hallazgo:
    regla: str
    mensaje: str
    donde: str              # ubicación legible: "Sección 1 · párrafo 12 (en una tabla)"
    extracto: str = ""
    seccion: int = None
    parrafo: int = None     # número de párrafo del cuerpo, contando los de las tablas


@dataclass
class Resultado:
    archivo: str
    incumplimientos: list = field(default_factory=list)
    excepciones: list = field(default_factory=list)

    @property
    def cumple(self):
        return not self.incumplimientos


# --- Utilidades -------------------------------------------------------------------

def _extracto(texto, largo=40):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo] + "…"


def _activado(el):
    """¿Está activado un ajuste de Word como <w:titlePg/> o <w:autoHyphenation w:val="1"/>?"""
    return el is not None and el.get(w("val")) not in ("0", "false", "off")


def _entero(valor):
    try:
        return int(valor)
    except (TypeError, ValueError):
        return 0


def _pt(puntos):
    return f"{puntos:g}".replace(".", ",")


def _dentro_de(p, *etiquetas):
    return next(p.iterancestors(*etiquetas), None) is not None


@dataclass
class Parrafo:
    """Un párrafo del cuerpo y dónde está."""
    p: object
    numero: int
    seccion: int
    en_tabla: bool
    en_cuadro: bool

    @property
    def donde(self):
        lugar = " (en un cuadro de texto)" if self.en_cuadro else " (en una tabla)" if self.en_tabla else ""
        return f"Sección {self.seccion} · párrafo {self.numero}{lugar}"


class Verificacion:
    """Recorre un documento abierto y anota incumplimientos y excepciones."""

    def __init__(self, doc, archivo):
        self.doc = doc
        self.estilos = Estilos(doc)
        self.resultado = Resultado(archivo)
        # Todos los párrafos del cuerpo, numerados como en el resto del programa
        # (radiografía, informe de la fuente…), con su sección.
        self.parrafos, numero = [], 0
        for seccion, bloques in enumerate(documento.bloques_por_seccion(doc), start=1):
            for bloque in bloques:
                for p in bloque.iter(w("p")):
                    numero += 1
                    self.parrafos.append(Parrafo(p, numero, seccion, _dentro_de(p, w("tc")),
                                                 _dentro_de(p, w("txbxContent"))))
        self.por_elemento = {info.p: info for info in self.parrafos}
        # Los del texto principal: ni en tablas ni en cuadros de texto (como en las reglas).
        self.principales = [i for i in self.parrafos if not i.en_tabla and not i.en_cuadro]

    def fallo(self, regla, mensaje, donde, extracto="", seccion=None, parrafo=None):
        self.resultado.incumplimientos.append(Hallazgo(regla, mensaje, donde, extracto, seccion, parrafo))

    def excepcion(self, regla, mensaje, donde, extracto="", seccion=None, parrafo=None):
        self.resultado.excepciones.append(Hallazgo(regla, mensaje, donde, extracto, seccion, parrafo))

    def fallo_en(self, regla, info, mensaje):
        self.fallo(regla, mensaje, info.donde, _extracto(texto_parrafo(info.p)), info.seccion, info.numero)

    def excepcion_en(self, regla, info, mensaje):
        self.excepcion(regla, mensaje, info.donde, _extracto(texto_parrafo(info.p)), info.seccion, info.numero)

    def comprobar(self):
        for comprobacion in (self.margenes, self.fuente, self.interlineado, self.alineacion,
                             self.sangria, self.paginacion):
            comprobacion()
        return self.resultado

    # --- 1.1 Márgenes -------------------------------------------------------------

    def margenes(self):
        for numero, seccion in enumerate(self.doc.sections, start=1):
            lados = (("superior", seccion.top_margin), ("inferior", seccion.bottom_margin),
                     ("izquierdo", seccion.left_margin), ("derecho", seccion.right_margin))
            mal = [f"{nombre} {cm(valor.twips)} cm" if valor is not None else f"{nombre} sin indicar"
                   for nombre, valor in lados
                   if valor is None or abs(valor.twips - margenes.MARGEN) > margenes.TOLERANCIA]
            if mal:
                self.fallo(MARGENES, f"Márgenes distintos de 2,54 cm: {', '.join(mal)}.",
                           f"Sección {numero}", seccion=numero)
            encuadernacion = seccion.gutter.twips if seccion.gutter is not None else 0
            if encuadernacion:
                self.fallo(MARGENES, f"Hay un margen de encuadernación de {cm(encuadernacion)} cm "
                                     "(debe ser 0).", f"Sección {numero}", seccion=numero)

    # --- 1.2 Fuente -------------------------------------------------------------------

    def _runs_con_texto(self, p):
        """(run, fuente, tamaño) de los fragmentos con letras que deben llevar la fuente del documento."""
        for r in fuente._runs_propios(p):
            if not fuente._letras(r) or en_ecuacion(r) or r.find(w("sym")) is not None:
                continue
            nombre = self.estilos.fuente(r, p)
            if fuente.es_de_simbolos(nombre):
                continue
            yield r, nombre, self.estilos.tamano(r, p)

    def _encabezados_y_pies(self):
        """(tipo, número de sección, elemento) de cada encabezado y pie propio, sin repetir los compartidos."""
        vistos = set()
        for numero, seccion in enumerate(self.doc.sections, start=1):
            for tipo, nombre in (("encabezado", "header"), ("encabezado de la primera página", "first_page_header"),
                                 ("encabezado de páginas pares", "even_page_header"), ("pie", "footer"),
                                 ("pie de la primera página", "first_page_footer"),
                                 ("pie de páginas pares", "even_page_footer")):
                cabecera = getattr(seccion, nombre)
                if cabecera.is_linked_to_previous or cabecera._element in vistos:
                    continue
                vistos.add(cabecera._element)
                yield tipo, numero, cabecera._element

    def fuente(self):
        codigo = [i for i in self.parrafos if fuente.es_codigo(i.p, self.estilos)]
        for info in codigo:
            nombre = next((n for _, n, _ in self._runs_con_texto(info.p)), "monoespaciada")
            self.excepcion_en(FUENTE, info, f"Código de programación: se deja en {nombre} (APA pide para "
                                            "el código una fuente monoespaciada).")

        # La fuente del documento es la que tiene más letras en el cuerpo.
        letras = {}
        de_codigo = {i.numero for i in codigo}
        cuerpo = [i for i in self.parrafos if i.numero not in de_codigo]
        for info in cuerpo:
            for r, nombre, tamano in self._runs_con_texto(info.p):
                letras[(nombre, tamano)] = letras.get((nombre, tamano), 0) + fuente._letras(r)
        if not letras:
            return
        principal = max(letras, key=letras.get)[0]
        familia = fuente.familia(principal)
        if familia is None:
            self.fallo(FUENTE, f"La fuente principal, {principal}, no es una de las recomendadas por APA "
                               f"({fuente._lista(fuente.RECOMENDADAS)}).", "Todo el documento")
            return
        puntos = fuente.RECOMENDADAS[familia]

        def revisar(p, con_tamano):
            """Lo que no está en la fuente o el tamaño del documento, p. ej. ["Calibri Light 16 pt"]."""
            distintos = []
            for _, nombre, tamano in self._runs_con_texto(p):
                if nombre != principal or (con_tamano and tamano != puntos):
                    descripcion = f"{nombre} {_pt(tamano)} pt" if con_tamano else nombre
                    if descripcion not in distintos:
                        distintos.append(descripcion)
            return distintos

        esperado = f"{principal} {_pt(puntos)} pt"
        for info in cuerpo:
            distintos = revisar(info.p, True)
            if distintos:
                self.fallo_en(FUENTE, info, f"Texto en {', '.join(distintos)} (debe ser {esperado}).")

        for tipo, numero, elemento in self._encabezados_y_pies():
            for p in elemento.iter(w("p")):
                distintos = revisar(p, True)
                if distintos:
                    self.fallo(FUENTE, f"Texto en {', '.join(distintos)} (debe ser {esperado}).",
                               f"Sección {numero} · {tipo}", _extracto(texto_parrafo(p)), numero)

        # Notas: misma fuente, pero el tamaño es el de Word por defecto.
        for tipo, parte, etiqueta in (("Nota al pie", "/word/footnotes.xml", "footnote"),
                                      ("Nota al final", "/word/endnotes.xml", "endnote")):
            for n, nota in enumerate(documento._notas(self.doc, parte, etiqueta), start=1):
                distintos = [d for p in nota.iter(w("p")) for d in revisar(p, False)]
                if distintos:
                    texto = " ".join(texto_parrafo(p) for p in nota.iter(w("p")))
                    self.fallo(FUENTE, f"Texto en {', '.join(dict.fromkeys(distintos))} (debe ser {principal}).",
                               f"{tipo} {n}", _extracto(texto))

    # --- 1.3 Interlineado -------------------------------------------------------------

    def _espacios_que_anulan(self, p):
        """Espaciado en líneas o automático (el de los textos pegados de la web), que manda sobre before/after."""
        anulan = []
        for atributo in ("beforeLines", "afterLines"):
            if _entero(self.estilos.parrafo(p, "spacing", atributo)):
                anulan.append("espacio medido en líneas")
        for atributo in ("beforeAutospacing", "afterAutospacing"):
            if self.estilos.parrafo(p, "spacing", atributo) in ("1", "true", "on"):
                anulan.append("espacio automático")
        return list(dict.fromkeys(anulan))

    def interlineado(self):
        for info in self.parrafos:
            if info.en_cuadro:
                continue
            linea, regla, antes, despues = interlineado._efectivo(self.estilos, info.p)
            problemas = []
            if info.en_tabla:
                if regla != "auto" or linea not in interlineado.ADMITIDOS_EN_TABLAS:
                    problemas.append(f"interlineado {interlineado._describir(linea, regla)} "
                                     "(en tablas: sencillo, 1,5 o doble)")
            elif (linea, regla) != (interlineado.DOBLE, "auto"):
                problemas.append(f"interlineado {interlineado._describir(linea, regla)} (debe ser doble)")
            if antes:
                problemas.append(f"{_pt(antes / interlineado.TWIPS_POR_PUNTO)} pt de espacio antes (debe ser 0)")
            if despues:
                problemas.append(f"{_pt(despues / interlineado.TWIPS_POR_PUNTO)} pt de espacio después "
                                 "(debe ser 0)")
            problemas += [f"{a} (debe ser 0)" for a in self._espacios_que_anulan(info.p)]
            if problemas:
                self.fallo_en(INTERLINEADO, info, "; ".join(problemas).capitalize() + ".")

    # --- 1.4 Alineación ----------------------------------------------------------------

    def _alineacion(self, p):
        return self.estilos.parrafo(p, "jc", "val") or "left"

    def alineacion(self):
        portada = set(alineacion.parece_portada([i.p for i in self.principales], self.estilos))
        for info in self.parrafos:
            jc = self._alineacion(info.p)
            if jc in alineacion.JUSTIFICADOS:
                self.fallo_en(ALINEACION, info, "Texto justificado (debe ir a la izquierda).")
            elif jc in alineacion.A_LA_DERECHA and not info.en_tabla and not info.en_cuadro:
                if info.p in portada:
                    self.excepcion_en(ALINEACION, info, "Alineado a la derecha en lo que parece la portada "
                                                        "de la universidad (se trata en el paso 5.3).")
                else:
                    self.fallo_en(ALINEACION, info, "Alineado a la derecha (debe ir a la izquierda).")

        for tipo, parte, etiqueta in (("Nota al pie", "/word/footnotes.xml", "footnote"),
                                      ("Nota al final", "/word/endnotes.xml", "endnote")):
            for n, nota in enumerate(documento._notas(self.doc, parte, etiqueta), start=1):
                if any(self._alineacion(p) in alineacion.JUSTIFICADOS for p in nota.iter(w("p"))):
                    texto = " ".join(texto_parrafo(p) for p in nota.iter(w("p")))
                    self.fallo(ALINEACION, "Texto justificado (debe ir a la izquierda).", f"{tipo} {n}",
                               _extracto(texto))

        if _activado(self.doc.settings.element.find(w("autoHyphenation"))):
            self.fallo(ALINEACION, "Los guiones automáticos a final de línea están activados.",
                       "Ajustes del documento")
        opcionales = sum(1 for _ in self.doc.element.body.iter(w("softHyphen")))
        if opcionales:
            self.excepcion(ALINEACION, f"Hay {opcionales} guion(es) opcional(es): quitarlos modificaría el "
                                       "texto, así que solo se avisa.", "Todo el documento")

    # --- 1.5 Sangría -----------------------------------------------------------------------

    def _primera_linea(self, p):
        """Sangría de primera línea efectiva en twips, o None si se mide en caracteres."""
        if _entero(self.estilos.parrafo(p, "ind", "firstLineChars")):
            return None
        return sangria._sangria(self.estilos, p, "firstLine")

    def sangria(self):
        for p, motivo in sangria.motivos_sin_sangria(self.doc, self.estilos):
            info = self.por_elemento[p]
            primera = self._primera_linea(p)
            francesa = sangria._sangria(self.estilos, p, "hanging")
            if motivo is None:
                if francesa or primera != sangria.SANGRIA:
                    actual = (f"francesa de {cm(francesa)} cm" if francesa else
                              "medida en caracteres" if primera is None else
                              f"de {cm(primera)} cm" if primera else "ninguna")
                    self.fallo_en(SANGRIA, info, f"Párrafo de texto sin la sangría de 1,27 cm en la primera "
                                                 f"línea (tiene {actual}).")
            elif motivo == "centrado o a la derecha" and self._alineacion(p) in alineacion.A_LA_DERECHA:
                continue   # el problema es la alineación (1.4); con ella corregida, se vería su sangría
            elif motivo == "tabulador":
                self.excepcion_en(SANGRIA, info, "Empieza con un tabulador o espacios usados como sangría: "
                                                 "no se le pone sangría (se avisa en el informe).")
            elif motivo != "vacío" and not francesa and primera != 0:
                medida = "en caracteres" if primera is None else f"de {cm(primera)} cm"
                self.fallo_en(SANGRIA, info, f"Lleva sangría de primera línea {medida}, pero no es un "
                                             f"párrafo de texto ({motivo}): no debe llevarla.")

    # --- 1.6 Número de página -------------------------------------------------------------

    def _revisar_encabezado(self, elemento):
        """Problemas de un encabezado: debe tener solo el número de página, a la derecha."""
        problemas = []
        con_numero = [p for p in elemento.iter(w("p")) if paginacion.tiene_numero_de_pagina(p)]
        if not con_numero:
            problemas.append("no tiene el número de página")
        elif not any(self._alineacion(p) in alineacion.A_LA_DERECHA for p in con_numero):
            problemas.append("el número de página no está a la derecha")
        textos = [paginacion._texto_sin_campos(p) for p in elemento.iter(w("p"))
                  if not _dentro_de(p, w("tbl"))]
        textos = [t for t in textos if t]
        if textos:
            problemas.append(f"tiene otro texto («{'», «'.join(_extracto(t) for t in textos)}»)")
        imagenes = sum(1 for _ in elemento.iter(w("drawing"), w("pict")))
        tablas = sum(1 for _ in elemento.iter(w("tbl")))
        if imagenes:
            problemas.append(f"tiene {imagenes} imagen(es)")
        if tablas:
            problemas.append(f"tiene {tablas} tabla(s)")
        return problemas

    def paginacion(self):
        if self.doc.settings.odd_and_even_pages_header_footer:
            self.fallo(PAGINACION, "Está activado «Pares e impares diferentes»: las páginas pares usan otro "
                                   "encabezado.", "Ajustes del documento")

        secciones = list(self.doc.sections)
        revisados = {}
        for numero, seccion in enumerate(secciones, start=1):
            donde = f"Sección {numero}"
            if seccion.different_first_page_header_footer:
                self.fallo(PAGINACION, "Está activado «Primera página diferente»: la primera página usa "
                                       "otro encabezado.", donde, seccion=numero)
            tipo = seccion._sectPr.find(w("pgNumType"))
            if tipo is not None and tipo.get(w("start")) is not None:
                self.fallo(PAGINACION, f"La numeración se reinicia en {tipo.get(w('start'))} (debe ser "
                                       "continua).", donde, seccion=numero)
            if tipo is not None and tipo.get(w("fmt")) not in (None, "decimal"):
                self.fallo(PAGINACION, f"La numeración usa el formato «{tipo.get(w('fmt'))}» (debe ser "
                                       "con números normales).", donde, seccion=numero)

            # El encabezado que se ve: el propio o, si está enlazado, el de la sección anterior.
            propio = next((s.header for s in reversed(secciones[:numero]) if not s.header.is_linked_to_previous),
                          None)
            if propio is None:
                self.fallo(PAGINACION, "No tiene encabezado, así que no se ve el número de página.",
                           donde, seccion=numero)
                continue
            if propio._element in revisados:   # compartido con otra sección: ya se ha dicho
                continue
            revisados[propio._element] = numero
            problemas = self._revisar_encabezado(propio._element)
            if problemas:
                self.fallo(PAGINACION, "El encabezado " + "; ".join(problemas) + ".", f"{donde} · encabezado",
                           seccion=numero)

        for tipo, numero, elemento in self._encabezados_y_pies():
            if not tipo.startswith("pie"):
                continue
            for p in elemento.iter(w("p")):
                if paginacion.tiene_numero_de_pagina(p):
                    self.fallo(PAGINACION, "El pie tiene un número de página (debe ir solo arriba a la derecha).",
                               f"Sección {numero} · {tipo}", _extracto(documento.texto_legible(p)), numero)
                elif paginacion._texto_sin_campos(p):
                    self.excepcion(PAGINACION, "El pie tiene texto que no es un número de página: se deja "
                                               "(se avisa en el informe).", f"Sección {numero} · {tipo}",
                                   _extracto(paginacion._texto_sin_campos(p)), numero)


def verificar(origen):
    """
    Comprueba las reglas de la fase 1 en un documento: una ruta a un .docx o un
    documento ya abierto con python-docx. Devuelve un Resultado.
    """
    if isinstance(origen, (str, Path)):
        doc, archivo = documento.abrir(origen), Path(origen).name
    else:
        doc, archivo = origen, "(documento en memoria)"
    return Verificacion(doc, archivo).comprobar()


# --- Presentación ---------------------------------------------------------------------

def _linea(h):
    extracto = f" «{h.extracto}»" if h.extracto else ""
    return f"   · {h.donde}{extracto}: {h.mensaje}"


def mostrar(resultado):
    """Convierte el resultado en un texto legible para la terminal."""
    L = ["=" * 72, f"📄 {resultado.archivo}", "=" * 72]
    for regla in REGLAS:
        hallazgos = [h for h in resultado.incumplimientos if h.regla == regla]
        if not hallazgos:
            L.append(f"✔ {regla}: cumple")
            continue
        L.append(f"✘ {regla}: {len(hallazgos)} incumplimiento{'s' if len(hallazgos) > 1 else ''}")
        L += [_linea(h) for h in hallazgos[:MAXIMO_POR_REGLA]]
        if len(hallazgos) > MAXIMO_POR_REGLA:
            L.append(f"   … y {len(hallazgos) - MAXIMO_POR_REGLA} más.")

    if resultado.excepciones:
        L.append("\nExcepciones admitidas (no son incumplimientos)")
        L += [_linea(h).replace("   · ", f"   · {h.regla} · ", 1) for h in resultado.excepciones]

    total = len(resultado.incumplimientos)
    L.append("\nResultado: " + ("cumple todas las reglas de la fase 1." if resultado.cumple else
                                f"{total} incumplimiento{'s' if total > 1 else ''}."))
    return "\n".join(L)


def main(argv=None):
    configurar_terminal()
    archivos = sys.argv[1:] if argv is None else argv
    if not archivos:
        print("Indica qué documento verificar. Ejemplo: "
              "python herramientas/verificar.py salida/01_ensayo_basico_APA.docx")
        return 2
    alguno_mal = False
    for archivo in archivos:
        try:
            resultado = verificar(archivo)
            print(mostrar(resultado))
            alguno_mal |= not resultado.cumple
        except ErrorDocumento as e:
            print(f"✘ {Path(archivo).name}: {e}")
            alguno_mal = True
        print()
    return 1 if alguno_mal else 0


if __name__ == "__main__":
    sys.exit(main())
