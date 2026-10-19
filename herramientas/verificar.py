"""
Verificador: comprueba si un documento Word cumple las reglas APA de la
fase 1 y los estilos de título (paso 4.1), y lista cada incumplimiento con su
ubicación (sección, número de párrafo y un extracto del texto). No modifica nada.

Cómo comprueba:
- Lee el XML del documento y calcula el formato efectivo de cada párrafo y de
  cada fragmento de texto (valores por defecto → estilos → formato directo).
  No pregunta a Word, que da "fuente mezclada" en rangos grandes aunque no lo
  estén (ver "Problemas abiertos" en ROADMAP.md).
- Respeta las mismas excepciones que las reglas (portada, código, párrafos con
  tabulador…), usando los mismos detectores. Lo que entra por excepción no es
  un incumplimiento: se lista aparte, como "excepción admitida".
- Qué párrafos deben llevar sangría y cuáles son títulos lo decide el
  clasificador de párrafos (apa_facil/estructura.py), igual que en las reglas.
  El verificador solo tiene la copia, así que la clasifica a ella; el
  programa, en cambio, clasificó el original. Desde el paso 4.1 cada título de
  la copia lleva su estilo de Word, así que su estructura queda escrita en ella
  y las dos clasificaciones coinciden (lo comprueba test_cinco_documentos.py).
  Un título que se quedara sin estilo aparece como incumplimiento.
- Limitación: si el clasificador se equivoca (por ejemplo, toma un párrafo de
  texto por un título), el verificador no lo ve. Para eso están las pruebas del
  clasificador con la clasificación esperada de los documentos de prueba.

Uso:
    python herramientas/verificar.py archivo.docx [otro.docx ...]

Termina con código 0 si todos los documentos cumplen y 1 si alguno no.
"""

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

# Para encontrar el paquete apa_facil al ejecutar este archivo directamente.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from apa_facil import documento, estructura  # noqa: E402
from apa_facil.__main__ import configurar_terminal  # noqa: E402
from apa_facil.documento import ErrorDocumento, Estilos, cm, en_ecuacion, texto_parrafo, w  # noqa: E402
from apa_facil.reglas import (alineacion, fuente, interlineado, margenes, paginacion, sangria,  # noqa: E402
                               titulos)
from apa_facil.reglas import portada as regla_portada  # noqa: E402
from apa_facil.reglas import resumen as resumen_regla  # noqa: E402
from apa_facil.reglas import apendices as regla_apendices  # noqa: E402
from apa_facil.reglas import citas as regla_citas  # noqa: E402
from apa_facil.reglas import leyendas  # noqa: E402
from apa_facil.reglas import orden as regla_orden  # noqa: E402
from apa_facil.reglas import referencias as regla_referencias  # noqa: E402
from apa_facil.reglas import tablas as regla_tablas  # noqa: E402

MARGENES = "1.1 Márgenes"
FUENTE = "1.2 Fuente"
INTERLINEADO = "1.3 Interlineado"
ALINEACION = "1.4 Alineación"
SANGRIA = "1.5 Sangría"
PAGINACION = "1.6 Número de página"
TITULOS = "4.1 Estilos de título"
FORMATO_TITULOS = "4.2 Formato de los títulos"
TEXTO_TITULOS = "4.3 Numeración y mayúsculas de los títulos"
TITULO_TRABAJO = "4.4 Título del trabajo"
TITULOS_EN_LINEA = "4.5 Títulos de nivel 4 y 5 en línea"
MAYUSCULAS_TITULOS = "4.6 Mayúsculas de los títulos"
PORTADA = "5.2 Portada"
PORTADA_PROPIA = "5.3 Portada propia"
INDICE = "5.4 Índice"
RESUMEN = "5.5 Resumen"
TABLAS = "6.1 Tablas"
FIGURAS = "6.2 Figuras"
CITAS = "6.3 Citas en bloque"
REFERENCIAS = "6.4 Referencias"
ORDEN = "6.5 Orden alfabético"
APENDICES = "6.7 Apéndices"
REGLAS = [MARGENES, FUENTE, INTERLINEADO, ALINEACION, SANGRIA, PAGINACION, TITULOS, FORMATO_TITULOS, TEXTO_TITULOS,
          TITULO_TRABAJO, TITULOS_EN_LINEA, MAYUSCULAS_TITULOS, PORTADA, PORTADA_PROPIA, INDICE, RESUMEN, TABLAS,
          FIGURAS, CITAS, REFERENCIAS, ORDEN, APENDICES]
SIN_LINEA = (None, "nil", "none")
SIN_FONDO = (None, "auto", "FFFFFF", "ffffff")

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
        self.clasificados = estructura.clasificar(doc, self.estilos)
        self.clasificado_de = {c.elemento: c for c in self.clasificados}

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
                             self.sangria, self.paginacion, self.titulos, self.formato_de_titulos,
                             self.texto_de_titulos, self.titulo_del_trabajo, self.titulos_en_linea,
                             self.mayusculas_de_titulos, self.portada, self.indice, self.resumen, self.tablas,
                             self.figuras, self.citas, self.referencias,
                             self.orden, self.apendices):
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
                                                        "de la universidad, que se mantiene con su diseño (paso 5.3).")
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
        for p, motivo in sangria.motivos_sin_sangria(self.doc, self.estilos, self.clasificados):
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
            elif self._titulo_con_sangria(p):
                continue   # los títulos de nivel 4 y 5 la llevan: la pone su estilo (se comprueba en 4.1)
            elif self._sigue_a_una_cita(p):
                continue   # el segundo párrafo de una cita en bloque y los siguientes la llevan (se comprueba en 6.3)
            elif motivo == "tabulador":
                self.excepcion_en(SANGRIA, info, "Empieza con un tabulador o espacios usados como sangría: "
                                                 "no se le pone sangría (se avisa en el informe).")
            elif motivo != "vacío" and not francesa and primera != 0:
                medida = "en caracteres" if primera is None else f"de {cm(primera)} cm"
                self.fallo_en(SANGRIA, info, f"Lleva sangría de primera línea {medida}, pero no es un "
                                             f"párrafo de texto ({motivo}): no debe llevarla.")

    def _titulo_con_sangria(self, p):
        c = self.clasificado_de.get(p)
        return c is not None and c.tipo == "titulo" and c.nivel is not None and titulos.FORMATO[c.nivel][2] > 0

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


    # --- 4.1 Estilos de título -----------------------------------------------------------

    def _del_estilo(self, estilo, capa, etiqueta, atributo=None):
        """El primer `etiqueta` de `capa` (pPr o rPr) en el estilo, en los que hereda o en los valores por defecto."""
        base = estilo.find(w("basedOn"))
        cadena = [estilo, *self.estilos.cadena(base.get(w("val")) if base is not None else None)]
        capas = [e.find(w(capa)) for e in cadena]
        capas.append(self.estilos.ppr_defecto if capa == "pPr" else self.estilos.rpr_defecto)
        for c in capas:
            el = c.find(w(etiqueta)) if c is not None else None
            if el is not None and (atributo is None or el.get(w(atributo)) is not None):
                return el
        return None

    def _problemas_de_estilo(self, estilo, nivel, esquema=True):
        """
        Lo que le falta a un estilo «Título N» para tener el formato APA de su nivel.
        Con esquema=False (el estilo «Title» del título del trabajo), no debe tener
        nivel de esquema: lo lleva el párrafo del título, no la portada que use el estilo.
        """
        alineacion_apa, cursiva, primera_linea = titulos.FORMATO[nivel]
        problemas = []
        jc = self._del_estilo(estilo, "pPr", "jc", "val")
        jc = jc.get(w("val")) if jc is not None else "left"
        if {"start": "left"}.get(jc, jc) != alineacion_apa:
            problemas.append("no está centrado" if alineacion_apa == "center" else "no está a la izquierda")
        if not _activado(self._del_estilo(estilo, "rPr", "b")):
            problemas.append("no va en negrita")
        if _activado(self._del_estilo(estilo, "rPr", "i")) != cursiva:
            problemas.append("no va en cursiva" if cursiva else "va en cursiva")
        medidas = {}
        for atributo in ("firstLine", "hanging", "left"):
            el = self._del_estilo(estilo, "pPr", "ind", atributo)
            medidas[atributo] = _entero(el.get(w(atributo))) if el is not None else 0
        if medidas["firstLine"] != primera_linea or medidas["hanging"] or medidas["left"]:
            esperada = "1,27 cm en la primera línea" if primera_linea else "ninguna"
            problemas.append(f"la sangría no es la de su nivel ({esperada})")
        if not _activado(self._del_estilo(estilo, "pPr", "keepNext")):
            problemas.append("no tiene «Conservar con el siguiente»")
        nivel_de_esquema = self._del_estilo(estilo, "pPr", "outlineLvl", "val")
        if esquema and (nivel_de_esquema is None or nivel_de_esquema.get(w("val")) != str(nivel - 1)):
            problemas.append(f"no tiene el nivel de esquema {nivel} (el que usa el panel de navegación)")
        if not esquema and nivel_de_esquema is not None and nivel_de_esquema.get(w("val")) != "9":
            problemas.append("tiene nivel de esquema (la portada que lo use parecería un título)")
        color = self._del_estilo(estilo, "rPr", "color", "val")
        if color is not None and color.get(w("val")) != "auto":
            problemas.append("va en color")
        subrayado = self._del_estilo(estilo, "rPr", "u")
        if subrayado is not None and subrayado.get(w("val")) != "none":
            problemas.append("va subrayado")
        if any(_activado(self._del_estilo(estilo, "rPr", e)) for e in ("caps", "smallCaps")):
            problemas.append("pone el texto en mayúsculas")
        if self._del_estilo(estilo, "pPr", "pBdr") is not None:
            problemas.append("tiene bordes")
        return problemas

    def titulos(self):
        por_nivel = titulos.estilos_de_titulo(self.doc)
        for nivel, estilos in por_nivel.items():
            ids = [e.get(w("styleId")) for e in estilos]
            if len(ids) != len(set(ids)):
                self.fallo(TITULOS, f"El estilo «Título {nivel}» está definido varias veces: Word puede usar "
                                    "cualquiera de las definiciones.", "Estilos de Word")
            for estilo in estilos:
                problemas = self._problemas_de_estilo(estilo, nivel)
                if problemas:
                    self.fallo(TITULOS, f"El estilo «Título {nivel}» no tiene el formato APA: "
                                        f"{'; '.join(problemas)}.", "Estilos de Word")
        for estilo in titulos.estilos_del_titulo_del_trabajo(self.doc):
            problemas = self._problemas_de_estilo(estilo, 1, esquema=False)
            if problemas:
                self.fallo(TITULOS, f"El estilo «Título» (el del título del trabajo) no tiene el formato APA del "
                                    f"nivel 1: {'; '.join(problemas)}.", "Estilos de Word")

        for c in self.clasificados:
            if c.tipo != "titulo" or not c.nivel or self._es_titulo_del_trabajo(c):
                continue
            if self.estilos.estilo_de(c.elemento) in {e.get(w("styleId")) for e in por_nivel[c.nivel]}:
                continue
            info = self.por_elemento[c.elemento]
            if c.confianza == estructura.BAJA:
                self.excepcion_en(TITULOS, info, "Parece un título, pero no es seguro: no se le pone estilo "
                                                 "(se avisa en el informe).")
            else:
                self.fallo_en(TITULOS, info, f"Título de nivel {c.nivel} sin el estilo «Título {c.nivel}» de Word.")


    # --- 4.2 Formato de cada título ------------------------------------------------------

    def _de_texto(self, r, p, etiqueta):
        """¿Está activado en el fragmento un ajuste de texto como la negrita ("b")? (formato efectivo)"""
        for capa in self.estilos._capas_texto(r, p):
            el = capa.find(w(etiqueta)) if capa is not None else None
            if el is not None:
                return el
        return None

    def formato_de_titulos(self):
        """
        Cada título, con su formato final (estilo y formato puesto a mano): el
        de su nivel. Los que no tienen su estilo ya salen en 4.1. Se admite la
        cursiva en una parte del título (el título de un libro, una especie).
        """
        con_estilo = {e.get(w("styleId")): nivel for nivel, estilos in titulos.estilos_de_titulo(self.doc).items()
                      for e in estilos}
        con_estilo.update({e.get(w("styleId")): 1 for e in titulos.estilos_del_titulo_del_trabajo(self.doc)})
        for c in self.clasificados:
            if c.tipo != "titulo" or con_estilo.get(self.estilos.estilo_de(c.elemento)) != c.nivel:
                continue
            p, info = c.elemento, self.por_elemento[c.elemento]
            alineacion_apa, cursiva, primera_linea = titulos.FORMATO[c.nivel]
            problemas = []
            if {"start": "left"}.get(self._alineacion(p), self._alineacion(p)) != alineacion_apa:
                problemas.append("no está centrado" if alineacion_apa == "center" else "no está a la izquierda")
            # La sangría de primera línea de los niveles 1 a 3 ya la comprueba 1.5 (no deben llevarla).
            sangrias = {a: _entero(self.estilos.parrafo(p, "ind", a)) for a in ("firstLine", "hanging", "left")}
            if sangrias["hanging"] or sangrias["left"] or \
                    (primera_linea and (sangrias["firstLine"] != primera_linea or
                                        _entero(self.estilos.parrafo(p, "ind", "firstLineChars")))):
                problemas.append("la sangría no es la de su nivel")
            if not _activado(self._capa_de_parrafo(p, "keepNext")):
                problemas.append("no tiene «Conservar con el siguiente»")
            fragmentos = titulos.fragmentos_con_letras(p)
            if any(not _activado(self._de_texto(r, p, "b")) for r in fragmentos):
                problemas.append("no va todo en negrita")
            if fragmentos and all(titulos.cursiva_efectiva(self.estilos, r, p) != cursiva for r in fragmentos):
                problemas.append("no va en cursiva" if cursiva else "va en cursiva")
            if any(_activado(self._de_texto(r, p, e)) for r in fragmentos for e in ("caps", "smallCaps")):
                problemas.append("tiene mayúsculas forzadas")
            enlaces = [r for r in fragmentos if next(r.iterancestors(w("hyperlink")), None) is not None]
            if any(r not in enlaces and (self._de_texto(r, p, "color") is not None and
                                         self._de_texto(r, p, "color").get(w("val")) not in (None, "auto"))
                   for r in fragmentos):
                problemas.append("va en color")
            if any(r not in enlaces and (self._de_texto(r, p, "u") is not None and
                                         self._de_texto(r, p, "u").get(w("val")) != "none")
                   for r in fragmentos):
                problemas.append("va subrayado")
            if problemas:
                self.fallo_en(FORMATO_TITULOS, info, f"Título de nivel {c.nivel}: {'; '.join(problemas)}.")

    # --- 4.3 Numeración y mayúsculas -------------------------------------------------------

    def texto_de_titulos(self):
        """
        APA no numera los títulos, pero quitar la numeración modifica el texto:
        por defecto se deja y se avisa (solo se quita con --quitar-numeracion).
        Por eso no es un incumplimiento, sino una excepción admitida. Las
        mayúsculas las comprueba 4.6.
        """
        for c in self.clasificados:
            if c.tipo != "titulo" or c.confianza == estructura.BAJA or self._es_titulo_del_trabajo(c):
                continue
            texto, info = texto_parrafo(c.elemento), self.por_elemento[c.elemento]
            if titulos.numeracion(texto):
                self.excepcion_en(TEXTO_TITULOS, info, "Título numerado: APA no numera los títulos, pero se deja "
                                                       "(se avisa en el informe; --quitar-numeracion lo quita).")

    # --- 4.4 Título del trabajo ---------------------------------------------------------------

    def _es_titulo_del_trabajo(self, c):
        """¿Es el título del trabajo repetido al principio del texto? (estilo «Title» fuera de la portada)"""
        estilo = self.estilos.por_id.get(self.estilos.estilo_de(c.elemento))
        return c.zona != "portada" and estilo is not None and \
            titulos._nombre(estilo) == titulos.NOMBRE_DEL_ESTILO_DEL_TITULO

    def titulo_del_trabajo(self):
        """
        El texto empieza con el título del trabajo: el párrafo con el estilo «Title» y
        el mismo texto que la portada, salvo mayúsculas y minúsculas (su formato lo
        comprueban 4.1 y 4.2, y sus mayúsculas, 4.6). Si no hay portada, no se puede
        saber cuál es. El rótulo «Introducción» se deja a propósito: es una excepción
        admitida.
        """
        portada = [texto_parrafo(c.elemento).strip().lower() for c in self.clasificados
                   if c.zona == "portada" and c.texto]
        primero = next((c for c in self.clasificados if c.zona == "cuerpo" and c.tipo != "vacio"), None)
        if not portada or primero is None:
            return
        info = self.por_elemento[primero.elemento]
        lineas = texto_parrafo(primero.elemento).lower().split("\n")
        en_la_portada = any(portada[i:i + len(lineas)] == lineas for i in range(len(portada)))
        if not self._es_titulo_del_trabajo(primero) or not en_la_portada:
            self.fallo_en(TITULO_TRABAJO, info, "El texto no empieza con el título del trabajo (el de la portada, "
                                                "con el estilo «Título» de Word).")
        rotulo = next((c for c in self.clasificados if c.zona == "cuerpo" and c.tipo == "titulo"
                       and not self._es_titulo_del_trabajo(c)), None)
        if rotulo is not None and rotulo.nivel == 1 and documento.es_rotulo_de_introduccion(texto_parrafo(rotulo.elemento)):
            self.excepcion_en(TITULO_TRABAJO, self.por_elemento[rotulo.elemento],
                              "Rótulo «Introducción»: APA no lo usa, pero se deja (se avisa en el informe; "
                              "--quitar-introduccion lo quita si no tiene subapartados).")

    # --- 4.5 Títulos de nivel 4 y 5 en línea -----------------------------------------------

    def titulos_en_linea(self):
        """
        Los títulos de nivel 4 y 5 seguidos de un párrafo de texto van en su misma línea
        (marca de párrafo oculta: el separador de estilo de Word) y terminan en punto y
        un espacio. Los que no van seguidos de texto no se pueden poner en línea, y los que
        terminan en dos puntos se dejan: son excepciones admitidas (se avisa en el informe).
        """
        por_elemento = {c.elemento: c for c in self.clasificados}
        for c in self.clasificados:
            if c.tipo != "titulo" or not c.nivel or c.nivel < 4 or c.confianza == estructura.BAJA:
                continue
            info, texto = self.por_elemento[c.elemento], texto_parrafo(c.elemento)
            detras = por_elemento.get(titulos.siguiente_bloque(c.elemento))
            if detras is None or detras.tipo != "cuerpo":
                if titulos.en_linea(c.elemento):
                    self.fallo_en(TITULOS_EN_LINEA, info, "Tiene la marca de párrafo oculta, pero no le sigue un "
                                                          "párrafo de texto: se uniría con lo que venga detrás.")
                else:
                    self.excepcion_en(TITULOS_EN_LINEA, info, "No va seguido de un párrafo de texto: no se puede "
                                                              "poner en la misma línea (se avisa en el informe).")
                continue
            problemas = []
            if not titulos.en_linea(c.elemento):
                problemas.append("no va en la misma línea que su texto")
            if not texto.endswith(" "):
                problemas.append("no termina con un espacio (el texto quedaría pegado)")
            if titulos.ACABA_CON_DOS_PUNTOS.search(texto.rstrip()):
                self.excepcion_en(TITULOS_EN_LINEA, info, "Termina en dos puntos o punto y coma: se deja (se avisa "
                                                          "en el informe).")
            elif not titulos.ACABA_CON_SIGNO.search(texto.rstrip()):
                problemas.append("no termina en punto")
            if problemas:
                self.fallo_en(TITULOS_EN_LINEA, info, f"Título de nivel {c.nivel}: {'; '.join(problemas)}.")

    # --- 4.6 Mayúsculas de los títulos --------------------------------------------------------

    def mayusculas_de_titulos(self):
        """
        Los títulos (niveles 1 a 5 y el del trabajo al principio del texto) llevan
        mayúscula inicial en las palabras principales, como pide la web oficial de APA
        (Title case capitalization): son tal como los deja titulos.mayusculas_apa. Las
        palabras dudosas («que», «lo», «vía») se admiten de las dos formas.
        """
        for c in self.clasificados:
            es_del_trabajo = self._es_titulo_del_trabajo(c) and c.zona == "cuerpo"
            if not es_del_trabajo and (c.tipo != "titulo" or c.confianza == estructura.BAJA):
                continue
            mal = []
            for linea in texto_parrafo(c.elemento).split("\n"):
                esperada = titulos.mayusculas_apa(linea, self.clasificados)[0]
                if esperada != linea:
                    mal.append(esperada.strip())
            if mal:
                self.fallo_en(MAYUSCULAS_TITULOS, self.por_elemento[c.elemento],
                              f"No lleva las mayúsculas de APA (sería «{' / '.join(mal)}»).")

    # --- 5.2 Portada -------------------------------------------------------------------------

    def portada(self):
        """
        La portada APA de estudiante (la que genera el paso 5.2: empieza con líneas en
        blanco y sigue con un párrafo con el estilo «Title»), como la describe la web
        oficial de APA: el título 3 o 4 líneas por debajo del margen superior, centrado
        y en negrita; una sola línea en blanco; debajo, cada dato centrado en su línea,
        sin líneas en blanco entre ellos; y el texto en una página nueva. El doble
        espacio y la sangría ya los comprueban 1.3 y 1.5.

        Paso 5.3: otra portada (la de la universidad, que se mantiene con
        --mantener-portada o sin datos) es una excepción admitida, y también tiene que
        dejar el texto en una página nueva.
        """
        portada = [c for c in self.clasificados if c.zona == "portada"]
        if not any(c.texto for c in portada):
            return
        blancos = next((n for n, c in enumerate(portada) if c.tipo != "vacio"), len(portada))
        if not blancos or blancos == len(portada) or not self._con_estilo_title(portada[blancos].elemento):
            primera = next(c for c in portada if c.texto)
            self.excepcion_en(PORTADA_PROPIA, self.por_elemento[primera.elemento],
                              "Portada propia, no la de estudiante de APA: se ha mantenido (se avisa en el informe).")
            self._texto_en_pagina_nueva(PORTADA_PROPIA, portada)
            return
        titulo, debajo = portada[blancos], portada[blancos + 1:]
        info = self.por_elemento[titulo.elemento]
        if blancos not in (3, 4):
            self.fallo_en(PORTADA, info, f"El título va {blancos} línea(s) por debajo del margen superior "
                                         "(APA: 3 o 4).")
        problemas = []
        if self._alineacion(titulo.elemento) != "center":
            problemas.append("no está centrado")
        runs = [r for r in fuente._runs_propios(titulo.elemento) if fuente._letras(r)]
        if not runs or not all(_activado(self._de_texto(r, titulo.elemento, "b")) for r in runs):
            problemas.append("no está en negrita")
        if problemas:
            self.fallo_en(PORTADA, info, f"Título de la portada: {' y '.join(problemas)}.")
        if not debajo or debajo[0].tipo != "vacio" or (len(debajo) > 1 and debajo[1].tipo == "vacio"):
            self.fallo_en(PORTADA, info, "Entre el título y los autores tiene que haber una sola línea en blanco.")
        for c in debajo[1:]:
            datos = self.por_elemento[c.elemento]
            if c.tipo == "vacio":
                self.fallo_en(PORTADA, datos, "Línea en blanco entre los datos de la portada (no debe haberla).")
            elif self._alineacion(c.elemento) != "center":
                self.fallo_en(PORTADA, datos, "Dato de la portada sin centrar.")
        self._texto_en_pagina_nueva(PORTADA, portada)

    def _texto_en_pagina_nueva(self, regla, portada):
        """Después de la portada, el texto empieza en una página nueva (sin contar las marcas que deja Word)."""
        ultimo = portada[-1].elemento
        siguiente = next((i for i in self.principales if i.numero > self.por_elemento[ultimo].numero), None)
        empieza_pagina = regla_portada._termina_en_salto(ultimo) or (
            siguiente is not None and alineacion._salto_antes(self.estilos, siguiente.p))
        if not empieza_pagina:
            self.fallo_en(regla, self.por_elemento[ultimo], "Después de la portada, el texto no empieza en una "
                                                            "página nueva.")

    # --- 5.4 Índice -------------------------------------------------------------------------

    def indice(self):
        """
        APA no usa el índice en los trabajos de estudiante (Order of pages), pero muchas
        universidades lo exigen: por defecto se deja (--quitar-indice lo quita). Por eso
        no es un incumplimiento, sino una excepción admitida.
        """
        primero = next((c for c in self.clasificados if c.zona == "indice"), None)
        if primero is not None:
            self.excepcion_en(INDICE, self.por_elemento[primero.elemento],
                              "Índice: APA no lo usa en los trabajos de estudiante, pero se deja (se avisa en el "
                              "informe; --quitar-indice lo quita).")

    # --- 5.5 Resumen ------------------------------------------------------------------------

    def resumen(self):
        """
        Las palabras clave, como las describe la guía oficial de APA (Abstract and
        Keywords Guide): el rótulo «Palabras clave:» en cursiva y sin negrita, las
        palabras sin cursiva ni negrita y sin punto al final (el de una abreviatura se
        admite: es una excepción). Y lo que sigue al resumen empieza en una página
        nueva. El rótulo «Resumen» y la sangría ya los comprueban 4.1, 4.2 y 1.5.
        """
        resumen = [c for c in self.clasificados if c.zona == "resumen"]
        if not resumen:
            return
        clave = next((c for c in resumen if documento.PALABRAS_CLAVE.match(c.texto)), None)
        if clave is not None:
            p, info = clave.elemento, self.por_elemento[clave.elemento]
            texto = texto_parrafo(p)
            rotulo = documento.PALABRAS_CLAVE.match(texto)
            final = rotulo.end(2) if rotulo.group(2) else rotulo.end(1)
            problemas, desde = [], 0
            for r in fuente._runs_propios(p):
                largo = sum(len(t.text or "") for t in r.iter(w("t")))
                if not "".join(t.text or "" for t in r.iter(w("t"))).strip():
                    desde += largo
                    continue
                cursiva, negrita = (_activado(self._de_texto(r, p, e)) for e in ("i", "b"))
                if desde < final and (not cursiva or negrita):
                    problemas.append("el rótulo tiene que ir en cursiva y sin negrita")
                elif desde >= final and (cursiva or negrita):
                    problemas.append("las palabras clave no van en cursiva ni en negrita")
                desde += largo
            if texto.endswith("."):
                if resumen_regla.ABREVIATURA_FINAL.search(texto[rotulo.end():]):
                    self.excepcion_en(RESUMEN, info, "Las palabras clave acaban en el punto de una abreviatura: se deja "
                                                     "(se avisa en el informe).")
                else:
                    problemas.append("terminan en punto")
            for problema in dict.fromkeys(problemas):
                self.fallo_en(RESUMEN, info, f"Palabras clave: {problema}.")
        ultimo = resumen[-1].elemento
        siguiente = next((i for i in self.principales if i.numero > self.por_elemento[ultimo].numero), None)
        if siguiente is not None and not regla_portada._termina_en_salto(ultimo) and \
                not alineacion._salto_antes(self.estilos, siguiente.p):
            self.fallo_en(RESUMEN, self.por_elemento[ultimo], "Lo que sigue al resumen no empieza en una página nueva.")

    # --- 6.1 Tablas ---------------------------------------------------------------------------

    def tablas(self):
        """
        Cada tabla del texto, como la describe la web oficial de APA (Table setup y
        Sample tables): sus leyendas (ver _leyendas_de); la fila de encabezados se
        repite en cada página; sin bordes verticales ni colores de fondo en las
        celdas. No se miran las tablas de la portada o del índice, los recuadros de
        una sola celda ni las tablas que colocan imágenes (como la regla).
        """
        bloques = estructura._bloques(self.doc)
        for k, tbl in enumerate(bloques):
            if not regla_tablas._es_tabla_de_datos(tbl) or self._en_portada_o_indice(bloques, k):
                continue
            info = self.por_elemento[tbl.find(f".//{w('p')}")]
            self._leyendas_de(TABLAS, leyendas.TABLA, bloques, k, info)
            problemas = []
            primera = tbl.find(w("tr"))
            if not _activado(primera.find(f"{w('trPr')}/{w('tblHeader')}") if primera is not None else None):
                problemas.append("la fila de encabezados no se repite en cada página")
            if self._bordes_verticales(tbl):
                problemas.append("tiene bordes verticales")
            if any(shd.get(w("fill")) not in SIN_FONDO for tc in tbl.iter(w("tc"))
                   for shd in tc.findall(f"{w('tcPr')}/{w('shd')}")):
                problemas.append("tiene colores de fondo")
            for problema in problemas:
                self.fallo_en(TABLAS, info, f"Tabla: {problema}.")

    # --- 6.2 Figuras --------------------------------------------------------------------------

    def figuras(self):
        """
        Cada figura del texto (un párrafo con solo una imagen, fuera de la portada y
        del índice, como en la regla), como la describe la web oficial de APA (Figure
        setup y Sample figures): sus leyendas, igual que las de las tablas.
        """
        bloques = estructura._bloques(self.doc)
        for k, p in enumerate(bloques):
            c = self.clasificado_de.get(p)
            if c is None or c.tipo != "imagen" or c.zona in leyendas.ZONAS_SIN_LEYENDAS_APA:
                continue
            self._leyendas_de(FIGURAS, leyendas.FIGURA, bloques, k, self.por_elemento[p])

    def _en_portada_o_indice(self, bloques, k):
        vecino = next((self.clasificado_de[b] for b in reversed(bloques[:k]) if b in self.clasificado_de), None)
        return vecino is not None and vecino.zona in leyendas.ZONAS_SIN_LEYENDAS_APA

    def _leyendas_de(self, regla, clase, bloques, k, info):
        """
        Encima del bloque k, «Tabla N» (o «Figura N») en negrita y, en la línea
        siguiente, el título en cursiva, los dos a la izquierda, sin sangría y con
        «Conservar con el siguiente»; debajo, si hay nota, «Nota.» en cursiva y el
        resto sin cursiva.
        """
        s = clase.singular
        encima = self._con_texto(bloques, k, -1, 2)
        numero, titulo = encima if len(encima) == 2 else (None, None)
        if numero is None or not clase.numero_o_letra.fullmatch(texto_parrafo(numero)):
            self.fallo_en(regla, info, f"La {s} no tiene encima su número («{clase.rotulo} 1») y, debajo, su título.")
        else:
            for p, negrita, cursiva, que in ((numero, True, False, "El número"), (titulo, False, True, "El título")):
                problemas = self._problemas_de_leyenda(p, negrita, cursiva, True)
                if problemas:
                    self.fallo_en(regla, self.por_elemento[p], f"{que} de la {s}: {'; '.join(problemas)}.")
        nota = next(iter(self._con_texto(bloques, k, +1, 1)), None)
        if nota is not None and estructura.NOTA.match(texto_parrafo(nota)):
            rotulo = re.match(r"Nota\.", texto_parrafo(nota))
            problemas = (self._problemas_de_leyenda(nota, False, rotulo.end(), False) if rotulo
                         else ["no empieza por «Nota.»"])
            if problemas:
                self.fallo_en(regla, self.por_elemento[nota], f"Nota de la {s}: {'; '.join(problemas)}.")

    def _con_texto(self, bloques, k, paso, cuantos):
        """Los `cuantos` párrafos con texto seguidos encima (paso -1) o debajo (+1) del bloque k, sin los vacíos."""
        encontrados, j = [], k + paso
        while 0 <= j < len(bloques) and len(encontrados) < cuantos and bloques[j].tag == w("p"):
            if texto_parrafo(bloques[j]).strip():
                encontrados.append(bloques[j])
            j += paso
        return encontrados[::-1] if paso < 0 else encontrados

    def _problemas_de_leyenda(self, p, negrita, cursiva, conservar):
        """`cursiva`: True (todo), False (nada) o hasta qué posición del texto (el rótulo «Nota.»)."""
        problemas = []
        if {"start": "left"}.get(self._alineacion(p), self._alineacion(p)) != "left":
            problemas.append("no está a la izquierda")
        if any(_entero(self.estilos.parrafo(p, "ind", a)) for a in ("firstLine", "hanging", "left")):
            problemas.append("tiene sangría")
        if conservar and not _activado(self._capa_de_parrafo(p, "keepNext")):
            problemas.append("no tiene «Conservar con el siguiente»")
        desde = 0
        for r in fuente._runs_propios(p):
            texto = "".join(t.text or "" for t in r.iter(w("t")))
            quiere = cursiva if isinstance(cursiva, bool) else desde < cursiva
            desde += len(texto)
            if not texto.strip():
                continue
            if _activado(self._de_texto(r, p, "b")) != negrita:
                problemas.append("va en negrita" if not negrita else "no va en negrita")
            if _activado(self._de_texto(r, p, "i")) != quiere:
                problemas.append("«Nota.» va en cursiva y el resto no" if not isinstance(cursiva, bool)
                                 else "va en cursiva" if not quiere else "no va en cursiva")
        return list(dict.fromkeys(problemas))

    def _bordes_verticales(self, tbl):
        """¿Se ve alguna línea vertical? (bordes de la tabla, de su estilo o de sus celdas)"""
        tblpr = tbl.find(w("tblPr"))
        estilo = tblpr.find(w("tblStyle")) if tblpr is not None else None
        cadena = self.estilos.cadena(estilo.get(w("val"))) if estilo is not None else []
        capas = [tblpr] + [e.find(w("tblPr")) for e in cadena]
        for lado in ("left", "right", "insideV"):
            borde = next((b for capa in capas if capa is not None
                          for b in capa.findall(f"{w('tblBorders')}/{w(lado)}")), None)
            if borde is not None and borde.get(w("val")) not in SIN_LINEA:
                return True
        return any(b.get(w("val")) not in SIN_LINEA for tc in tbl.iter(w("tc"))
                   for lado in ("left", "right", "start", "end")
                   for b in tc.findall(f"{w('tcPr')}/{w('tcBorders')}/{w(lado)}"))

    # --- 6.3 Citas en bloque ----------------------------------------------------------------------

    def citas(self):
        """
        Cada cita en bloque (párrafos con sangría a la izquierda seguidos, de 40
        palabras o más entre todos), como la describe la web oficial de APA
        (Quotations): todo el bloque sangrado 1,27 cm a la izquierda, nada a la
        derecha, sin sangría de primera línea (salvo en los párrafos segundo y
        siguientes, que llevan 1,27 cm más) y sin cursiva en todo el texto. Las
        comillas no se quitan (cambiaría el texto): son una excepción admitida.
        """
        for bloque in regla_citas.bloques_de_cita(self.clasificados):
            if regla_citas.palabras(bloque) < estructura.PALABRAS_DE_CITA_EN_BLOQUE:
                continue
            for n, c in enumerate(bloque):
                p, info = c.elemento, self.por_elemento[c.elemento]
                problemas = []
                izquierda = max(_entero(self.estilos.parrafo(p, "ind", a)) for a in ("left", "start"))
                if izquierda != regla_citas.SANGRIA or _entero(self.estilos.parrafo(p, "ind", "leftChars")):
                    problemas.append("todo el bloque no está sangrado 1,27 cm a la izquierda")
                if any(_entero(self.estilos.parrafo(p, "ind", a)) for a in ("right", "end", "rightChars")):
                    problemas.append("tiene sangría a la derecha")
                esperada = 0 if n == 0 else regla_citas.SANGRIA
                if self._primera_linea(p) != esperada or sangria._sangria(self.estilos, p, "hanging"):
                    problemas.append("la primera línea no lleva sangría" if n == 0 else
                                     "la primera línea de este párrafo no lleva 1,27 cm más que el resto")
                for problema in problemas:
                    self.fallo_en(CITAS, info, f"Cita en bloque: {problema}.")
            if regla_citas._cursiva_en_todo(self.estilos, bloque):
                self.fallo_en(CITAS, self.por_elemento[bloque[0].elemento], "Cita en bloque: va toda en cursiva.")
            if regla_citas.entre_comillas(bloque):
                self.excepcion_en(CITAS, self.por_elemento[bloque[0].elemento],
                                  "Cita en bloque entre comillas: APA no las pone, pero quitarlas cambiaría el texto "
                                  "(se avisa en el informe).")

    def _sigue_a_una_cita(self, p):
        """¿Es el párrafo segundo o siguiente de una cita en bloque? (lleva sangría de primera línea: 6.3)"""
        for bloque in regla_citas.bloques_de_cita(self.clasificados):
            if regla_citas.palabras(bloque) >= estructura.PALABRAS_DE_CITA_EN_BLOQUE and \
                    any(c.elemento is p for c in bloque[1:]):
                return True
        return False

    # --- 6.4 Referencias ----------------------------------------------------------------------------

    def referencias(self):
        """
        La lista de referencias, como la describe la web oficial de APA (Reference
        list setup): empieza en una página nueva, con el rótulo «Referencias» (o
        «Referencia») centrado y en negrita, y cada referencia con sangría francesa de
        1,27 cm. Otro rótulo («Bibliografía») no es un incumplimiento sino una
        excepción admitida: se deja con --rotulo-tal-cual o si hay varias listas, y se
        avisa en el informe.
        """
        for n, (rotulo, referencias) in enumerate(regla_referencias.listas(self.clasificados)):
            p, info = rotulo.elemento, self.por_elemento[rotulo.elemento]
            texto = texto_parrafo(p)
            if texto not in documento.ROTULOS_APA_DE_REFERENCIAS:
                self.excepcion_en(REFERENCIAS, info, "Rótulo distinto de «Referencias»: se ha dejado (se avisa en el "
                                                     "informe; sin --rotulo-tal-cual y con una sola lista, se cambia).")
            problemas = []
            if self._alineacion(p) != "center":
                problemas.append("no está centrado")
            if any(not _activado(self._de_texto(r, p, "b")) for r in titulos.fragmentos_con_letras(p)):
                problemas.append("no va en negrita")
            anterior = next((i for i in reversed(self.principales) if i.numero < info.numero), None)
            if n == 0 and not alineacion._salto_antes(self.estilos, p) and \
                    not (anterior is not None and regla_portada._termina_en_salto(anterior.p)):
                problemas.append("la lista no empieza en una página nueva")
            for problema in problemas:
                self.fallo_en(REFERENCIAS, info, f"Rótulo de las referencias: {problema}.")
            for c in referencias:
                q = c.elemento
                izquierda = max(_entero(self.estilos.parrafo(q, "ind", a)) for a in ("left", "start"))
                francesa = _entero(self.estilos.parrafo(q, "ind", "hanging"))
                if izquierda != regla_referencias.SANGRIA or francesa != regla_referencias.SANGRIA or \
                        _entero(self.estilos.parrafo(q, "ind", "firstLine")):
                    self.fallo_en(REFERENCIAS, self.por_elemento[q], "Referencia sin la sangría francesa de 1,27 cm.")

    # --- 6.5 Orden alfabético ------------------------------------------------------------------------

    def orden(self):
        """
        Las referencias, en orden alfabético (la clave de reglas/orden.py: autores
        palabra a palabra, año y título; las direcciones web sueltas, al final). Con
        varias listas no se ordenan: es una excepción admitida (se avisa).
        """
        todas = regla_referencias.listas(self.clasificados)
        if len(todas) > 1:
            rotulo = todas[0][0]
            self.excepcion_en(ORDEN, self.por_elemento[rotulo.elemento], f"Hay {len(todas)} listas de referencias: no "
                                                                         "se ordenan hasta que se unan (se avisa).")
            return
        for _, referencias in todas:
            claves = [regla_orden.clave(texto_parrafo(c.elemento)) for c in referencias]
            for anterior, (c, k) in zip(claves, list(zip(referencias, claves))[1:]):
                if k < anterior:
                    self.fallo_en(ORDEN, self.por_elemento[c.elemento],
                                  "Referencia fuera del orden alfabético (va antes que la anterior).")
                    break

    # --- 6.7 Apéndices ------------------------------------------------------------------------------

    def apendices(self):
        """
        Cada apéndice, como lo describe la web oficial de APA (Appendices setup): en
        una página nueva, con el rótulo y, en la línea siguiente, el título, los dos
        centrados y en negrita. «Anexo» en lugar de «Apéndice» es una excepción
        admitida (se avisa; --rotulo-apendice lo cambia).
        """
        todos = regla_apendices.apendices(self.clasificados)
        for n, (rotulo, titulo) in enumerate(todos):
            p, info = rotulo.elemento, self.por_elemento[rotulo.elemento]
            texto = texto_parrafo(p).strip()
            if regla_apendices.ROTULO_DE_APENDICE.match(texto)["titulo"]:
                self.fallo_en(APENDICES, info, "El rótulo y el título del apéndice van en la misma línea.")
                continue
            if texto != regla_apendices.rotulo_apa(n, len(todos)):
                self.excepcion_en(APENDICES, info, f"Rótulo «{texto}» en lugar de "
                                                   f"«{regla_apendices.rotulo_apa(n, len(todos))}»: se deja (se avisa en "
                                                   "el informe; --rotulo-apendice lo cambia).")
            anterior = next((i for i in reversed(self.principales) if i.numero < info.numero), None)
            if not alineacion._salto_antes(self.estilos, p) and \
                    not (anterior is not None and regla_portada._termina_en_salto(anterior.p)):
                self.fallo_en(APENDICES, info, "El apéndice no empieza en una página nueva.")
            if titulo is None:
                self.fallo_en(APENDICES, info, "El apéndice no tiene título en la línea siguiente al rótulo.")
            for q, que in ((p, "El rótulo"), (titulo.elemento if titulo is not None else None, "El título")):
                if q is None:
                    continue
                problemas = []
                if self._alineacion(q) != "center":
                    problemas.append("no está centrado")
                if any(not _activado(self._de_texto(r, q, "b")) for r in titulos.fragmentos_con_letras(q)):
                    problemas.append("no va en negrita")
                if problemas:
                    self.fallo_en(APENDICES, self.por_elemento[q], f"{que} del apéndice: {'; '.join(problemas)}.")

    def _con_estilo_title(self, p):
        estilo = self.estilos.por_id.get(self.estilos.estilo_de(p))
        return estilo is not None and titulos._nombre(estilo) == titulos.NOMBRE_DEL_ESTILO_DEL_TITULO

    def _capa_de_parrafo(self, p, etiqueta):
        """El primer `etiqueta` en las capas de formato del párrafo (directo, estilos, por defecto)."""
        for capa in self.estilos._capas_parrafo(p):
            el = capa.find(w(etiqueta)) if capa is not None else None
            if el is not None:
                return el
        return None


def verificar(origen):
    """
    Comprueba las reglas de la fase 1 y los títulos (pasos 4.1 a 4.5) en un documento: una ruta a un .docx o un
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
    L.append("\nResultado: " + ("cumple todas las reglas comprobadas." if resultado.cumple else
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
