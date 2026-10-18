"""
Paso 4.1: estilos de título de Word (Título 1 a Título 5) con el formato APA,
cada título con el estilo de su nivel y avisos de estructura.

Los casos concretos se prueban con documentos creados aquí mismo; después se
comprueba la regla con los 5 documentos de pruebas/ (que no se modifican).
"""

import copy
from pathlib import Path

import docx
import pytest
from docx.enum.text import WD_BREAK
from docx.shared import Pt

from apa_facil import estructura, pipeline
from apa_facil.documento import CambioPrevisto, Estilos, comparar, radiografia, texto_parrafo, w
from apa_facil.informe import Informe
from apa_facil.reglas import titulos
from herramientas.verificar import (FORMATO_TITULOS, MAYUSCULAS_TITULOS, PORTADA_PROPIA, TITULO_TRABAJO, TITULOS, TITULOS_EN_LINEA,
                                   verificar)

PRUEBAS = sorted((Path(__file__).resolve().parent.parent / "pruebas").glob("*.docx"))
TEXTO = ("Este es un párrafo normal de texto con bastantes palabras para que no parezca un título "
         "ni una línea de portada, y que termina en punto como cualquier frase.")


def documento_con_portada():
    """Un documento cuya primera página es una portada (para que no estorbe en las pruebas)."""
    doc = docx.Document()
    doc.add_paragraph("Portada").runs[0].add_break(WD_BREAK.PAGE)
    return doc


def aplicar(doc):
    informe = Informe()
    titulos.aplicar(doc, {}, informe)
    return informe


def titulo_falso(doc, texto):
    """Un título hecho a mano: corto, en negrita, con el estilo Normal y seguido de texto."""
    p = doc.add_paragraph()
    p.add_run(texto).bold = True
    doc.add_paragraph(TEXTO)
    return p


def estilo_de(doc, p):
    return Estilos(doc).estilo_de(p._p)


def unico(doc, nivel):
    estilos = titulos.estilos_de_titulo(doc)[nivel]
    assert len(estilos) == 1
    return estilos[0]


def propio(estilo, capa, etiqueta, atributo="val"):
    el = estilo.find(f"{w(capa)}/{w(etiqueta)}")
    if el is None:
        return None
    return el.get(w(atributo), "activado")


def sin_el_titulo(mensajes):
    """
    Los mensajes del informe sin los del título del trabajo (paso 4.4): la portada de
    estos documentos de prueba («Portada») se repite al principio del texto.
    """
    return [m for m in mensajes if "título del trabajo" not in m.lower()]


def previstos_sin_el_titulo(informe):
    return [c for c in informe.cambios_previstos if c.accion != "añadir_titulo"]


def problemas_del_verificador(doc, regla=TITULOS):
    return [h.mensaje for h in verificar(doc).incumplimientos if h.regla == regla]


# --- Formato de los estilos ---------------------------------------------------------------

@pytest.mark.parametrize("nivel, alineacion, cursiva, sangria", [
    (1, "center", "0", "0"),
    (2, "left", "0", "0"),
    (3, "left", "activado", "0"),
    (4, "left", "0", "720"),
    (5, "left", "activado", "720"),
])
def test_formato_apa_de_cada_nivel(nivel, alineacion, cursiva, sangria):
    doc = documento_con_portada()
    doc.add_heading("Un título", level=nivel)
    aplicar(doc)
    estilo = unico(doc, nivel)
    assert propio(estilo, "pPr", "jc") == alineacion
    assert propio(estilo, "rPr", "b") == "activado"
    assert propio(estilo, "rPr", "i") == cursiva
    assert propio(estilo, "pPr", "ind", "firstLine") == sangria
    assert propio(estilo, "pPr", "keepNext") == "activado", "no se queda solo al final de la página"
    assert propio(estilo, "pPr", "outlineLvl") == str(nivel - 1), "nivel de esquema: panel de navegación"


def test_se_quitan_color_subrayado_mayusculas_y_bordes():
    doc = documento_con_portada()
    doc.add_heading("Un título", level=1)
    estilo = unico(doc, 1)
    rpr, ppr = estilo.get_or_add_rPr(), estilo.get_or_add_pPr()
    for etiqueta in ("u", "caps", "smallCaps"):
        rpr.append(copy.deepcopy(rpr.makeelement(w(etiqueta), {w("val"): "single"} if etiqueta == "u" else {})))
    ppr.append(ppr.makeelement(w("pBdr"), {}))
    assert estilo.find(f"{w('rPr')}/{w('color')}") is not None   # el azul de Word
    aplicar(doc)
    for etiqueta in ("color", "u", "caps", "smallCaps"):
        assert estilo.find(f"{w('rPr')}/{w(etiqueta)}") is None, etiqueta
    assert estilo.find(f"{w('pPr')}/{w('pBdr')}") is None


def test_el_estilo_de_caracter_vinculado_pierde_el_color():
    doc = documento_con_portada()
    doc.add_heading("Un título", level=1)
    vinculado = doc.styles.element.makeelement(w("style"), {w("type"): "character", w("styleId"): "Heading1Char"})
    rpr = vinculado.makeelement(w("rPr"), {})
    rpr.append(rpr.makeelement(w("color"), {w("val"): "2E74B5"}))
    vinculado.append(rpr)
    doc.styles.element.append(vinculado)
    unico(doc, 1).get_or_add_pPr().addprevious(unico(doc, 1).makeelement(w("link"), {w("val"): "Heading1Char"}))
    aplicar(doc)
    assert vinculado.find(f"{w('rPr')}/{w('color')}") is None


def test_estilos_duplicados_quedan_en_uno():
    """En 02, 03 y 05 hay dos definiciones de «Heading 1»: el resultado dependería de cuál lea Word."""
    doc = documento_con_portada()
    p = doc.add_heading("Un título", level=1)
    doc.styles.element.append(copy.deepcopy(unico(doc, 1)))
    assert len(titulos.estilos_de_titulo(doc)[1]) == 2
    assert any("definido varias veces" in m for m in problemas_del_verificador(doc))
    aplicar(doc)
    unico(doc, 1)
    assert estilo_de(doc, p) == "Heading1"
    assert problemas_del_verificador(doc) == []


def test_estilo_que_falta_se_crea_basado_en_normal():
    doc = documento_con_portada()
    for estilo in titulos.estilos_de_titulo(doc)[1]:
        estilo.getparent().remove(estilo)
    p = titulo_falso(doc, "Introducción")
    informe = aplicar(doc)
    estilo = unico(doc, 1)
    assert estilo.get(w("styleId")) == "Heading1"
    assert propio(estilo, "basedOn", "basedOn") is None   # basedOn es hijo directo del estilo
    assert estilo.find(w("basedOn")).get(w("val")) == "Normal"
    assert estilo_de(doc, p) == "Heading1"
    assert propio(estilo, "pPr", "jc") == "center"
    assert any("1 título («Introducción») usa ahora el estilo de título" in c for c in informe.cambios)


def test_estilo_con_identificador_en_espanol():
    """Word en español guarda «Título 1» con el identificador «Ttulo1» y el nombre interno «heading 1»."""
    doc = documento_con_portada()
    estilo = unico(doc, 1)
    estilo.set(w("styleId"), "Ttulo1")
    estilo.find(w("name")).set(w("val"), "heading 1")
    p = titulo_falso(doc, "Introducción")
    aplicar(doc)
    assert unico(doc, 1).get(w("styleId")) == "Ttulo1", "no se crea otro «Título 1»"
    assert estilo_de(doc, p) == "Ttulo1"


def test_numeracion_automatica_del_estilo_solo_se_avisa():
    doc = documento_con_portada()
    doc.add_heading("Un título", level=1)
    ppr = unico(doc, 1).get_or_add_pPr()
    ppr.get_or_add_numPr().get_or_add_numId().set(w("val"), "1")
    informe = aplicar(doc)
    assert ppr.find(w("numPr")) is not None
    assert any("numera automáticamente" in a for a in informe.avisos)


# --- Asignación ---------------------------------------------------------------------------

def test_titulos_falsos_reciben_el_estilo_de_su_nivel():
    doc = documento_con_portada()
    p1 = titulo_falso(doc, "1. Marco teórico")
    p2 = titulo_falso(doc, "1.1. Antecedentes")
    p3 = titulo_falso(doc, "1.2. Estado de la cuestión")
    aplicar(doc)
    assert [estilo_de(doc, p) for p in (p1, p2, p3)] == ["Heading1", "Heading2", "Heading2"]
    assert problemas_del_verificador(doc) == []


def test_los_que_ya_tenian_su_estilo_no_cuentan_como_cambio():
    doc = documento_con_portada()
    doc.add_heading("Método", level=1)
    doc.add_paragraph(TEXTO)
    informe = aplicar(doc)
    cambios = sin_el_titulo(informe.cambios)
    assert len(cambios) == 1 and cambios[0].startswith("Títulos: los estilos")


def test_titulo_dudoso_no_recibe_estilo_y_se_avisa():
    doc = documento_con_portada()
    p = titulo_falso(doc, "Esto es importante.")
    c = next(c for c in estructura.clasificar(doc) if c.elemento is p._p)
    assert (c.tipo, c.confianza) == ("titulo", estructura.BAJA)
    informe = aplicar(doc)
    assert estilo_de(doc, p) == "Normal"
    assert sin_el_titulo(informe.avisos) == ["«Esto es importante.» parece un título, pero no es seguro, así que se ha dejado "
                              "como estaba. Si es un título, ponle en Word el estilo «Título 1»."]
    resultado = verificar(doc)
    assert [h.regla for h in resultado.excepciones if h.regla != PORTADA_PROPIA] == [TITULOS]
    assert not [h for h in resultado.incumplimientos if h.regla == TITULOS]


def test_sin_titulos_no_se_informa_de_nada():
    doc = documento_con_portada()
    doc.add_paragraph(TEXTO)
    informe = aplicar(doc)
    assert sin_el_titulo(informe.cambios) == sin_el_titulo(informe.avisos) == []


def test_el_texto_no_cambia():
    doc = documento_con_portada()
    titulo_falso(doc, "INTRODUCCIÓN")
    titulo_falso(doc, "1.1. Antecedentes")
    doc.add_heading("Método", level=1)
    antes = radiografia(doc)
    informe = aplicar(doc)
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []


def test_el_verificador_ve_un_titulo_sin_estilo():
    doc = documento_con_portada()
    titulo_falso(doc, "Introducción")
    assert "Título de nivel 1 sin el estilo «Título 1» de Word." in problemas_del_verificador(doc)


def test_el_verificador_ve_un_estilo_estropeado():
    doc = documento_con_portada()
    doc.add_heading("Método", level=1)
    doc.add_heading("Participantes", level=2)
    aplicar(doc)
    assert problemas_del_verificador(doc) == []
    unico(doc, 2).find(f"{w('rPr')}/{w('b')}").set(w("val"), "0")
    unico(doc, 1).find(f"{w('pPr')}/{w('keepNext')}").getparent().remove(unico(doc, 1).find(f"{w('pPr')}/{w('keepNext')}"))
    assert problemas_del_verificador(doc) == [
        "El estilo «Título 1» no tiene el formato APA: no tiene «Conservar con el siguiente».",
        "El estilo «Título 2» no tiene el formato APA: no va en negrita.",
    ]


# --- 4.2 Formato puesto a mano ----------------------------------------------------------------

def titulo(doc, texto, nivel):
    p = doc.add_heading(texto, level=nivel)
    doc.add_paragraph(TEXTO)
    return p


def poner(el, etiqueta, **atributos):
    """Añade a mano un ajuste de formato (w:jc, w:i…) a unas propiedades."""
    hijo = el.makeelement(w(etiqueta), {w(k): v for k, v in atributos.items()})
    el.append(hijo)
    return hijo


def formato_corregido(informe):
    return [c for c in informe.cambios if "formato puesto a mano" in c]


def test_alineacion_a_mano_que_contradice_el_nivel():
    doc = documento_con_portada()
    p = titulo(doc, "Participantes", 2)
    poner(p._p.get_or_add_pPr(), "jc", val="center")
    informe = aplicar(doc)
    assert p._p.pPr.find(w("jc")) is None
    assert formato_corregido(informe) == ["Títulos: se ha quitado el formato puesto a mano que no corresponde a "
                                          "su nivel en 1 título («Participantes»: la alineación)."]
    assert problemas_del_verificador(doc, FORMATO_TITULOS) == []


def test_lo_que_coincide_con_el_nivel_no_se_cuenta():
    doc = documento_con_portada()
    p1 = titulo(doc, "Método", 1)
    poner(p1._p.get_or_add_pPr(), "jc", val="center")
    p4 = titulo(doc, "Edad.", 4)
    p4._p.get_or_add_pPr().get_or_add_ind().set(w("firstLine"), "720")
    for r in p1.runs:
        r.bold = True
    informe = aplicar(doc)
    assert p1._p.pPr.find(w("jc")) is not None and p4._p.pPr.find(w("ind")) is not None
    assert formato_corregido(informe) == []


def test_sangria_a_mano_en_un_titulo():
    doc = documento_con_portada()
    p1 = titulo(doc, "Método", 1)
    titulo(doc, "Muestra", 2)
    titulo(doc, "Edad", 3)
    p4 = titulo(doc, "Niños.", 4)
    aplicar(doc)   # los estilos, ya con el formato APA
    p1._p.get_or_add_pPr().get_or_add_ind().set(w("left"), "720")
    p4._p.get_or_add_pPr().get_or_add_ind().set(w("firstLine"), "0")
    assert problemas_del_verificador(doc, FORMATO_TITULOS) == [
        "Título de nivel 1: la sangría no es la de su nivel.", "Título de nivel 4: la sangría no es la de su nivel."]
    informe = aplicar(doc)
    assert p1._p.pPr.find(w("ind")) is None and p4._p.pPr.find(w("ind")) is None
    assert "«Método»: la sangría; «Niños.»: la sangría" in formato_corregido(informe)[0]
    assert problemas_del_verificador(doc, FORMATO_TITULOS) == []


def test_cursiva_en_todo_el_titulo_se_quita():
    doc = documento_con_portada()
    p = titulo(doc, "Método", 1)
    p.runs[0].italic = True
    informe = aplicar(doc)
    assert p.runs[0].italic is None
    assert "«Método»: la cursiva)" in formato_corregido(informe)[0]


def test_cursiva_desactivada_en_un_nivel_3():
    doc = documento_con_portada()
    titulo(doc, "Método", 1)
    titulo(doc, "Muestra", 2)
    p = titulo(doc, "Edad", 3)
    p.runs[0].italic = False
    informe = aplicar(doc)
    assert p.runs[0].italic is None
    assert "«Edad»: la cursiva desactivada" in formato_corregido(informe)[0]


def test_cursiva_en_una_parte_del_titulo_se_respeta():
    """APA usa la cursiva dentro de un título para el título de un libro o el nombre de una especie."""
    doc = documento_con_portada()
    p = doc.add_heading("Lecturas de ", level=1)
    p.add_run("Cien años de soledad").italic = True
    doc.add_paragraph(TEXTO)
    informe = aplicar(doc)
    assert p.runs[1].italic is True
    assert formato_corregido(informe) == []
    assert problemas_del_verificador(doc, FORMATO_TITULOS) == []


def test_negrita_desactivada_mayusculas_bordes_y_conservar():
    doc = documento_con_portada()
    p = titulo(doc, "Método", 1)
    rpr = p.runs[0]._r.get_or_add_rPr()
    poner(rpr, "b", val="0")
    poner(rpr, "smallCaps")
    ppr = p._p.get_or_add_pPr()
    poner(ppr, "keepNext", val="0")
    poner(ppr, "pBdr")
    antes = radiografia(doc)
    assert problemas_del_verificador(doc, FORMATO_TITULOS)
    informe = aplicar(doc)
    for etiqueta in ("b", "smallCaps"):
        assert rpr.find(w(etiqueta)) is None, etiqueta
    for etiqueta in ("keepNext", "pBdr"):
        assert ppr.find(w(etiqueta)) is None, etiqueta
    assert formato_corregido(informe) == [
        "Títulos: se ha quitado el formato puesto a mano que no corresponde a su nivel en 1 título («Método»: los "
        "bordes o el sombreado, «Conservar con el siguiente» desactivado, la negrita desactivada y las "
        "mayúsculas forzadas)."]
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == [], \
        "las mayúsculas son formato: el texto no cambia"
    assert problemas_del_verificador(doc, FORMATO_TITULOS) == []


def test_el_verificador_ve_el_formato_a_mano():
    doc = documento_con_portada()
    p = titulo(doc, "Método", 1)
    aplicar(doc)
    p.runs[0].italic = True
    p.runs[0].underline = True
    assert problemas_del_verificador(doc, FORMATO_TITULOS) == ["Título de nivel 1: va en cursiva; va subrayado."]


def test_numeracion_de_word_en_un_titulo_solo_se_avisa():
    doc = documento_con_portada()
    p = titulo(doc, "Método", 1)
    p._p.get_or_add_pPr().get_or_add_numPr().get_or_add_numId().set(w("val"), "1")
    informe = aplicar(doc)
    assert p._p.pPr.find(w("numPr")) is not None
    assert [a for a in informe.avisos if "numeración automática" in a] == [
        "1 título («Método») lleva la numeración automática de Word. APA no numera los títulos, pero se ha "
        "dejado así: quítala en Word si no la exige tu universidad."]


# --- 4.5 Niveles 4 y 5 en la misma línea que su texto ----------------------------------------

def con_niveles(*titulos_de_nivel):
    """Un documento con títulos de nivel 1 a 3 y, debajo, los de nivel 4 y 5 indicados, cada uno con su texto."""
    doc = documento_con_portada()
    for texto_titulo, nivel in (("Método", 1), ("Muestra", 2), ("Edad", 3), *titulos_de_nivel):
        titulo(doc, texto_titulo, nivel)
    return doc


def parrafo_que_empieza(doc, inicio):
    return next(p for p in doc.element.body.iter(w("p")) if texto_parrafo(p).startswith(inicio))


def el_parrafo(doc, inicio):
    """El mismo párrafo, como objeto de python-docx (con sus fragmentos de texto, runs)."""
    return next(par for par in doc.paragraphs if par.text.startswith(inicio))


@pytest.mark.parametrize("original, queda", [
    ("Niños", "Niños. "),                      # se añade el punto y un espacio
    ("Niñas.", "Niñas. "),                     # ya tenía el punto: solo el espacio
    ("¿Por qué leen?", "¿Por Qué Leen? "),     # con interrogación, APA no añade punto (y 4.6: mayúsculas)
    ("Los «nativos digitales»", "Los «Nativos Digitales». "),
])
def test_nivel_4_en_la_misma_linea_con_su_punto(original, queda):
    doc = con_niveles((original, 4))
    antes = radiografia(doc)
    informe = aplicar(doc)
    p = parrafo_que_empieza(doc, original[:4])
    assert texto_parrafo(p) == queda
    assert titulos.en_linea(p), "la marca de párrafo, oculta (separador de estilo de Word)"
    assert texto_parrafo(p.getnext()) == TEXTO, "el texto de debajo no cambia"
    assert Estilos(doc).nombre(Estilos(doc).estilo_de(p)).lower() == "heading 4", "sigue siendo un título"
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert problemas_del_verificador(doc, TITULOS_EN_LINEA) == []


def test_nivel_5_tambien_y_el_punto_va_con_el_formato_del_titulo():
    doc = con_niveles(("Niños", 4), ("Menores de 6", 5))
    par = el_parrafo(doc, "Menores de 6")
    par.runs[0].italic = True
    p = par._p
    informe = aplicar(doc)
    assert texto_parrafo(p) == "Menores de 6. " and titulos.en_linea(p)
    assert p.findall(w("r"))[-1].find(f"{w('rPr')}/{w('i')}") is not None, "el punto, con el formato del título"
    assert any(c.startswith("Títulos: 2 títulos de nivel 4 o 5 («Niños», «Menores de 6») van ahora en la "
                            "misma línea que su texto y terminan en punto") for c in informe.cambios)


def test_con_dos_puntos_se_une_pero_se_avisa():
    doc = con_niveles(("Factores de riesgo:", 4))
    informe = aplicar(doc)
    p = parrafo_que_empieza(doc, "Factores")
    assert texto_parrafo(p) == "Factores de Riesgo: " and titulos.en_linea(p)
    assert any("termina en dos puntos o punto y coma" in a for a in informe.avisos)
    assert [h.regla for h in verificar(doc).excepciones if h.regla == TITULOS_EN_LINEA] == [TITULOS_EN_LINEA]


@pytest.mark.parametrize("detras", ["tabla", "titulo", "lista", "nada"])
def test_si_no_le_sigue_texto_no_se_une(detras):
    doc = con_niveles()
    p = doc.add_heading("Niños", level=4)
    if detras == "tabla":
        doc.add_table(rows=1, cols=1)
    elif detras == "titulo":
        doc.add_heading("Niñas", level=4)
        doc.add_paragraph(TEXTO)
    elif detras == "lista":
        doc.add_paragraph("primer elemento", style="List Bullet")
    antes = radiografia(doc)
    informe = aplicar(doc)
    assert texto_parrafo(p._p) == "Niños" and not titulos.en_linea(p._p)
    assert any(a.startswith("1 título de nivel 4 o 5 («Niños») no va seguido de un párrafo de texto")
               for a in informe.avisos)
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert problemas_del_verificador(doc, TITULOS_EN_LINEA) == []


def test_un_marcador_entre_el_titulo_y_el_texto_no_estorba():
    doc = con_niveles(("Niños", 4))
    p = parrafo_que_empieza(doc, "Niños")
    p.addnext(p.makeelement(w("bookmarkEnd"), {w("id"): "0"}))
    aplicar(doc)
    assert titulos.en_linea(p)


def test_la_integridad_solo_acepta_el_punto_y_el_espacio():
    doc = con_niveles(("Niños", 4))
    antes = radiografia(doc)
    numero = next(i for i, (_, t) in enumerate(antes.textos, start=1) if t == "Niños")
    par = el_parrafo(doc, "Niños")
    par.runs[0].text = "Niños. "
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "añadir_al_final", nuevo=". ")]) == []
    par.runs[0].text = "Niños y niñas. "
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "añadir_al_final", nuevo=" y niñas. ")])
    doc2 = con_niveles(("Niños.", 4))
    antes2 = radiografia(doc2)
    el_parrafo(doc2, "Niños").runs[0].text = "Niños.. "
    assert comparar(antes2, radiografia(doc2), [CambioPrevisto(numero, "añadir_al_final", nuevo=". ")]), \
        "no se pone un segundo punto"


def test_el_verificador_ve_un_nivel_4_que_no_va_en_linea():
    doc = con_niveles(("Niños", 4))
    aplicar(doc)
    p = parrafo_que_empieza(doc, "Niños")
    p.find(f"{w('pPr')}/{w('rPr')}").getparent().remove(p.find(f"{w('pPr')}/{w('rPr')}"))
    assert problemas_del_verificador(doc, TITULOS_EN_LINEA) == [
        "Título de nivel 4: no va en la misma línea que su texto."]


# --- 4.3 Numeración y mayúsculas --------------------------------------------------------------

TAL_CUAL = {"titulos_tal_cual": True}   # para probar la numeración sin las mayúsculas de APA (paso 4.6)


def aplicar_con(doc, opciones):
    informe = Informe()
    titulos.aplicar(doc, opciones, informe)
    return informe


def texto(p):
    return texto_parrafo(p._p)


@pytest.mark.parametrize("titulo_original, sin_numeracion", [
    ("1. Introducción", "Introducción"),
    ("1.1. Antecedentes", "Antecedentes"),
    ("1.1 Antecedentes", "Antecedentes"),
    ("2.3.1. Muestra", "Muestra"),
    ("3) Resultados", "Resultados"),
    ("IV. Discusión", "Discusión"),
    ("a) Evolución demográfica", "Evolución demográfica"),
    ("20 años de democracia", "20 años de democracia"),   # el número es parte del título
    ("2020: un año difícil", "2020: un año difícil"),
])
def test_que_numeracion_se_quita(titulo_original, sin_numeracion):
    assert (titulos.numeracion(titulo_original) is None) == (titulo_original == sin_numeracion)
    doc = documento_con_portada()
    p = titulo(doc, titulo_original, 1)
    antes = radiografia(doc)
    informe = aplicar_con(doc, {"quitar_numeracion": True, **TAL_CUAL})
    assert texto(p) == sin_numeracion
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []


def test_por_defecto_la_numeracion_solo_se_avisa_y_las_mayusculas_se_cambian():
    doc = documento_con_portada()
    p1 = titulo(doc, "1. Introducción", 1)
    p2 = titulo(doc, "2. MÉTODO", 1)
    p3 = titulo(doc, "3. Análisis de los datos", 1)
    antes = radiografia(doc)
    informe = aplicar(doc)
    assert (texto(p1), texto(p2), texto(p3)) == ("1. Introducción", "2. Método", "3. Análisis de los Datos")
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert [c.accion for c in previstos_sin_el_titulo(informe)] == ["mayusculas_apa", "mayusculas_apa"]
    assert [a for a in informe.avisos if "numeración" in a] == [
        "3 títulos empiezan con una numeración («1. Introducción», «2. MÉTODO», «3. Análisis de los datos»). APA no "
        "numera los títulos, pero muchas universidades lo exigen, así que se ha dejado como estaba. Si la tuya no lo "
        "pide, la herramienta puede quitarla: usa la opción --quitar-numeracion."]
    assert ("Títulos: 2 títulos llevan ahora mayúscula inicial en las palabras principales, como pide APA "
            "(«2. MÉTODO» → «2. Método», «3. Análisis de los datos» → «3. Análisis de los Datos»). Se han respetado "
            "los nombres propios y las siglas que el trabajo escribe así en el texto; revisa que hayan quedado bien. "
            "Si prefieres dejar los títulos como estaban, usa la opción --titulos-tal-cual.") in informe.cambios


def test_titulos_tal_cual():
    doc = documento_con_portada()
    p = titulo(doc, "2. Marco teórico", 1)
    informe = aplicar_con(doc, TAL_CUAL)
    assert texto(p) == "2. Marco teórico" and previstos_sin_el_titulo(informe) == []
    assert any(a.startswith("1 título no lleva las mayúsculas de APA («2. Marco teórico»)") and
               "--titulos-tal-cual" in a for a in informe.avisos)


def test_las_dos_cosas_son_independientes():
    doc = documento_con_portada()
    p = titulo(doc, "2. MÉTODO", 1)
    aplicar_con(doc, {})
    assert texto(p) == "2. Método"
    doc = documento_con_portada()
    p = titulo(doc, "2. MÉTODO", 1)
    aplicar_con(doc, {"quitar_numeracion": True, **TAL_CUAL})
    assert texto(p) == "MÉTODO"
    doc = documento_con_portada()
    p = titulo(doc, "2. MÉTODO", 1)
    aplicar_con(doc, {"quitar_numeracion": True})
    assert texto(p) == "Método"


def test_numeracion_repartida_en_varios_fragmentos():
    doc = documento_con_portada()
    p = doc.add_heading("", level=1)
    for trozo in ("1", ".", "1. ", "Antece", "dentes"):
        p.add_run(trozo)
    p.runs[3].italic = True
    doc.add_paragraph(TEXTO)
    antes = radiografia(doc)
    informe = aplicar_con(doc, {"quitar_numeracion": True, **TAL_CUAL})
    assert texto(p) == "Antecedentes"
    assert [r.text for r in p.runs] == ["", "", "", "Antece", "dentes"] and p.runs[3].italic, "el formato no se mueve"
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []


def test_mayusculas_conservan_nombres_propios_siglas_y_romanos():
    doc = documento_con_portada()
    p = titulo(doc, "LA UE Y ESPAÑA EN EL SIGLO XXI", 1)
    doc.add_paragraph("Desde que España entró en la UE, el país ha cambiado. " + TEXTO)
    doc.add_paragraph("Hoy el estado de la cuestión es distinto. " + TEXTO)
    antes = radiografia(doc)
    informe = aplicar_con(doc, {})
    assert texto(p) == "La UE y España en el Siglo XXI"
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []


def test_mayusculas_en_varios_fragmentos_sin_mover_el_formato():
    doc = documento_con_portada()
    p = doc.add_heading("", level=1)
    p.add_run("MARCO ")
    p.add_run("TEÓRICO").italic = True
    doc.add_paragraph(TEXTO)
    aplicar_con(doc, {})
    assert [r.text for r in p.runs] == ["Marco ", "Teórico"] and p.runs[1].italic


def test_titulo_que_ya_cumple_no_cambia():
    doc = documento_con_portada()
    p = titulo(doc, "El Caso de la ONU", 1)
    informe = aplicar_con(doc, {})
    assert texto(p) == "El Caso de la ONU" and previstos_sin_el_titulo(informe) == []


def test_palabras_dudosas_en_los_titulos():
    doc = documento_con_portada()
    p = titulo(doc, "lo que dicen los datos", 2)
    informe = aplicar_con(doc, {})
    assert texto(p) == "Lo que Dicen los Datos"
    assert any(a.startswith("En los títulos se ha dejado como estaba «que»") for a in informe.avisos)


@pytest.mark.parametrize("original, esperado", [
    ("a) evolución demográfica", "a) Evolución Demográfica"),   # la numeración, como está (no «A)»)
    ("IV. marco teórico", "IV. Marco Teórico"),
    ("1.1. objetivos", "1.1. Objetivos"),
    ("2.3. LA UE EN CIFRAS", "2.3. La Ue en Cifras"),          # sin el trabajo, «UE» no se reconoce como sigla
])
def test_mayusculas_apa_respetan_la_numeracion(original, esperado):
    assert titulos.mayusculas_apa(original)[0] == esperado


def test_el_verificador_ve_un_titulo_sin_las_mayusculas_de_apa():
    doc = documento_con_portada()
    titulo(doc, "2. Marco teórico", 1)
    assert problemas_del_verificador(doc, MAYUSCULAS_TITULOS) == [
        "No lleva las mayúsculas de APA (sería «2. Marco Teórico»)."]
    aplicar(doc)
    assert problemas_del_verificador(doc, MAYUSCULAS_TITULOS) == []


@pytest.mark.parametrize("escrito", ["Lo que Dicen los Datos", "Lo Que Dicen los Datos"])
def test_el_verificador_admite_las_dudosas_de_las_dos_formas(escrito):
    doc = documento_con_portada()
    titulo(doc, escrito, 2)
    assert problemas_del_verificador(doc, MAYUSCULAS_TITULOS) == []


def test_aviso_de_remisiones_y_del_indice():
    doc = documento_con_portada()
    doc.add_paragraph("Índice")   # un índice escrito a mano: el clasificador lo reconoce por el rótulo
    titulo(doc, "1. Método", 1)
    doc.add_paragraph("Como se explica en el apartado 2.1, los datos se recogieron en dos fases. " + TEXTO)
    titulo(doc, "2. Resultados", 1)
    informe = aplicar_con(doc, {"quitar_numeracion": True})
    assert any(a.startswith("El texto remite a apartados por su número («apartado 2.1»)") for a in informe.avisos)
    assert any(a.startswith("El índice todavía muestra los títulos como estaban") for a in informe.avisos)


def test_la_integridad_solo_acepta_lo_previsto():
    doc = documento_con_portada()
    p = titulo(doc, "1. MÉTODO", 1)
    antes = radiografia(doc)
    numero = 2   # el párrafo del título (el 1 es la portada)
    p.runs[0].text = "Método nuevo"
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "quitar_numeracion", "1. ")])
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "mayusculas_apa", nuevo="Método nuevo")])
    p.runs[0].text = "ÉTODO"
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "quitar_numeracion", "1. M")]), \
        "solo se puede quitar una numeración, aunque lo quitado coincida"
    p.runs[0].text = "Método"
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "quitar_numeracion", "1. "),
                                              CambioPrevisto(numero, "mayusculas_apa", nuevo="Método")]) == []
    p.runs[0].text = "1. Método"
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "mayusculas_apa", nuevo="1. Método")]) == []
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "mayusculas_apa", nuevo="1. Métodos")])
    assert comparar(antes, radiografia(doc), [CambioPrevisto(numero, "mayusculas_apa", nuevo="1. MÉTODO")]), \
        "un cambio de mayúsculas anotado tiene que cambiar algo"


# --- 4.4 Título del trabajo y rótulo «Introducción» -------------------------------------------

def trabajo(portada, apartados=(("Introducción", 1), ("Método", 1))):
    """
    Un documento con portada (cada línea: texto, tamaño en puntos, negrita y, si se da,
    estilo) que acaba en un salto de página, y apartados con texto debajo.
    """
    doc = docx.Document()
    for linea in portada:
        texto_linea, tamano, negrita, *estilo = linea
        p = doc.add_paragraph(style=estilo[0] if estilo else None)
        r = p.add_run(texto_linea)
        r.font.size, r.bold = Pt(tamano), negrita
    p.runs[-1].add_break(WD_BREAK.PAGE)
    for texto_apartado, nivel in apartados:
        titulo(doc, texto_apartado, nivel)
    return doc


PORTADA = [("El sueño en la universidad", 16, True), ("Ana Ruiz", 12, False), ("Psicología", 12, False)]


def primero_del_texto(doc):
    clasificados = estructura.clasificar(doc)
    return next(c for c in clasificados if c.zona == "cuerpo" and c.tipo != "vacio")


def titulos_del_trabajo(doc):
    return [p for p in doc.element.body.iter(w("p"))
            if Estilos(doc).nombre(Estilos(doc).estilo_de(p)).lower() == "title"]


def test_el_titulo_de_la_portada_se_repite_al_principio_del_texto():
    doc = trabajo(PORTADA)
    antes = radiografia(doc)
    informe = aplicar(doc)
    [nuevo] = titulos_del_trabajo(doc)
    assert texto_parrafo(nuevo) == "El Sueño en la Universidad"   # con las mayúsculas de APA (paso 4.6)
    assert nuevo.getnext() is not None and texto_parrafo(nuevo.getnext()) == "Introducción"
    assert nuevo.find(f"{w('pPr')}/{w('outlineLvl')}").get(w("val")) == "0", "sale en el panel de navegación"
    estilo = titulos.estilos_del_titulo_del_trabajo(doc)[0]
    assert estilo.find(f"{w('pPr')}/{w('outlineLvl')}") is None, "el estilo, sin nivel: no sale en el índice"
    assert estilo.find(f"{w('pPr')}/{w('jc')}").get(w("val")) == "center"
    assert "Título del trabajo: se ha repetido «El Sueño en la Universidad» al principio del texto, centrado y " \
           "en negrita, como pide APA." in informe.cambios
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert problemas_del_verificador(doc, TITULO_TRABAJO) == []
    assert problemas_del_verificador(doc) == problemas_del_verificador(doc, FORMATO_TITULOS) == []


def test_el_titulo_con_el_estilo_title_manda_sobre_el_tamano():
    doc = trabajo([("UNIVERSIDAD DE SALAMANCA", 18, True), ("Hábitos de lectura", 12, False, "Title")])
    aplicar(doc)
    assert [texto_parrafo(p) for p in titulos_del_trabajo(doc)][-1] == "Hábitos de Lectura"


def test_con_subtitulo_va_en_dos_lineas_sin_anadir_nada():
    doc = trabajo([("EL TURISMO RURAL", 18, True), ("estudio de cuatro municipios", 16, True),
                   ("Autora: Lucía Prieto", 12, False)])
    antes = radiografia(doc)
    informe = aplicar(doc)
    [nuevo] = titulos_del_trabajo(doc)
    assert texto_parrafo(nuevo) == "El Turismo Rural\nEstudio de Cuatro Municipios"   # el subtítulo, con mayúscula
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert any(c.startswith("Título del trabajo: al principio del texto lleva mayúscula inicial en las palabras "
                            "principales, como pide APA («EL TURISMO RURAL estudio de cuatro municipios» → «El "
                            "Turismo Rural Estudio de Cuatro Municipios»)") for c in informe.cambios)


def test_titulo_del_trabajo_tal_cual():
    doc = trabajo([("EL TURISMO RURAL", 18, True), ("estudio de cuatro municipios", 16, True),
                   ("Autora: Lucía Prieto", 12, False)])
    antes = radiografia(doc)
    informe = Informe()
    titulos.aplicar(doc, {"titulo_tal_cual": True}, informe)
    [nuevo] = titulos_del_trabajo(doc)
    assert texto_parrafo(nuevo) == "EL TURISMO RURAL\nestudio de cuatro municipios"
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert any(a.startswith("El título del trabajo está escrito en mayúsculas («EL TURISMO RURAL»") and
               "--titulo-tal-cual" in a for a in informe.avisos)


def test_una_linea_de_datos_no_es_el_subtitulo():
    """«Autora: …» debajo del título, aunque vaya en negrita y grande, es un dato de la portada."""
    doc = trabajo([("El turismo rural", 18, True), ("Autora: Lucía Prieto", 14, True), ("Grado", 12, False)])
    aplicar(doc)
    assert [texto_parrafo(p) for p in titulos_del_trabajo(doc)] == ["El Turismo Rural"]


def test_el_salto_de_pagina_y_el_formato_de_la_portada_no_se_copian():
    doc = docx.Document()
    doc.add_paragraph("Ana Ruiz").runs[0].font.size = Pt(12)
    p = doc.add_paragraph()
    rojo = p.add_run("Hábitos de ")
    rojo.font.size, rojo.bold, rojo.font.color.rgb = Pt(16), True, docx.shared.RGBColor(0xFF, 0, 0)
    obra = p.add_run("Cien años de soledad ")
    obra.font.size, obra.bold, obra.italic = Pt(16), True, True
    p.add_run().add_break(WD_BREAK.PAGE)   # el salto de página, al final del título
    titulo(doc, "Método", 1)
    informe = aplicar(doc)
    [nuevo] = titulos_del_trabajo(doc)
    assert texto_parrafo(nuevo) == "Hábitos de Cien Años de Soledad"
    assert next(nuevo.iter(w("br")), None) is None, "sin el salto de página"
    assert [r.find(f"{w('rPr')}/{w('i')}") is not None for r in nuevo.iter(w("r"))] == [False, True], \
        "la cursiva (el título de una obra) se conserva"
    assert next(nuevo.iter(w("sz")), None) is None and next(nuevo.iter(w("color")), None) is None
    assert informe.cambios_previstos[-1].nuevo == "Hábitos de Cien Años de Soledad"


def test_sin_titulo_reconocible_solo_se_avisa():
    doc = trabajo([("Ana Ruiz", 12, False), ("Psicología", 12, False)])
    informe = aplicar(doc)
    assert titulos_del_trabajo(doc) == []
    assert any(a.startswith("No se ha encontrado el título del trabajo en la portada") for a in informe.avisos)


def test_si_ya_estaba_no_se_repite():
    doc = trabajo(PORTADA, apartados=())
    doc.add_paragraph("El sueño en la universidad")
    doc.add_paragraph(TEXTO)
    informe = aplicar(doc)
    assert titulos_del_trabajo(doc) == [] and previstos_sin_el_titulo(informe) == informe.cambios_previstos


def test_si_el_primer_apartado_empezaba_pagina_la_empieza_el_titulo():
    doc = trabajo([("El sueño", 16, False)])   # sin salto: el estilo del título lo pone
    unico(doc, 1).get_or_add_pPr().get_or_add_pageBreakBefore()
    aplicar(doc)
    [nuevo] = titulos_del_trabajo(doc)
    assert nuevo.find(f"{w('pPr')}/{w('pageBreakBefore')}") is not None
    assert nuevo.getnext().find(f"{w('pPr')}/{w('pageBreakBefore')}").get(w("val")) == "0"


def test_la_integridad_solo_acepta_una_copia_de_la_portada():
    doc = trabajo(PORTADA)
    antes = radiografia(doc)
    aplicar(doc)
    # Portada: párrafos 1 a 3; «Introducción», el 4.
    # La copia puede cambiar las mayúsculas (paso 4.6), y nada más.
    previsto = CambioPrevisto(4, "añadir_titulo", nuevo="El Sueño en la Universidad", desde=(1,))
    assert comparar(antes, radiografia(doc), [previsto]) == []
    for falso in (CambioPrevisto(4, "añadir_titulo", nuevo="El Sueño en la Universidad", desde=()),
                  CambioPrevisto(4, "añadir_titulo", nuevo="El Sueño en la Universidad", desde=(2,)),
                  CambioPrevisto(4, "añadir_titulo", nuevo="El Sueño en la Universidad", desde=(1, 3)),
                  CambioPrevisto(4, "añadir_titulo", nuevo="El Sueño en la Universidad", desde=(4,)),
                  CambioPrevisto(4, "añadir_titulo", nuevo="El Sueño en la Universidad.", desde=(1,)),
                  CambioPrevisto(4, "añadir_titulo", nuevo="Un título inventado", desde=(1,))):
        assert comparar(antes, radiografia(doc), [falso]), falso
    assert comparar(antes, radiografia(doc)), "sin anotarlo, la integridad lo ve"


def test_el_verificador_ve_que_falta_el_titulo():
    doc = trabajo(PORTADA)
    aplicar(doc)
    [nuevo] = titulos_del_trabajo(doc)
    nuevo.getparent().remove(nuevo)
    assert problemas_del_verificador(doc, TITULO_TRABAJO) == [
        "El texto no empieza con el título del trabajo (el de la portada, con el estilo «Título» de Word)."]


def test_introduccion_por_defecto_solo_se_avisa():
    doc = trabajo(PORTADA)
    informe = aplicar(doc)
    assert any(texto_parrafo(p) == "Introducción" for p in doc.element.body.iter(w("p")))
    assert [a for a in informe.avisos if "Introducción" in a] == [
        "El primer apartado se titula «Introducción». En APA la introducción no lleva título: el título del "
        "trabajo, al principio del texto, hace sus veces. Se ha dejado así; quítalo tú o usa la opción "
        "--quitar-introduccion."]
    assert [h.regla for h in verificar(doc).excepciones if h.regla != PORTADA_PROPIA] == [TITULO_TRABAJO]


def test_introduccion_con_la_opcion_se_quita():
    doc = trabajo(PORTADA, apartados=(("1. Introducción", 1), ("2. Método", 1)))
    rotulo = next(p for p in doc.element.body.iter(w("p")) if texto_parrafo(p) == "1. Introducción")
    marcador = rotulo.makeelement(w("bookmarkStart"), {w("id"): "0", w("name"): "_Toc1"})
    rotulo.append(marcador)
    antes = radiografia(doc)
    informe = aplicar_con(doc, {"quitar_introduccion": True})
    assert rotulo.getparent() is None
    assert marcador.getparent() is not None, "el marcador pasa al párrafo siguiente"
    assert "Títulos: se ha quitado el rótulo «1. Introducción» del primer apartado: en APA la introducción no " \
           "lleva título." in informe.cambios
    [nuevo] = titulos_del_trabajo(doc)
    assert texto_parrafo(nuevo.getnext()) == TEXTO, "el título del trabajo ocupa su lugar"
    assert comparar(antes, radiografia(doc), informe.cambios_previstos) == []
    assert [h for h in verificar(doc).excepciones if h.regla == TITULO_TRABAJO] == []


def test_introduccion_con_subapartados_no_se_quita():
    doc = trabajo(PORTADA, apartados=(("Introducción", 1), ("Objetivos", 2), ("Hipótesis", 2), ("Método", 1)))
    informe = aplicar_con(doc, {"quitar_introduccion": True})
    assert any(texto_parrafo(p) == "Introducción" for p in doc.element.body.iter(w("p")))
    assert any(a.startswith("No se ha quitado el rótulo «Introducción» porque el apartado tiene subapartados "
                            "(«Objetivos», «Hipótesis»)") for a in informe.avisos)


def test_introduccion_que_no_es_el_primer_apartado_no_se_toca():
    doc = trabajo(PORTADA, apartados=(("Marco teórico", 1), ("Introducción", 1)))
    informe = aplicar_con(doc, {"quitar_introduccion": True})
    assert any(texto_parrafo(p) == "Introducción" for p in doc.element.body.iter(w("p")))
    assert not [m for m in informe.avisos + informe.cambios if "Introducción" in m]


def test_la_integridad_solo_borra_el_rotulo_introduccion():
    doc = trabajo(PORTADA)
    antes = radiografia(doc)
    metodo = next(p for p in doc.element.body.iter(w("p")) if texto_parrafo(p) == "Método")
    metodo.getparent().remove(metodo)
    # Portada: párrafos 1 a 3; «Introducción», el 4; «Método», el 6.
    assert comparar(antes, radiografia(doc), [CambioPrevisto(6, "quitar_introduccion")]), "«Método» no"
    assert comparar(antes, radiografia(doc), [CambioPrevisto(4, "quitar_introduccion")]), "«Introducción» sí, " \
        "pero no es lo que se ha borrado"


def test_cinco_documentos_con_quitar_introduccion(tmp_path):
    """01, 02 y 03 la pierden; el 05 no (tiene subapartados), y el 04 no la tiene."""
    quitadas = {}
    for original in PRUEBAS:
        salida, informe = pipeline.procesar(original, tmp_path, {"quitar_introduccion": True})
        assert verificar(salida).cumple
        quitadas[original.name[:2]] = sum(c.accion == "quitar_introduccion" for c in informe.cambios_previstos)
    assert quitadas == {"01": 1, "02": 1, "03": 1, "04": 0, "05": 0}


# --- Avisos de estructura ---------------------------------------------------------------------

def test_aviso_de_un_solo_subapartado():
    doc = documento_con_portada()
    for texto, nivel in (("Método", 1), ("Participantes", 2), ("Resultados", 1), ("Datos", 2), ("Análisis", 2)):
        doc.add_heading(texto, level=nivel)
        doc.add_paragraph(TEXTO)
    informe = aplicar(doc)
    assert sin_el_titulo(informe.avisos) == ["El apartado «Método» tiene un solo subapartado («Participantes»). APA recomienda "
                              "dividir un apartado solo si tiene al menos dos subapartados: añade otro o integra "
                              "su contenido en el apartado."]


def test_los_nietos_no_cuentan_como_subapartados():
    doc = documento_con_portada()
    for texto, nivel in (("Método", 1), ("Muestra", 2), ("Edad", 3), ("Sexo", 3), ("Instrumentos", 2)):
        doc.add_heading(texto, level=nivel)
        doc.add_paragraph(TEXTO)
    assert aplicar(doc).avisos == []


def test_aviso_de_salto_de_nivel():
    doc = documento_con_portada()
    for texto, nivel in (("Método", 1), ("Edad", 3), ("Sexo", 3)):
        doc.add_heading(texto, level=nivel)
        doc.add_paragraph(TEXTO)
    avisos = aplicar(doc).avisos
    assert avisos == ["El título «Edad» es de nivel 3, pero el anterior («Método») es de nivel 1: se salta un "
                      "nivel. APA pide no saltarse niveles: revisa si debería ser de nivel 2."]


def test_aviso_si_el_primer_titulo_no_es_de_nivel_1():
    doc = documento_con_portada()
    for texto in ("Muestra", "Instrumentos"):
        doc.add_heading(texto, level=2)
        doc.add_paragraph(TEXTO)
    assert aplicar(doc).avisos == ["El primer título del trabajo, «Muestra», es de nivel 2. APA pide empezar por "
                                   "el nivel 1: revisa su nivel."]


# --- Documentos de prueba ---------------------------------------------------------------------

UN_SOLO_SUBAPARTADO = {   # apartado → su único subapartado
    "01": [],
    "02": [],
    "03": [("2. Datos y resultados", "2.1. Fuentes de energía")],
    "04": [("Efectos sobre la autoestima", "Factores de riesgo:")],
    "05": [("1. Introducción", "1.1. Objetivos")],
}
SIN_ESTILO_EN_EL_ORIGINAL = {"01": 5, "02": 0, "03": 0, "04": 4, "05": 1}   # títulos falsos (05: «RESUMEN»)


@pytest.mark.parametrize("original", PRUEBAS, ids=lambda p: p.stem)
def test_documentos_de_prueba(original, tmp_path):
    clave = original.name[:2]
    sin_estilo = [m for m in problemas_del_verificador(docx.Document(str(original))) if "sin el estilo" in m]
    assert len(sin_estilo) == SIN_ESTILO_EN_EL_ORIGINAL[clave]

    salida, informe = pipeline.procesar(original, tmp_path)
    copia = docx.Document(str(salida))
    assert problemas_del_verificador(copia) == problemas_del_verificador(copia, FORMATO_TITULOS) == []
    assert formato_corregido(informe) == [], "en los 5 documentos no queda formato a mano que contradiga"
    en_linea = [c for c in informe.cambios if "en la misma línea que su texto" in c]
    assert len(en_linea) == (1 if clave == "05" else 0)   # «a) Evolución demográfica» y «b) Percepción…»
    assert problemas_del_verificador(copia, TITULOS_EN_LINEA) == []
    for nivel in titulos.NIVELES:
        ids = [e.get(w("styleId")) for e in titulos.estilos_de_titulo(copia)[nivel]]
        assert len(ids) == len(set(ids)), f"«Título {nivel}» sin duplicados"

    estilos = Estilos(copia)
    por_id = {e.get(w("styleId")): e for e in copia.styles.element.iter(w("style"))}
    for c in estructura.clasificar(copia):
        if c.tipo == "titulo":
            nombre = por_id[estilos.estilo_de(c.elemento)].find(w("name")).get(w("val")).lower()
            assert nombre == f"heading {c.nivel}" or (nombre == "title" and c.nivel == 1), c.texto

    esperados = [f"El apartado «{a}» tiene un solo subapartado («{b}»)" for a, b in UN_SOLO_SUBAPARTADO[clave]]
    de_estructura = [a for a in informe.avisos if "subapartado" in a or "se salta un nivel" in a]
    assert [a.split(". APA")[0] for a in de_estructura] == esperados
