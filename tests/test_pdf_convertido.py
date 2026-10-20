"""
Revisión del propietario (2026-10-20): un trabajo que era un PDF convertido a Word salía
con medias páginas en blanco, casi todo como citas en bloque, la portada cambiada, el
índice de tablas roto y el número de página pegado al borde. Estas pruebas usan un
documento que imita lo que hace el conversor (el trabajo real no se sube al repositorio).
"""

import docx
import pytest
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Pt, Twips

from apa_facil import documento, pipeline
from apa_facil.documento import CambioPrevisto, comparar, radiografia, texto_parrafo, w
from apa_facil.estructura import clasificar
from apa_facil.reglas import portada as regla_portada

LARGO = ("Este párrafo cuenta con detalle cómo se hizo el estudio con la alumna y qué pruebas se le pasaron "
         "en el centro, para que se entienda bien el resto del trabajo y sus conclusiones.")
CORTADO = "Una vez hecho el diagnóstico, se entrega el informe a la familia. La psicopedagoga"
SIGUE = "insiste en que su hija necesita una reeducación especializada con el apoyo de la familia y del centro."
CORTADO_EN_EL = "Aprovechando que la niña se encuentra en el"
SIGUE_EN_MAYUSCULA = "S. XXI, se va a ampliar el material a otras herramientas y programas de apoyo pedagógico."
SANGRIA_DEL_PDF = 501


def _seccion(doc, nueva=WD_SECTION.NEW_PAGE):
    seccion = doc.add_section(nueva)
    seccion.header_distance = Twips(0)
    return seccion


def _texto(doc, texto, sangria=SANGRIA_DEL_PDF):
    p = doc.add_paragraph()
    r = p.add_run(texto)
    r._r.get_or_add_rPr().append(_espaciado(24))   # letras separadas para imitar el justificado
    p.paragraph_format.left_indent = Twips(sangria)
    return p


def _espaciado(valor):
    el = documento.OxmlElement("w:spacing")
    el.set(w("val"), str(valor))
    return el


def convertido():
    """Portada en una tabla, índice de tablas, y una sección por «página» con párrafos partidos."""
    doc = docx.Document()
    portada = doc.add_table(rows=2, cols=1)
    portada.cell(0, 0).paragraphs[0].add_run("Universidad de Ejemplo").font.size = Pt(14)
    titulo = portada.cell(1, 0).paragraphs[0]
    titulo.add_run("La dislexia: un caso real").font.size = Pt(40)
    titulo.paragraph_format.space_before = Pt(24)
    doc.add_paragraph()
    _seccion(doc)
    doc.add_paragraph("ÍNDICE DE TABLAS Y FIGURAS")
    doc.add_paragraph("TABLAS", style="Heading 3")
    for n in (1, 2):
        linea = doc.add_paragraph(f"Tabla {n}. Resultados de las pruebas del informe psicopedagógico\t{20 + n}")
        linea.paragraph_format.tab_stops.add_tab_stop(Twips(8318), leader=4)   # guiones
        linea.paragraph_format.right_indent = Twips(634)
    _seccion(doc)
    doc.add_heading("Introducción", level=1)
    for pagina in range(12):
        _texto(doc, LARGO)
        if pagina == 3:
            _texto(doc, CORTADO)
        elif pagina == 6:
            _texto(doc, CORTADO_EN_EL)
        else:
            _texto(doc, LARGO)
        _seccion(doc)
        if pagina == 3:
            _texto(doc, SIGUE)
        elif pagina == 6:
            _texto(doc, SIGUE_EN_MAYUSCULA)
        if pagina == 8:
            doc.add_heading("Conclusiones", level=1)
    _texto(doc, LARGO)
    return doc


@pytest.fixture(scope="module")
def procesado(tmp_path_factory):
    carpeta = tmp_path_factory.mktemp("pdf")
    ruta = carpeta / "convertido.docx"
    convertido().save(str(ruta))
    salida, informe = pipeline.procesar(ruta, carpeta / "salida")
    return docx.Document(str(ruta)), docx.Document(str(salida)), informe


def _textos(doc):
    return [texto_parrafo(p) for p in doc.element.body.iter(w("p"))]


def test_se_reconoce_como_pdf_convertido():
    clasificados = clasificar(convertido())
    assert clasificados.convertido_de_pdf
    assert not any(c.tipo == "cita_bloque" for c in clasificados), "la sangría del PDF no es una cita en bloque"


def test_los_saltos_a_mitad_del_texto_pasan_a_continuos(procesado):
    _, copia, informe = procesado
    tipos = []
    for p in copia.element.body.iter(w("p")):
        seccion = p.find(f"{w('pPr')}/{w('sectPr')}")
        if seccion is not None:
            tipo = seccion.find(w("type"))
            siguiente = p.getnext()
            tipos.append((None if tipo is None else tipo.get(w("val")),
                          texto_parrafo(siguiente) if siguiente is not None else ""))
    # Se mantienen: tras la portada, tras el índice y antes de «Conclusiones» (título de nivel 1).
    empiezan_pagina = [siguiente for tipo, siguiente in tipos if tipo != "continuous"]
    assert empiezan_pagina == ["ÍNDICE DE TABLAS Y FIGURAS", "Introducción", "Conclusiones"]
    assert any(c.startswith("Saltos de sección: 11 saltos") for c in informe.cambios)
    assert len(copia.sections) == len(procesado[0].sections), "las secciones siguen ahí"


def test_se_unen_los_parrafos_partidos(procesado):
    _, copia, informe = procesado
    textos = _textos(copia)
    assert f"{CORTADO} {SIGUE}" in textos
    assert f"{CORTADO_EN_EL} {SIGUE_EN_MAYUSCULA}" in textos, "«en el» nunca acaba un párrafo"
    assert SIGUE not in textos and SIGUE_EN_MAYUSCULA not in textos
    assert any(c.startswith("Párrafos partidos: se han unido 2 párrafos") for c in informe.cambios)


def test_el_texto_va_pegado_al_margen(procesado):
    _, copia, _ = procesado
    p = next(p for p in copia.paragraphs if p.text == LARGO)
    assert p.paragraph_format.left_indent in (None, 0)
    assert p.paragraph_format.first_line_indent == Twips(720)


def test_se_quita_el_espaciado_entre_letras(procesado):
    _, copia, informe = procesado
    p = next(p for p in copia.paragraphs if p.text == LARGO)
    assert all(r._r.find(f"{w('rPr')}/{w('spacing')}") is None for r in p.runs)
    assert any(c.startswith("Espaciado entre letras") for c in informe.cambios)
    assert any("parece un PDF convertido" in a for a in informe.avisos)


def test_la_portada_en_tabla_no_se_toca(procesado):
    original, copia, informe = procesado
    titulo = next(p for p in copia.element.body.iter(w("p")) if texto_parrafo(p) == "La dislexia: un caso real")
    r = titulo.find(w("r"))
    assert documento.Estilos(copia).tamano(r, titulo) == 40
    assert documento.Estilos(copia).parrafo(titulo, "spacing", "before") == "480"
    assert copia.sections[0].left_margin == original.sections[0].left_margin, "la sección de la portada"
    assert copia.sections[1].left_margin == Twips(1440)
    assert any(a.startswith("Se ha mantenido tu portada («Universidad de Ejemplo»)") for a in informe.avisos)


def test_portada_con_todo_su_texto_en_tablas():
    clasificados = clasificar(convertido())
    assert regla_portada.portada_existente(clasificados)
    textos = {texto_parrafo(p) for p in regla_portada.parrafos_de(clasificados)}
    assert {"Universidad de Ejemplo", "La dislexia: un caso real"} <= textos


def test_indice_de_tablas(procesado):
    _, copia, informe = procesado
    clasificados = clasificar(convertido())
    tipos = {c.texto: c.tipo for c in clasificados if c.texto}
    assert tipos["ÍNDICE DE TABLAS Y FIGURAS"] == tipos["TABLAS"] == "indice"
    assert all(tipo == "indice" for texto, tipo in tipos.items() if texto.startswith("Tabla "))
    linea = next(p for p in copia.paragraphs if p.text.startswith("Tabla 1."))
    paradas = [(t.alignment, t.leader, t.position) for t in linea.paragraph_format.tab_stops]
    assert len(paradas) == 1 and paradas[0][2] == Twips(copia.sections[1].page_width.twips - 2880)
    assert linea.paragraph_format.right_indent in (None, 0)
    assert any(c.startswith("Índice: 2 líneas del índice escrito a mano") for c in informe.cambios)


def test_encabezado_separado_del_borde(procesado):
    _, copia, informe = procesado
    assert all(s.header_distance == Twips(708) for s in copia.sections[1:]), "estaban a 0"
    assert any("pegado al borde" in c for c in informe.cambios)


def test_seccion_horizontal_no_pasa_a_continua(tmp_path):
    doc = docx.Document()
    doc.add_paragraph(LARGO)
    doc.add_paragraph(LARGO)
    horizontal = doc.add_section(WD_SECTION.NEW_PAGE)
    horizontal.orientation = WD_ORIENT.LANDSCAPE
    horizontal.page_width, horizontal.page_height = horizontal.page_height, horizontal.page_width
    doc.add_paragraph(LARGO)
    ruta = tmp_path / "horizontal.docx"
    doc.save(str(ruta))
    salida, _ = pipeline.procesar(ruta, tmp_path / "salida")
    seccion = next(docx.Document(str(salida)).element.body.iter(w("sectPr")))
    assert seccion.find(w("type")) is None or seccion.find(w("type")).get(w("val")) != "continuous"


# --- La integridad: «unir_parrafos» -------------------------------------------------------

def _dos_partes(antes, despues):
    doc = docx.Document()
    doc.add_paragraph(antes)
    doc.add_paragraph("")
    doc.add_paragraph(despues)
    return doc


@pytest.mark.parametrize("antes, despues, valido", [
    ("la psicopedagoga", "insiste en que", True),
    ("se encuentra en el", "S. XXI", True),
    ("Fin del párrafo.", "otro párrafo", False),   # el primero ya estaba acabado
    ("la psicopedagoga", "Insiste en que", False),   # mayúscula sin palabra que lo pida
    ("los resultados son:", "primero, la lectura", False),
])
def test_integridad_de_unir_parrafos(antes, despues, valido):
    original = radiografia(_dos_partes(antes, despues))
    unido = docx.Document()
    unido.add_paragraph(f"{antes} {despues}")
    unido.add_paragraph("")
    diferencias = comparar(original, radiografia(unido), [CambioPrevisto(1, "unir_parrafos", desde=(3,))])
    assert (diferencias == []) == valido


def test_unir_parrafos_no_admite_texto_en_medio():
    doc = docx.Document()
    for texto in ("la psicopedagoga", "otro texto", "insiste en que"):
        doc.add_paragraph(texto)
    unido = docx.Document()
    unido.add_paragraph("la psicopedagoga insiste en que")
    unido.add_paragraph("otro texto")
    assert comparar(radiografia(doc), radiografia(unido), [CambioPrevisto(1, "unir_parrafos", desde=(3,))])
