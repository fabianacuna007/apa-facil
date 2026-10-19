"""
Paso 7.1: detector de citas en el texto (apa_facil/revision/citas.py), según la
web oficial de APA (Author–date citation system, Parenthetical versus narrative
in-text citations, Citing multiple works, Quotations y Direct quotation of
material with page numbers). Los ejemplos de APA son los de esas páginas (no los
de normas-apa.org: ver la fase 7 del ROADMAP), y sus versiones en español.

El detector solo lee y no juzga: guarda lo escrito tal cual, también las formas
mal escritas («(Martínez 2019)», «et. al», «s.f», «p.45»), para el paso 7.4.
"""

import hashlib
from pathlib import Path

import docx
import pytest
from lxml import etree

from apa_facil import documento
from apa_facil.__main__ import main
from apa_facil.revision import citas
from apa_facil.revision.citas import ENTRE_PARENTESIS as P, NARRATIVA as N, SIN_PARENTESIS as S

RAIZ = Path(__file__).resolve().parent.parent
PRUEBAS = {ruta.name[:2]: ruta for ruta in sorted((RAIZ / "pruebas").glob("*.docx"))}


def resumen(cita):
    """Lo principal de una cita, para compararlo en las pruebas."""
    return (cita.tipo, cita.autores, cita.año, cita.localizador)


def en(texto):
    return [resumen(c) for c in citas.citas_del_texto(texto)]


# --- Ejemplos de la web oficial de APA ------------------------------------------------

@pytest.mark.parametrize("texto, esperado", [
    # Author–date citation system: la tabla de formatos básicos.
    ("(Luna, 2020)", [(P, ["Luna"], "2020", None)]),
    ("Luna (2020)", [(N, ["Luna"], "2020", None)]),
    ("(Salas & D’Agostino, 2020)", [(P, ["Salas", "D’Agostino"], "2020", None)]),
    ("Salas and D’Agostino (2020)", [(N, ["Salas", "D’Agostino"], "2020", None)]),
    ("(Martin et al., 2020)", [(P, ["Martin"], "2020", None)]),
    ("Martin et al. (2020)", [(N, ["Martin"], "2020", None)]),
    ("(National Institute of Mental Health [NIMH], 2020)", [(P, ["National Institute of Mental Health"], "2020", None)]),
    ("(NIMH, 2020)", [(P, ["NIMH"], "2020", None)]),
    ("National Institute of Mental Health (NIMH, 2020)", [(N, ["National Institute of Mental Health"], "2020", None)]),
    ("NIMH (2020)", [(N, ["NIMH"], "2020", None)]),
    ("(Stanford University, 2020)", [(P, ["Stanford University"], "2020", None)]),
    ("Stanford University (2020)", [(N, ["Stanford University"], "2020", None)]),
    # Parenthetical versus narrative.
    ("Falsely balanced news coverage can distort the public’s perception (Koehler, 2016).",
     [(P, ["Koehler"], "2016", None)]),
    ("(see Koehler, 2016, for more detail)", [(P, ["Koehler"], "2016", None)]),
    ("(e.g., falsely balanced news coverage; Koehler, 2016)", [(P, ["Koehler"], "2016", None)]),
    ("Koehler (2016) noted the dangers of falsely balanced news coverage.", [(N, ["Koehler"], "2016", None)]),
    # Citing multiple works.
    ("(Adams et al., 2019; Shumway & Shulman, 2015; Westinghouse, 2017)",
     [(P, ["Adams"], "2019", None), (P, ["Shumway", "Shulman"], "2015", None), (P, ["Westinghouse"], "2017", None)]),
    ("(Department of Veterans Affairs, n.d., 2017a, 2017b, 2019)",
     [(P, ["Department of Veterans Affairs"], año, None) for año in ("n.d.", "2017a", "2017b", "2019")]),
    ("Zhou (n.d., 2000, 2016, in press)", [(N, ["Zhou"], año, None) for año in ("n.d.", "2000", "2016", "in press")]),
    ("(Sampson & Hughes, 2020; see also Augustine, 2017; Melara et al., 2018; Pérez, 2014)",
     [(P, ["Sampson", "Hughes"], "2020", None), (P, ["Augustine"], "2017", None),
      (P, ["Melara"], "2018", None), (P, ["Pérez"], "2014", None)]),
    ("Suliman (2018), Gutiérrez (2012, 2017), and Medina and Reyes (2019) examined",
     [(N, ["Suliman"], "2018", None), (N, ["Gutiérrez"], "2012", None), (N, ["Gutiérrez"], "2017", None),
      (N, ["Medina", "Reyes"], "2019", None)]),
    # Quotations y page numbers.
    ("“high performance along one domain does not translate” (Ervin et al., 2018, p. 470).",
     [(P, ["Ervin"], "2018", "p. 470")]),
    ("of inner speech and its cognitive and neural underpinnings. (Alderson-Day & Fernyhough, 2015, p. 957)",
     [(P, ["Alderson-Day", "Fernyhough"], "2015", "p. 957")]),
    ("(Pérez, 2019, pp. 34–36)", [(P, ["Pérez"], "2019", "pp. 34–36")]),
    ("(Pérez, 2019, pp. 67, 72)", [(P, ["Pérez"], "2019", "pp. 67, 72")]),
    ("(Pérez, 2019, p. S41)", [(P, ["Pérez"], "2019", "p. S41")]),
    ("Webster-Stratton (2016, pp. 152–153) described", [(N, ["Webster-Stratton"], "2016", "pp. 152–153")]),
])
def test_ejemplos_de_la_web_de_apa(texto, esperado):
    assert en(texto) == esperado


@pytest.mark.parametrize("texto, esperado", [
    ("(Salas y D’Agostino, 2020)", [(P, ["Salas", "D’Agostino"], "2020", None)]),
    ("Salas y D’Agostino (2020)", [(N, ["Salas", "D’Agostino"], "2020", None)]),
    ("(Gómez, s. f.)", [(P, ["Gómez"], "s. f.", None)]),
    ("Gómez (s. f.)", [(N, ["Gómez"], "s. f.", None)]),
    ("(Zhou, s. f., 2000, en prensa)", [(P, ["Zhou"], año, None) for año in ("s. f.", "2000", "en prensa")]),
    ("(véase también Augustine, 2017)", [(P, ["Augustine"], "2017", None)]),
    ("(Organización Mundial de la Salud [OMS], 2020)", [(P, ["Organización Mundial de la Salud"], "2020", None)]),
    ("Organización Mundial de la Salud (OMS, 2020)", [(N, ["Organización Mundial de la Salud"], "2020", None)]),
    ("(“Título del informe”, 2019)", [(P, ["“Título del informe”"], "2019", None)]),
    ("Pérez, Gómez y Ruiz (2018)", [(N, ["Pérez", "Gómez", "Ruiz"], "2018", None)]),
])
def test_en_español(texto, esperado):
    assert en(texto) == esperado


def test_los_detalles_de_una_cita():
    texto = "Como se ha visto (Castro et al., 2019; National Institute of Mental Health [NIMH], 2020, p. 3)."
    castro, nimh = citas.citas_del_texto(texto, "Párrafo 4", 4)
    assert castro.fragmento == "(Castro et al., 2019; National Institute of Mental Health [NIMH], 2020, p. 3)"
    assert texto[castro.inicio:castro.fin] == castro.fragmento
    assert (castro.et_al, castro.et_al_escrito, castro.autores_escritos) == (True, "et al.", "Castro et al.")
    assert castro.grupo == nimh.grupo and castro.obras_en_el_grupo == 2
    assert (nimh.abreviatura, nimh.localizador, nimh.coma_antes_del_año) == ("NIMH", "p. 3", True)
    assert (castro.ubicacion, castro.numero) == ("Párrafo 4", 4)


def test_el_conector_y_el_prefijo_se_guardan():
    (amp,) = citas.citas_del_texto("(Campos & Vidal, 2022)")
    (y,) = citas.citas_del_texto("Campos y Vidal (2022)")
    (vease,) = citas.citas_del_texto("(véase Campos y Vidal, 2022)")
    assert (amp.conector, y.conector, vease.conector) == ("&", "y", "y")
    assert (amp.prefijo, vease.prefijo) == (None, "véase")


# --- Formas mal escritas: se reconocen y se guardan tal cual (para el 7.4) ------------

def test_formas_mal_escritas():
    (sin_coma,) = citas.citas_del_texto("(Martínez 2019)")
    assert (sin_coma.autores, sin_coma.año, sin_coma.coma_antes_del_año) == (["Martínez"], "2019", False)
    (et_al,) = citas.citas_del_texto("(Pérez et. al, 2020)")
    assert (et_al.et_al, et_al.et_al_escrito, et_al.autores) == (True, "et. al", ["Pérez"])
    (sf,) = citas.citas_del_texto("Gómez (s.f)")
    assert (sf.tipo, sf.año) == (N, "s.f")
    (pagina,) = citas.citas_del_texto("(Consejo Económico de Ejemplo, 2022, p.45)")
    assert pagina.localizador == "p.45"


# --- Lo que no es una cita ------------------------------------------------------------

@pytest.mark.parametrize("texto", [
    "Participaron 180 estudiantes (112 mujeres y 68 hombres) con edades de 18 a 24 años.",
    "el Índice de Calidad de Sueño de Pittsburgh (PSQI)",
    "Estadísticos descriptivos (N = 180)",
    "Consumo eléctrico medio mensual (kWh)",
    "fuentes de energía de los hogares (elaboración propia)",
    "Evolución de la población (2001-2024)",
    "según la Ley 39/2015 (BOE núm. 236)",
    "In 2016, Koehler noted the dangers of falsely balanced news coverage.",
    "el informe del Consejo Económico de Ejemplo advierte:",
    "Flores et al. described: (p. 311)",
    "lo que ocurrió en (2019)",
    "Entre 2001 y 2024 la población bajó",
])
def test_no_son_citas(texto):
    assert en(texto) == []


def test_el_nombre_de_una_cita_narrativa():
    """Se para en la primera palabra en minúscula o en un signo, y quita las palabras que empiezan frase."""
    assert [c.autores_escritos for c in citas.citas_del_texto("Según Ortega y Blanco (2018), el control")] \
        == ["Ortega y Blanco"]
    assert [c.autores_escritos for c in citas.citas_del_texto(
        "lo que confirma las conclusiones de Molina (2017) y Herranz (2015).")] == ["Molina", "Herranz"]
    assert [c.autores_escritos for c in citas.citas_del_texto("Según la revisión de Castro et al. (2019), entre")] \
        == ["Castro et al."]
    assert [c.autores_escritos for c in citas.citas_del_texto("ciudades. Molina (2017) sostiene")] == ["Molina"]
    # La coma solo se cruza en «Pérez, Gómez y Ruiz» (con un «y» detrás).
    assert [c.autores_escritos for c in citas.citas_del_texto("Por su parte, Herranz (2015) señala")] == ["Herranz"]
    assert [c.autores_escritos for c in citas.citas_del_texto("En España, Pérez (2019) señala")] == ["Pérez"]


# --- Los cinco documentos de prueba ---------------------------------------------------

ESPERADAS = {  # (párrafo, tipo, autores como están escritos, año, páginas)
    "01": [(7, P, "Ramírez Soto", "2015", None), (8, N, "Lefèvre", "2009", None),
           (10, N, "Ortega y Blanco", "2018", None), (11, P, "Castell", "2011", None),
           (13, P, "Ramírez Soto", "2015", None), (14, P, "Lefèvre", "2009", "p. 112"),
           (16, P, "Ortega y Blanco", "2018", None)],
    "02": [(6, P, "Navarro y Gil", "2020", None), (7, N, "Castro et al.", "2019", None),
           (65, P, "Navarro y Gil", "2020", None), (65, P, "Castro et al.", "2019", None)],
    "03": [(4, P, "Herrero", "2021", None), (10, N, "Herrero", "2021", None),
           (10, N, "Campos y Vidal", "2022", None)],
    "04": [(7, P, "Martínez", "2019", None), (11, P, "Pérez et. al", "2020", None),
           (12, N, "Gómez", "s.f", None),
           # La columna «Fuente» de la tabla (sin paréntesis hace falta la coma: «Martínez 2019» no cuenta).
           (44, S, "Pérez et. al", "2020", None), (51, N, "INE", "s.f", None), (58, S, "Gómez", "s.f.", None)],
    "05": [(20, P, "Collantes y Pinilla", "2011", None), (21, P, "Molina", "2017", None),
           (21, P, "Sáenz Gil", "2020", None), (31, N, "Molina", "2017", None),
           (32, P, "Consejo Económico de Ejemplo", "2022", "p.45"), (33, N, "Herranz", "2015", None),
           (63, P, "Ruiz Montero", "2023", None), (65, N, "Molina", "2017", None),
           (65, N, "Herranz", "2015", None)],
}


@pytest.mark.parametrize("clave", sorted(ESPERADAS))
def test_cinco_documentos(clave):
    encontradas = citas.detectar(documento.abrir(PRUEBAS[clave]))
    assert [(c.numero, c.tipo, c.autores_escritos, c.año, c.localizador) for c in encontradas] == ESPERADAS[clave]


def test_las_citas_de_las_tablas_dicen_que_estan_en_una_tabla():
    encontradas = citas.detectar(documento.abrir(PRUEBAS["04"]))
    assert [c.ubicacion for c in encontradas if c.numero > 40] == ["Párrafo 44 (dentro de una tabla)",
                                                                   "Párrafo 51 (dentro de una tabla)",
                                                                   "Párrafo 58 (dentro de una tabla)"]


def test_la_posicion_corresponde_al_texto_del_parrafo():
    doc = documento.abrir(PRUEBAS["05"])
    parrafos = list(doc.element.body.iter(documento.w("p")))
    for c in citas.detectar(doc):
        assert documento.texto_parrafo(parrafos[c.numero - 1])[c.inicio:c.fin] == c.fragmento


def test_no_busca_en_la_portada_ni_en_las_referencias():
    """Las referencias del 01 tienen «Ramírez Soto, M. (2015)»; la portada, «Curso 2025-2026»."""
    zonas = {c.zona for clave in PRUEBAS for c in citas.detectar(documento.abrir(PRUEBAS[clave]))}
    assert zonas <= {"cuerpo", "resumen", "anexos"}


def test_busca_en_las_notas_al_pie():
    doc = documento.abrir(PRUEBAS["05"])
    assert not [c for c in citas.detectar(doc) if c.zona == "notas"]   # sus tres notas no citan nada

    def citar_en_la_primera_nota(nombre, raiz):   # python-docx no carga las notas como XML
        if nombre == "/word/footnotes.xml":
            nota = next(n for n in raiz.iter(documento.w("footnote")) if n.get(documento.w("type")) is None)
            next(nota.iter(documento.w("t"))).text = "Así lo define Molina (2017, p. 12)."
    documento.editar_partes_de_texto(doc, citar_en_la_primera_nota)
    (cita,) = [c for c in citas.detectar(doc) if c.zona == "notas"]
    assert (cita.ubicacion, cita.numero, resumen(cita)) == ("Nota al pie 1", None, (N, ["Molina"], "2017", "p. 12"))


def test_no_cambia_el_documento():
    doc = documento.abrir(PRUEBAS["05"])
    antes = etree.tostring(doc.element)
    huella = hashlib.sha256(PRUEBAS["05"].read_bytes()).hexdigest()
    citas.detectar(doc)
    assert etree.tostring(doc.element) == antes
    assert hashlib.sha256(PRUEBAS["05"].read_bytes()).hexdigest() == huella


def test_un_documento_sin_citas():
    doc = docx.Document()
    doc.add_paragraph("Un párrafo sin ninguna cita (y con un paréntesis).")
    assert citas.detectar(doc) == []
    assert citas.como_texto([]) == "No se ha encontrado ninguna cita en el texto."


# --- Opción --citas -------------------------------------------------------------------

def test_opcion_citas_no_genera_copias(tmp_path, capsys):
    assert main([str(PRUEBAS["05"]), "--citas"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert list(tmp_path.iterdir()) == []
    assert "Citas en el texto de 05_tfg_completo.docx" in salida
    assert "Consejo Económico de Ejemplo" in salida and "p.45" in salida
    assert "Resumen: 9 citas (5 entre paréntesis y 4 narrativas) de 6 obras distintas." in salida
    assert "1 de 1 documento(s) revisado(s). No se ha generado ninguna copia." in salida


def test_opcion_citas_con_comodin_y_otra_opcion(capsys):
    assert main([str(RAIZ / "pruebas" / "*.docx"), "--citas", "--fuente", "Arial"]) == 0
    salida = capsys.readouterr().out
    assert "la opción --fuente no se usa en la revisión de citas" in salida
    assert "5 de 5 documento(s) revisado(s)" in salida
    assert "6 citas (2 entre paréntesis, 2 narrativas y 2 sin paréntesis en tablas) de 5 obras distintas" in salida


def test_opcion_citas_con_un_archivo_que_no_existe(capsys):
    assert main(["no_existe.docx", "--citas"]) == 1


def test_la_ayuda_explica_la_opcion_citas(capsys):
    with pytest.raises(SystemExit):
        main(["--help"])
    ayuda = " ".join(capsys.readouterr().out.split())
    assert "--citas" in ayuda and "muestra las citas que hay en el texto" in ayuda
