"""
Paso 7.4: errores de forma en las citas y las referencias
(apa_facil/revision/forma.py), según la web oficial de APA y la especificación
en español del ROADMAP. Cada error trae lo escrito y la forma correcta que se
propone; el documento no cambia.

Cada error se prueba con un caso que lo tiene y otro correcto que no debe dar
nada. Fuera de este paso (decisión del propietario): la cursiva y la raya corta
de los rangos de páginas, que saltarían en el 02 (el documento de control).
"""

from pathlib import Path

import pytest

from apa_facil import documento, estructura
from apa_facil.__main__ import main
from apa_facil.revision import citas, forma, referencias
from apa_facil.revision.forma import (AMPERSAND, AÑO_SIN_PARENTESIS, CIUDAD, COMA, DOI_ANTIGUO, ESPACIO_EN_PAGINA,
                                      ET_AL, FALTA_AÑO, INCOMPLETA, NO_RECONOCIDA, NUMERO_ABREVIADO, ORDEN, P_O_PP,
                                      PUNTO_TRAS_AUTOR, PUNTO_TRAS_FECHA, PUNTO_Y_COMA, RECUPERADO_DE, SIN_FECHA)

RAIZ = Path(__file__).resolve().parent.parent
PRUEBAS = {ruta.name[:2]: ruta for ruta in sorted((RAIZ / "pruebas").glob("*.docx"))}


def en_citas(texto):
    return [(e.tipo, e.escrito, e.correccion) for e in forma.revisar(citas.citas_del_texto(texto, "Párrafo 1", 1), [])]


def en_referencias(*textos):
    leidas = [referencias.leer_texto(t, ubicacion=f"Párrafo {i}", numero=i) for i, t in enumerate(textos, start=10)]
    return [(e.tipo, e.escrito, e.correccion) for e in forma.revisar([], leidas)]


# --- Citas ----------------------------------------------------------------------------

@pytest.mark.parametrize("texto, esperado", [
    ("(Martínez 2019)", [(COMA, "(Martínez 2019)", "(Martínez, 2019)")]),
    ("(Pérez et. al, 2020)", [(ET_AL, "(Pérez et. al, 2020)", "(Pérez et al., 2020)")]),
    ("Pérez et al (2020)", [(ET_AL, "Pérez et al (2020)", "Pérez et al. (2020)")]),
    ("Gómez (s.f)", [(SIN_FECHA, "Gómez (s.f)", "Gómez (s. f.)")]),
    ("(Gómez, s.f.)", [(SIN_FECHA, "(Gómez, s.f.)", "(Gómez, s. f.)")]),
    ("(Gómez, n.d.)", [(SIN_FECHA, "(Gómez, n.d.)", "(Gómez, s. f.)")]),
    ("Zhou (in press)", [(SIN_FECHA, "Zhou (in press)", "Zhou (en prensa)")]),
    ("(Navarro & Gil, 2020)", [(AMPERSAND, "(Navarro & Gil, 2020)", "(Navarro y Gil, 2020)")]),
    ("(Consejo Económico de Ejemplo, 2022, p.45)",
     [(ESPACIO_EN_PAGINA, "(Consejo Económico de Ejemplo, 2022, p.45)", "(Consejo Económico de Ejemplo, 2022, p. 45)")]),
    ("(Lefèvre, 2009, p. 112–115)", [(P_O_PP, "(Lefèvre, 2009, p. 112–115)", "(Lefèvre, 2009, pp. 112–115)")]),
    ("(Lefèvre, 2009, pp. 112)", [(P_O_PP, "(Lefèvre, 2009, pp. 112)", "(Lefèvre, 2009, p. 112)")]),
])
def test_errores_en_las_citas(texto, esperado):
    assert en_citas(texto) == esperado


@pytest.mark.parametrize("texto", [
    "(Martínez, 2019)", "(Pérez et al., 2020)", "Pérez et al. (2020)", "Gómez (s. f.)", "(Gómez, s. f.)",
    "Zhou (en prensa)", "(Navarro y Gil, 2020)", "Navarro y Gil (2020)", "(Lefèvre, 2009, p. 112)",
    "(Lefèvre, 2009, pp. 112–115)", "(Pérez, 2019, pp. 67, 72)", "Webster-Stratton (2016, pp. 152–153)",
])
def test_citas_correctas(texto):
    assert en_citas(texto) == []


def test_varios_errores_en_una_cita_y_uno_por_parentesis():
    """Cada error se dice una vez por paréntesis, aunque el paréntesis cite varias obras."""
    assert en_citas("(Pérez et. al, 2020; Gómez et. al, 2021)") == [
        (ET_AL, "(Pérez et. al, 2020; Gómez et. al, 2021)", "(Pérez et al., 2020; Gómez et. al, 2021)")]
    assert [tipo for tipo, _, _ in en_citas("(Navarro & Gil 2020, p.45)")] == [COMA, AMPERSAND, ESPACIO_EN_PAGINA]


def test_lo_que_ya_dice_el_cruce_no_se_repite():
    """«et al.» con dos autores o coautores que faltan son del 7.3: aquí no salen."""
    assert en_citas("(Navarro et al., 2020)") == []


# --- Referencias ----------------------------------------------------------------------

def test_ampersand_y_punto_y_coma():
    assert en_referencias("Campos, F., & Vidal, N. (2022). Autoconsumo. Revista de Energía, 8, 15-29.") == [
        (AMPERSAND, "Campos, F., & Vidal, N.", "Campos, F. y Vidal, N.")]
    assert en_referencias("Pérez, L.; Romero, A.; Saiz, T. (2020). Autoestima. Revista Joven, 7(2), 1-9.") == [
        (PUNTO_Y_COMA, "Pérez, L.; Romero, A.; Saiz, T.", "Pérez, L., Romero, A. y Saiz, T.")]


@pytest.mark.parametrize("doi, correcto", [
    ("doi:10.0000/rerf.2015.041", "https://doi.org/10.0000/rerf.2015.041"),
    ("http://dx.doi.org/10.1037/a0040251", "https://doi.org/10.1037/a0040251"),
    ("https://doi.org/10.1037/a0040251", None),
])
def test_doi(doi, correcto):
    errores = en_referencias(f"Herranz, M. (2015). Turismo. Revista de Estudios, 41, 77-96. {doi}")
    assert errores == ([(DOI_ANTIGUO, doi, correcto)] if correcto else [])


@pytest.mark.parametrize("fuente, ciudad, editorial", [
    ("Editorial Síntesis, Madrid", "Madrid", "Editorial Síntesis"),
    ("Paris: Éditions Exemple", "Paris", "Éditions Exemple"),
])
def test_ciudad_de_la_editorial(fuente, ciudad, editorial):
    assert en_referencias(f"Lefèvre, A. (2009). Industrialisation. {fuente}.") == [(CIUDAD, ciudad, editorial)]


def test_recuperado_de():
    assert en_referencias("Martínez, J. (2019). Redes sociales. Recuperado de https://www.example.org/redes") == [
        (RECUPERADO_DE, "Recuperado de https://www.example.org/redes", "https://www.example.org/redes")]
    # Con la fecha de consulta sí vale (APA, Elements: retrieval dates, para las obras que cambian).
    assert en_referencias("Martínez, J. (2019). Redes sociales. Recuperado el 30 de septiembre de 2020, de "
                          "https://www.example.org/redes") == []


def test_el_año():
    assert en_referencias("Herranz, M. Turismo rural. Revista de Estudios, 41, 77-96.") == [
        (FALTA_AÑO, "Herranz, M.", "Herranz, M. (s. f.).")]
    assert en_referencias("Lefèvre, A. 2009. Industrialisation. Éditions Exemple.") == [
        (AÑO_SIN_PARENTESIS, "2009", "Lefèvre, A. (2009).")]
    (detras,) = forma.revisar([], [referencias.leer_texto("Martínez, J. Redes sociales. 2019. https://example.org/r")])
    assert detras.tipo == AÑO_SIN_PARENTESIS and "justo detrás del autor" in detras.explicacion


def test_puntos_entre_elementos():
    assert en_referencias("Consejo Económico de Ejemplo (2022). Informe sobre el medio rural. Editorial.") == [
        (PUNTO_TRAS_AUTOR, "Consejo Económico de Ejemplo (2022)", "Consejo Económico de Ejemplo. (2022)")]
    assert en_referencias("Molina, R. (2017) Turismo y territorio. Editorial Ejemplo.") == [
        (PUNTO_TRAS_FECHA, "(2017)", "(2017).")]
    assert en_referencias("National Institute of Mental Health. (2018, July). Anxiety disorders. "
                          "https://www.nimh.nih.gov/health") == []


def test_numero_de_la_revista_con_abreviatura():
    assert en_referencias("Pérez, L. (2020). Autoestima digital. Revista Joven, 12, nº 7, 1-9.") == [
        (NUMERO_ABREVIADO, "nº 7", "12(7)")]


@pytest.mark.parametrize("texto, explicacion", [
    ("Pérez, L. (2020). Autoestima digital. Revista Ejemplo Joven, (7).", "Faltan el volumen y las páginas del artículo."),
    ("Torres, E. (2019). Agroturismo y empleo femenino.", "Falta la editorial (o el enlace, si es una obra de internet)."),
])
def test_referencias_incompletas(texto, explicacion):
    (error,) = forma.revisar([], [referencias.leer_texto(texto)])
    assert (error.tipo, error.explicacion) == (INCOMPLETA, explicacion)


@pytest.mark.parametrize("texto", [
    # Ejemplos de la web de APA: sin volumen (Lipscomb), con número de artículo, organización sin editorial.
    "Lipscomb, A. Y. (2021, Winter). Addressing trauma. The Journal of College Admission, (249), 30–33.",
    "Jerrentrup, A. (2018). Teaching medicine. PLoS ONE, 13(3), Article e0193972. https://doi.org/10.1371/x",
    "Consejo Económico de Ejemplo. (2022). Informe sobre el medio rural.",
    "Martínez, J. (2019). Redes sociales y adolescencia. https://www.example.org/martinez-redes",
    "Castro, P., Llorente, M. y Benítez, J. (2019). Calidad de sueño. Revista de Psicología, 12(3), 201-219.",
])
def test_referencias_correctas(texto):
    assert en_referencias(texto) == []


def test_no_reconocida():
    (error,) = forma.revisar([], [referencias.leer_texto("www.example.org/datos", ubicacion="Párrafo 68", numero=68)])
    assert (error.tipo, error.correccion) == (NO_RECONOCIDA, None)
    assert error.explicacion.startswith("Solo tiene una dirección web") and "APA pide autor, fecha" in error.explicacion


def test_orden_alfabetico():
    """Fuera de su sitio está la que rompe el orden, no las de detrás: en «Molina, Collantes, Herranz», Molina."""
    errores = forma.revisar([], [referencias.leer_texto(t, ubicacion=f"Párrafo {n}", numero=n) for n, t in (
        (1, "Molina, R. (2017). Turismo. Editorial."), (2, "Collantes, F. (2011). Lugares. Editorial."),
        (3, "Herranz, M. (2015). Rural. Editorial."), (4, "Sáenz Gil, P. (2020). Pobladores. Editorial."))])
    (error,) = errores
    assert (error.tipo, error.numero, error.escrito) == (ORDEN, 1, "Molina, R.")
    assert error.explicacion.endswith("después de «Herranz, M.» y antes de «Sáenz Gil, P.».")
    assert forma.revisar([], [referencias.leer_texto("Collantes, F. (2011). Lugares. Editorial.", numero=1),
                              referencias.leer_texto("Molina, R. (2017). Turismo. Editorial.", numero=2)]) == []


# --- Los cinco documentos de prueba (la tabla de la fase 7 del ROADMAP) ---------------

ESPERADO = {
    "01": [(20, CIUDAD), (21, ORDEN), (22, CIUDAD), (22, AÑO_SIN_PARENTESIS), (22, ORDEN), (23, NO_RECONOCIDA)],
    "02": [],
    "03": [(29, AMPERSAND), (29, ORDEN)],
    "04": [(7, COMA), (11, ET_AL), (12, SIN_FECHA), (44, ET_AL), (51, SIN_FECHA), (58, SIN_FECHA),
           (66, AÑO_SIN_PARENTESIS), (67, PUNTO_Y_COMA), (67, PUNTO_TRAS_FECHA), (67, NUMERO_ABREVIADO),
           (67, INCOMPLETA), (68, NO_RECONOCIDA)],
    "05": [(32, ESPACIO_EN_PAGINA), (68, ORDEN), (70, CIUDAD), (70, PUNTO_TRAS_AUTOR), (71, DOI_ANTIGUO),
           (71, FALTA_AÑO)],
}


def errores_de(clave):
    doc = documento.abrir(PRUEBAS[clave])
    clasificacion = estructura.clasificar(doc)
    return forma.revisar(citas.detectar(doc, clasificacion), referencias.leer(doc, clasificacion))


@pytest.mark.parametrize("clave", sorted(ESPERADO))
def test_cinco_documentos(clave):
    assert [(e.numero, e.tipo) for e in errores_de(clave)] == ESPERADO[clave]


def test_las_correcciones_de_los_documentos():
    quinto = {e.tipo: (e.escrito, e.correccion) for e in errores_de("05")}
    assert quinto[DOI_ANTIGUO] == ("doi:10.0000/rerf.2015.041", "https://doi.org/10.0000/rerf.2015.041")
    assert quinto[FALTA_AÑO] == ("Herranz, M.", "Herranz, M. (s. f.).")
    primero = [(e.escrito, e.correccion) for e in errores_de("01") if e.tipo == CIUDAD]
    assert primero == [("Madrid", "Editorial Síntesis Ficticia"), ("Paris", "Éditions Exemple")]


# --- Opción --forma -------------------------------------------------------------------

def test_opcion_forma_no_genera_copias(tmp_path, capsys):
    assert main([str(PRUEBAS["04"]), "--forma"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert list(tmp_path.iterdir()) == []
    assert "Errores de forma de 04_formato_caotico.docx" in salida
    assert "En las citas (6)" in salida and "En las referencias (6)" in salida
    assert "«(Martínez 2019)» → «(Martínez, 2019)»" in salida
    assert "1 de 1 documento(s) revisado(s). No se ha generado ninguna copia." in salida


def test_opcion_forma_sin_errores(capsys):
    assert main([str(PRUEBAS["02"]), "--forma", "--fuente", "Arial"]) == 0
    salida = capsys.readouterr().out
    assert "la opción --fuente no se usa en la revisión de la forma" in salida
    assert "No hay errores de forma en las citas ni en las referencias." in salida


def test_opcion_forma_con_un_archivo_que_no_existe(capsys):
    assert main(["no_existe.docx", "--forma"]) == 1


def test_la_ayuda_explica_la_opcion_forma(capsys):
    with pytest.raises(SystemExit):
        main(["--help"])
    ayuda = " ".join(capsys.readouterr().out.split())
    assert "--forma" in ayuda and "muestra los errores de forma de las citas y las referencias" in ayuda
