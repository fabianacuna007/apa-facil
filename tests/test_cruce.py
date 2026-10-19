"""
Paso 7.3: cruce de citas y referencias (apa_facil/revision/cruce.py), según la
web oficial de APA: cada obra citada debe estar en la lista y cada referencia
debe citarse, con la misma grafía de los autores y la misma fecha (Basic
principles of citation); cada cita corresponde a una sola referencia
(Author–date citation system); las comunicaciones personales no llevan
referencia (Personal communications).

Los casos pequeños parten de texto: las citas las reconoce el detector del 7.1 y
las referencias, el lector del 7.2, así que se prueba la cadena entera.
"""

from pathlib import Path

import pytest

from apa_facil import documento, estructura
from apa_facil.__main__ import main
from apa_facil.revision import citas, cruce, referencias
from apa_facil.revision.cruce import (AMBIGUA, AUTORES_DISTINTOS, AÑO_DISTINTO, NO_CITADA, NO_RECONOCIDA,
                                      SIN_REFERENCIA)

RAIZ = Path(__file__).resolve().parent.parent
PRUEBAS = {ruta.name[:2]: ruta for ruta in sorted((RAIZ / "pruebas").glob("*.docx"))}


def cruzar(texto, *lista):
    """Cruza las citas de `texto` (un párrafo) con las referencias de `lista` (una por elemento)."""
    encontradas = citas.citas_del_texto(texto, "Párrafo 1", 1)
    leidas = [referencias.leer_texto(r, ubicacion=f"Párrafo {i}", numero=i) for i, r in enumerate(lista, start=10)]
    return cruce.cruzar(encontradas, leidas)


def resumen(resultado):
    """{tipo: [(autores de la cita, año), …] o [autores de la referencia, …]} de lo que hay que revisar."""
    return {tipo: [(h.citas[0].autores_escritos, h.citas[0].año) if h.citas else h.referencias[0].autores_escritos
                   for h in hallazgos]
            for tipo, hallazgos in resultado.grupos() if hallazgos}


MOLINA = "Molina, R. (2017). Turismo y territorio. Editorial Ejemplo."
NAVARRO = "Navarro, A. y Gil, S. (2020). Sueño, memoria y aprendizaje. Editorial Ejemplo."
CASTRO = "Castro, P., Llorente, M. y Benítez, J. (2019). Calidad de sueño. Revista de Psicología, 12(3), 201-219."


# --- Cada grupo por separado ----------------------------------------------------------

def test_todo_coincide():
    resultado = cruzar("Según Molina (2017) y otros (Navarro y Gil, 2020; Castro et al., 2019).", MOLINA, NAVARRO, CASTRO)
    assert resumen(resultado) == {}
    assert len(resultado.emparejadas) == 3
    assert [h.referencias[0].numero for h in resultado.emparejadas] == [10, 11, 12]


def test_cita_sin_referencia_y_referencia_sin_cita():
    resultado = cruzar("Lo dice Castell (2011).", MOLINA)
    assert resumen(resultado) == {SIN_REFERENCIA: [("Castell", "2011")], NO_CITADA: ["Molina, R."]}


def test_una_obra_citada_varias_veces_sale_una_vez_con_todas_sus_ubicaciones():
    encontradas = (citas.citas_del_texto("Según Gómez (s.f), lo importante.", "Párrafo 12", 12)
                   + citas.citas_del_texto("Gómez, s.f.", "Párrafo 58 (dentro de una tabla)", 58, celda=True))
    resultado = cruce.cruzar(encontradas, [])
    (hallazgo,) = resultado.sin_referencia
    assert [c.numero for c in hallazgo.citas] == [12, 58]
    assert "párrafos 12 y 58 (tabla)" in cruce.como_texto(resultado)


@pytest.mark.parametrize("cita, referencia, detalle", [
    ("Herranz (2015)", "Herranz, M. Turismo rural. Revista de Estudios, 41, 77-96.", "la referencia no tiene año"),
    ("(Molina, 2016)", MOLINA, "en la referencia, 2017"),
])
def test_años_que_no_coinciden(cita, referencia, detalle):
    resultado = cruzar(cita, referencia)
    (hallazgo,) = resultado.años_distintos
    assert hallazgo.detalle == detalle
    assert resultado.no_citadas == []   # la referencia sí se cita, aunque con otro año


def test_s_f_mal_escrito_coincide_igual():
    """Cómo se escribe «s. f.» lo revisa el 7.4; aquí «s.f», «s. f.» y «n.d.» son lo mismo."""
    for cita in ("(Gómez, s.f)", "Gómez (s. f.)", "(Gómez, n.d.)"):
        assert resumen(cruzar(cita, "Gómez, L. (s. f.). Redes y autoestima. Editorial Ejemplo.")) == {}


def test_y_o_ampersand_y_et_al_mal_escrito_coinciden_igual():
    assert resumen(cruzar("(Navarro & Gil, 2020)", NAVARRO)) == {}
    assert resumen(cruzar("(Castro et. al, 2019)", CASTRO)) == {}


@pytest.mark.parametrize("cita, referencia, parte_del_detalle", [
    ("(Ramirez Soto, 2015)", "Ramírez Soto, M. (2015). Fábricas. Editorial.", "cambian las tildes"),
    ("(Saenz-Gil, 2020)", "Sáenz Gil, P. (2020). Pobladores. Editorial.", "cambian las tildes"),
    ("(Ramírez, 2015)", "Ramírez Soto, M. (2015). Fábricas. Editorial.", "falta o sobra un apellido"),
    ("(Colantes y Pinilla, 2011)", "Collantes, F. y Pinilla, V. (2011). Lugares. Editorial.", "parece una errata"),
    ("(Navarro, 2020)", NAVARRO, "faltan autores en la cita: la referencia tiene 2 autores (Navarro y Gil)"),
    ("(Navarro, Gil y Ruiz, 2020)", NAVARRO, "sobran autores en la cita"),
    ("(Navarro et al., 2020)", NAVARRO, "la cita pone «et al.», pero la referencia tiene 2 autores"),
    ("(Consejo Económico, 2022)", "Consejo Económico de Ejemplo (2022). Informe. Madrid.",
     "solo tiene el principio del nombre"),
    ("(Consejo Economico de Ejemplo, 2022)", "Consejo Económico de Ejemplo (2022). Informe. Madrid.",
     "cambian las tildes"),
])
def test_autores_escritos_de_forma_distinta(cita, referencia, parte_del_detalle):
    resultado = cruzar(cita, referencia)
    (hallazgo,) = resultado.autores_distintos
    assert parte_del_detalle in hallazgo.detalle
    assert resultado.no_citadas == []


@pytest.mark.parametrize("cita", ["(Molina, 2017)", "(Gil, 2020)", "(Pérez, 2020)"])
def test_un_apellido_corto_distinto_es_otro_autor(cita):
    """Una letra distinta en un apellido corto («Gil» y «Gal») es otro autor, no una errata."""
    resultado = cruzar(cita, "Gal, S. (2020). Otro libro. Editorial.")
    assert resultado.sin_referencia and not resultado.autores_distintos


def test_parecida_pero_de_otro_año_es_sin_referencia_con_pista():
    resultado = cruzar("(Ramirez, 2014)", "Ramírez Soto, M. (2015). Fábricas. Editorial.")
    (hallazgo,) = resultado.sin_referencia
    assert hallazgo.detalle == ("se parece a «Ramírez Soto, M. (2015)», párrafo 10, pero no coinciden ni los "
                                "autores ni el año")


def test_citas_ambiguas():
    """APA: dos obras del mismo autor y año llevan letra (2017a, 2017b); dos Pérez distintos, sus iniciales."""
    mismo_año = cruzar("(Molina, 2017)", MOLINA, "Molina, R. (2017). Otro libro. Editorial Ejemplo.")
    assert resumen(mismo_año) == {AMBIGUA: [("Molina", "2017")]}
    assert "APA las distingue con una letra detrás del año" in mismo_año.ambiguas[0].detalle
    con_letras = cruzar("(Molina, 2017)", "Molina, R. (2017a). Uno. Editorial.", "Molina, R. (2017b). Dos. Editorial.")
    assert resumen(con_letras) == {AMBIGUA: [("Molina", "2017")]}
    assert "falta la letra del año" in con_letras.ambiguas[0].detalle
    dos_perez = cruzar("(Pérez, 2019)", "Pérez, A. (2019). Uno. Editorial.", "Pérez, L. (2019). Dos. Editorial.")
    assert resumen(dos_perez) == {AMBIGUA: [("Pérez", "2019")]}
    bien = cruzar("(Molina, 2017a; Molina, 2017b)", "Molina, R. (2017a). Uno. Editorial.",
                  "Molina, R. (2017b). Dos. Editorial.")
    assert resumen(bien) == {}


def test_organizacion_con_abreviatura():
    """APA: la abreviatura se define la primera vez («[NIMH]») y después vale por el nombre entero."""
    resultado = cruzar("Primero (National Institute of Mental Health [NIMH], 2018) y luego (NIMH, 2018).",
                       "National Institute of Mental Health. (2018, July). Anxiety disorders. U.S. Department of "
                       "Health and Human Services. https://www.nimh.nih.gov/health")
    assert resumen(resultado) == {}
    assert len(resultado.emparejadas) == 1 and len(resultado.emparejadas[0].citas) == 2


def test_obra_sin_autor_citada_por_su_titulo():
    """APA (Missing reference information): «Título. (Fecha). Fuente.», citada por el principio del título."""
    resultado = cruzar("(“Estudio sobre el sueño”, 2019)",
                       "Estudio sobre el sueño en universitarios. (2019). Revista de Psicología, 12(3), 201-219.")
    assert resumen(resultado) == {}


def test_comunicacion_personal_no_pide_referencia():
    resultado = cruzar("Lo confirmó (T. Nguyen, comunicación personal, 24 de febrero de 2020), y "
                       "E.-M. Paradis (personal communication, August 8, 2019) también.")
    assert resumen(resultado) == {}


def test_referencias_no_reconocidas_no_se_dan_por_no_citadas():
    resultado = cruzar("(Molina, 2017)", MOLINA, "www.example.org/datos-uso-redes-2024")
    assert resumen(resultado) == {NO_RECONOCIDA: [None]}
    assert resultado.no_citadas == []


# --- Los cinco documentos de prueba (la tabla de la fase 7 del ROADMAP) ---------------

ESPERADO = {
    "01": {SIN_REFERENCIA: [("Castell", "2011")], NO_RECONOCIDA: [None]},
    "02": {},
    "03": {},
    "04": {SIN_REFERENCIA: [("Gómez", "s.f"), ("INE", "s.f")], NO_RECONOCIDA: [None]},
    "05": {SIN_REFERENCIA: [("Ruiz Montero", "2023")], NO_CITADA: ["Torres, E."],
           AÑO_DISTINTO: [("Herranz", "2015")]},
}
EMPAREJADAS = {"01": 3, "02": 2, "03": 2, "04": 2, "05": 4}


def cruce_de(clave):
    doc = documento.abrir(PRUEBAS[clave])
    clasificacion = estructura.clasificar(doc)
    return cruce.cruzar(citas.detectar(doc, clasificacion), referencias.leer(doc, clasificacion))


@pytest.mark.parametrize("clave", sorted(ESPERADO))
def test_cinco_documentos(clave):
    resultado = cruce_de(clave)
    assert resumen(resultado) == ESPERADO[clave]
    assert len(resultado.emparejadas) == EMPAREJADAS[clave]


def test_las_ubicaciones_de_los_documentos():
    cuarto, quinto = cruce_de("04"), cruce_de("05")
    assert [[c.numero for c in h.citas] for h in cuarto.sin_referencia] == [[12, 58], [51]]
    (herranz,) = quinto.años_distintos
    assert ([c.numero for c in herranz.citas], herranz.referencias[0].numero, herranz.detalle) == (
        [33, 65], 71, "la referencia no tiene año")
    assert [h.referencias[0].numero for h in quinto.no_citadas] == [73]


# --- Opción --cruce -------------------------------------------------------------------

def test_opcion_cruce_no_genera_copias(tmp_path, capsys):
    assert main([str(PRUEBAS["05"]), "--cruce"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert list(tmp_path.iterdir()) == []
    assert "Cruce de citas y referencias de 05_tfg_completo.docx" in salida
    assert "Citas sin referencia en la lista (1)" in salida and "Ruiz Montero (2023) en el texto: párrafo 63" in salida
    assert "Referencias que no se citan en el texto (1)" in salida and "Torres, E. (2019), párrafo 73" in salida
    assert "↳ la referencia no tiene año" in salida
    assert "1 de 1 documento(s) revisado(s). No se ha generado ninguna copia." in salida


def test_opcion_cruce_sin_nada_que_revisar(capsys):
    assert main([str(PRUEBAS["02"]), "--cruce", "--fuente", "Arial"]) == 0
    salida = capsys.readouterr().out
    assert "la opción --fuente no se usa en el cruce de citas y referencias" in salida
    assert "No hay nada que revisar: todas las citas tienen su referencia y todas las referencias se citan." in salida


def test_opcion_cruce_con_un_archivo_que_no_existe(capsys):
    assert main(["no_existe.docx", "--cruce"]) == 1


def test_la_ayuda_explica_la_opcion_cruce(capsys):
    with pytest.raises(SystemExit):
        main(["--help"])
    ayuda = " ".join(capsys.readouterr().out.split())
    assert "--cruce" in ayuda and "cruza las citas del texto con la lista de referencias" in ayuda
