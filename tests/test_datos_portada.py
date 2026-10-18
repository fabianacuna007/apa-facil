"""
Paso 5.1: los datos de la portada.

Llegan en un archivo JSON (--datos) o con opciones sueltas (--titulo, --autor…),
que mandan sobre el archivo dato a dato. Se guardan tal cual (solo sin los
espacios de los extremos), lo que falta se avisa sin rellenarlo y lo que está
mal escrito para todo con un mensaje en español. Todavía no cambian el
documento: la portada la genera el paso 5.2.

tests/datos/01_datos.json lleva los datos de la portada del documento 01
(título, autor, asignatura y profesora) y, para lo que esa portada no tiene,
valores de prueba claramente ficticios («Universidad de Ejemplo»,
«Departamento de Ejemplo», «EJM 101» y una fecha), decisión del propietario.
"""

import json
import zipfile
from datetime import date
from pathlib import Path

import pytest

from apa_facil import pipeline
from apa_facil.__main__ import main
from apa_facil.datos_portada import DATOS, DatosPortada, ErrorDatos, avisar, combinar, leer_fecha, leer_json
from apa_facil.informe import Informe

RAIZ = Path(__file__).resolve().parent.parent
DOC_01 = RAIZ / "pruebas" / "01_ensayo_basico.docx"
DATOS_01 = RAIZ / "tests" / "datos" / "01_datos.json"
EJEMPLO = RAIZ / "docs" / "datos_ejemplo.json"


def escribir(carpeta, contenido, nombre="datos.json"):
    ruta = carpeta / nombre
    if isinstance(contenido, bytes):
        ruta.write_bytes(contenido)
    else:
        ruta.write_text(contenido if isinstance(contenido, str) else json.dumps(contenido, ensure_ascii=False),
                        encoding="utf-8")
    return ruta


def error_de(funcion, *args):
    with pytest.raises(ErrorDatos) as e:
        funcion(*args)
    return str(e.value)


# --- Lectura del archivo -------------------------------------------------------------------

def test_datos_del_documento_01():
    datos = combinar(DATOS_01)
    assert datos == DatosPortada(
        titulo="La Revolución Industrial y sus efectos en el mundo del trabajo",
        autores=("Javier Moreno Díaz",), departamento="Departamento de Ejemplo",
        universidad="Universidad de Ejemplo", codigo_asignatura="EJM 101",
        asignatura="Historia Contemporánea", profesores=("Elena Sáez",), fecha=date(2026, 10, 16))
    assert datos.faltan() == [] and datos.dados() == list(DATOS)


def test_el_ejemplo_de_docs_se_lee_completo():
    datos = combinar(EJEMPLO)
    assert datos.faltan() == []
    assert datos.autores == ("Ana Pérez Gómez", "Luis Martín Ruiz")


def test_el_texto_se_guarda_tal_cual(tmp_path):
    # Ni mayúsculas, ni espacios de dentro, ni signos: solo se quitan los espacios de los extremos.
    titulo = "LA revolución  industrial: ¿un cambio?"
    ruta = escribir(tmp_path, {"titulo": f"  {titulo}\t", "autores": ["  ana pérez ", "LUIS  GÓMEZ"],
                               "codigo_asignatura": "his-201 "})
    datos = combinar(ruta)
    assert datos.titulo == titulo
    assert datos.autores == ("ana pérez", "LUIS  GÓMEZ")
    assert datos.codigo_asignatura == "his-201"


def test_archivo_con_bom(tmp_path):
    # El Bloc de notas de Windows guarda en UTF-8 con una marca (BOM) al principio.
    ruta = escribir(tmp_path, '{"titulo": "Turismo rural en España"}'.encode("utf-8-sig"))
    assert combinar(ruta).titulo == "Turismo rural en España"


def test_claves_con_tildes_mayusculas_y_espacios(tmp_path):
    ruta = escribir(tmp_path, {"Título": "T", "Código asignatura": "C 1", "FECHA": "16/10/2026",
                               "Universidad": "U"})
    datos = combinar(ruta)
    assert (datos.titulo, datos.codigo_asignatura, datos.fecha, datos.universidad) == \
           ("T", "C 1", date(2026, 10, 16), "U")


def test_sinonimos_y_un_solo_texto_en_las_listas(tmp_path):
    ruta = escribir(tmp_path, {"autora": "Lucía Fernández Prieto", "profesores": ["Dr. A", "Dra. B"]})
    datos = combinar(ruta)
    assert datos.autores == ("Lucía Fernández Prieto",)
    assert datos.profesores == ("Dr. A", "Dra. B")


@pytest.mark.parametrize("vacio", [None, "", "   ", [], ["", " "]])
def test_lo_vacio_cuenta_como_que_falta(vacio, tmp_path):
    ruta = escribir(tmp_path, {"titulo": "T", "autores": vacio})
    assert "autores" in combinar(ruta).faltan()


def test_un_archivo_vacio_de_datos_es_una_portada_sin_nada(tmp_path):
    datos = combinar(escribir(tmp_path, "{}"))
    assert datos is not None and datos.faltan() == list(DATOS)


# --- Errores del archivo: se explica qué corregir ----------------------------------------------

def test_archivo_que_no_existe(tmp_path):
    assert "No se encuentra el archivo de datos" in error_de(leer_json, tmp_path / "no_existe.json")


def test_json_mal_escrito_dice_la_linea(tmp_path):
    ruta = escribir(tmp_path, '{\n  "titulo": "T",\n}')   # coma después del último dato
    mensaje = error_de(leer_json, ruta)
    assert "está mal escrito (línea 3, columna 1)" in mensaje
    assert "sin coma después del último" in mensaje


def test_archivo_que_no_esta_en_utf8(tmp_path):
    ruta = escribir(tmp_path, '{"autores": "Javier Moreno Díaz"}'.encode("cp1252"))
    assert "no está guardado en UTF-8" in error_de(leer_json, ruta)


def test_los_datos_tienen_que_ir_entre_llaves(tmp_path):
    assert "entre llaves" in error_de(leer_json, escribir(tmp_path, '["T", "A"]'))


def test_clave_desconocida_con_sugerencia(tmp_path):
    mensaje = error_de(leer_json, escribir(tmp_path, {"titulo": "T", "autr": "A"}))
    assert mensaje.startswith("«autr» no es un dato de portada. ¿Querías decir «autores»?")
    assert "titulo, autores, departamento" in mensaje


def test_clave_desconocida_sin_sugerencia(tmp_path):
    mensaje = error_de(leer_json, escribir(tmp_path, {"color": "rojo"}))
    assert mensaje.startswith("«color» no es un dato de portada. Los datos posibles")


def test_sugerencia_de_un_sinonimo(tmp_path):
    # «profesorr» se parece a «profesor»; «profesra», a «profesora», que es el dato «profesor».
    assert "¿Querías decir «profesor»?" in error_de(leer_json, escribir(tmp_path, {"profesra": "X"}))


@pytest.mark.parametrize("contenido", ['{"titulo": "A", "titulo": "B"}', '{"autor": "A", "autores": ["B"]}'])
def test_un_dato_repetido(contenido, tmp_path):
    assert "aparece dos veces" in error_de(leer_json, escribir(tmp_path, contenido))


@pytest.mark.parametrize("contenido, esperado", [
    ({"titulo": 2026}, "tiene que ser un texto entre comillas"),
    ({"titulo": ["A", "B"]}, "tiene que ser un texto entre comillas"),
    ({"autores": ["Ana", 3]}, "tiene que ser un texto o una lista de textos"),
    ({"profesor": {"nombre": "X"}}, "tiene que ser un texto o una lista de textos"),
    ({"fecha": 20261016}, "La fecha tiene que ser un texto"),
])
def test_tipos_incorrectos(contenido, esperado, tmp_path):
    assert esperado in error_de(leer_json, escribir(tmp_path, contenido))


# --- Fechas ---------------------------------------------------------------------------

@pytest.mark.parametrize("texto, esperada", [
    ("16/10/2026", date(2026, 10, 16)), ("2026-10-16", date(2026, 10, 16)),
    ("16 de octubre de 2026", date(2026, 10, 16)), ("16-10-2026", date(2026, 10, 16)),
    ("1/9/2026", date(2026, 9, 1)), ("1 de Septiembre de 2026", date(2026, 9, 1)),
    ("5 de setiembre del 2026", date(2026, 9, 5)), ("  29/02/2028 ", date(2028, 2, 29)),
])
def test_fechas_validas(texto, esperada):
    assert leer_fecha(texto) == esperada


@pytest.mark.parametrize("texto", ["31/02/2026", "29/02/2027", "16/13/2026", "2026-10-32", "0 de mayo de 2026"])
def test_fechas_que_no_existen(texto):
    assert error_de(leer_fecha, texto) == f"La fecha «{texto}» no existe. Revisa el día y el mes."


@pytest.mark.parametrize("texto", ["16/10/26", "octubre de 2026", "16 de octubrr de 2026", "2026/10/16",
                                   "mañana", "16 10 2026"])
def test_fechas_que_no_se_entienden(texto):
    assert error_de(leer_fecha, texto).startswith(f"No se entiende la fecha «{texto}». Escríbela así: «16/10/2026»")


# --- Opciones sueltas ---------------------------------------------------------------------

def test_sin_ningun_dato_no_hay_portada():
    assert combinar() is None
    assert combinar(None, {clave: None for clave in DATOS}) is None


def test_las_opciones_mandan_sobre_el_archivo_dato_a_dato():
    datos = combinar(DATOS_01, {"fecha": "20 de octubre de 2026", "autores": ["Ana", "Luis"], "titulo": None})
    assert datos.fecha == date(2026, 10, 20)
    assert datos.autores == ("Ana", "Luis")      # la lista entera, no se suma a la del archivo
    assert datos.titulo.startswith("La Revolución Industrial")   # None: opción no usada
    assert datos.profesores == ("Elena Sáez",)


def test_una_opcion_vacia_deja_el_dato_sin_dar():
    assert "universidad" in combinar(DATOS_01, {"universidad": ""}).faltan()


def test_la_fecha_de_una_opcion_tambien_se_comprueba():
    assert "no existe" in error_de(combinar, None, {"fecha": "31/04/2026"})


# --- Avisos ------------------------------------------------------------------------------

def test_con_todos_los_datos_no_hay_aviso():
    informe = Informe()
    avisar(combinar(DATOS_01), informe)
    assert informe.avisos == []


def test_aviso_de_un_dato_que_falta():
    informe = Informe()
    avisar(combinar(DATOS_01, {"fecha": ""}), informe)
    assert informe.avisos == ["Faltan datos para la portada: la fecha de entrega. No se han rellenado: "
                              "añádelos en tu archivo de datos o con la opción --fecha."]


def test_aviso_de_varios_datos_que_faltan():
    informe = Informe()
    avisar(combinar(None, {"titulo": "T", "autores": ["A"], "asignatura": "S", "profesor": ["P"]}), informe)
    assert informe.avisos == ["Faltan datos para la portada: el departamento, la universidad, el código de la "
                              "asignatura y la fecha de entrega. No se han rellenado: añádelos en tu archivo de "
                              "datos o con las opciones --departamento, --universidad, --codigo-asignatura y --fecha."]


# --- Línea de comandos y pipeline --------------------------------------------------------------

def contenido_docx(ruta):
    """Todas las partes del .docx menos las propiedades (llevan la hora de guardado)."""
    with zipfile.ZipFile(ruta) as z:
        return {n: z.read(n) for n in z.namelist() if not n.startswith("docProps/")}


def test_el_documento_01_con_datos_json(tmp_path, capsys):
    con, sin = tmp_path / "con", tmp_path / "sin"
    assert main([str(DOC_01), "--datos", str(DATOS_01)], carpeta_salida=con) == 0
    assert "Faltan datos para la portada" not in capsys.readouterr().out
    informe = json.loads((con / "01_ensayo_basico_APA.json").read_text(encoding="utf-8"))
    assert informe["opciones"]["portada"] == {"dados": list(DATOS), "faltan": []}
    # En el 5.1 los datos todavía no cambian el documento: la copia es la misma que sin ellos.
    assert main([str(DOC_01)], carpeta_salida=sin) == 0
    assert contenido_docx(con / "01_ensayo_basico_APA.docx") == contenido_docx(sin / "01_ensayo_basico_APA.docx")


def test_los_valores_no_se_guardan_en_el_json(tmp_path):
    # Son datos personales: el informe solo dice qué datos se han dado.
    salida, _ = pipeline.procesar(DOC_01, tmp_path, {"portada": combinar(None, {"autores": ["Persona Zzyzx"],
                                                                               "universidad": "Universidad Qwfp"})})
    texto = salida.with_suffix(".json").read_text(encoding="utf-8")
    assert "Zzyzx" not in texto and "Qwfp" not in texto
    assert json.loads(texto)["opciones"]["portada"]["dados"] == ["autores", "universidad"]


def test_opciones_sueltas_y_aviso_en_el_informe(tmp_path, capsys):
    assert main([str(DOC_01), "--titulo", "T", "--autor", "Ana", "--autor", "Luis", "--profesor", "P",
                 "--fecha", "16/10/2026"], carpeta_salida=tmp_path) == 0
    salida = capsys.readouterr().out
    assert ("Faltan datos para la portada: el departamento, la universidad, el código de la asignatura y el "
            "nombre de la asignatura.") in salida
    informe = json.loads((tmp_path / "01_ensayo_basico_APA.json").read_text(encoding="utf-8"))
    aviso = next(e for e in informe["avisos"] if e["mensaje"].startswith("Faltan datos"))
    assert (aviso["regla"], aviso["nombre"]) == ("portada", "Portada")
    assert informe["opciones"]["portada"]["dados"] == ["titulo", "autores", "profesor", "fecha"]


def test_sin_datos_de_portada_no_hay_aviso(tmp_path, capsys):
    assert main([str(DOC_01)], carpeta_salida=tmp_path) == 0
    assert "Faltan datos para la portada" not in capsys.readouterr().out
    informe = json.loads((tmp_path / "01_ensayo_basico_APA.json").read_text(encoding="utf-8"))
    assert informe["opciones"]["portada"] is None


@pytest.mark.parametrize("argumentos, esperado", [
    (["--fecha", "31/02/2026"], "La fecha «31/02/2026» no existe."),
    (["--datos", "no_existe.json"], "No se encuentra el archivo de datos «no_existe.json»."),
])
def test_un_dato_mal_escrito_no_procesa_nada(argumentos, esperado, tmp_path, capsys):
    with pytest.raises(SystemExit) as fin:
        main([str(DOC_01), *argumentos], carpeta_salida=tmp_path)
    assert fin.value.code == 2
    error = capsys.readouterr().err
    assert error.startswith(esperado) and error.rstrip().endswith("No se ha procesado ningún documento.")
    assert not list(tmp_path.iterdir()), "no se ha guardado ninguna copia"


def test_falta_el_archivo_despues_de_datos(capsys):
    with pytest.raises(SystemExit):
        main([str(DOC_01), "--datos"])
    assert "Falta indicar el valor después de --datos." in capsys.readouterr().err


def test_los_mismos_datos_para_varios_documentos(tmp_path, capsys):
    otro = RAIZ / "pruebas" / "02_informe_tablas.docx"
    assert main([str(DOC_01), str(otro), "--datos", str(DATOS_01)], carpeta_salida=tmp_path) == 0
    assert "Aviso: se usarán los mismos datos de portada en los 2 documentos." in capsys.readouterr().out


def test_el_diagnostico_no_usa_los_datos(capsys):
    # Ni siquiera lee el archivo: el diagnóstico no cambia nada.
    assert main([str(DOC_01), "--diagnostico", "--datos", "no_existe.json"]) == 0
    assert "Aviso: los datos de la portada no se usan en el modo diagnóstico" in capsys.readouterr().out


def test_la_ayuda_explica_los_datos_de_portada(capsys):
    with pytest.raises(SystemExit):
        main(["-h"])
    ayuda = capsys.readouterr().out
    assert "datos de la portada:" in ayuda
    for opcion, _ in DATOS.values():
        assert opcion in ayuda
    assert "--datos datos.json" in ayuda and "docs/datos_ejemplo.json" in ayuda
