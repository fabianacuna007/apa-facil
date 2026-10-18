"""
Paso 5.1: los datos de la portada de estudiante.

Llegan de un archivo JSON (opción --datos) y de opciones sueltas de la línea de
comandos (--titulo, --autor…), que mandan sobre el archivo dato a dato:

    {
      "titulo": "La Revolución Industrial y sus efectos en el mundo del trabajo",
      "autores": ["Javier Moreno Díaz"],
      "departamento": "Departamento de Historia",
      "universidad": "Universidad de …",
      "codigo_asignatura": "HIS 201",
      "asignatura": "Historia Contemporánea",
      "profesor": "Elena Sáez",
      "fecha": "16/10/2026"
    }

Aquí solo se leen y se comprueban; la portada la genera el paso 5.2.

- El texto se guarda tal cual: solo se quitan los espacios de los extremos.
- Lo que falta no se rellena (ni siquiera la fecha con la de hoy): se avisa.
- Un dato mal escrito (JSON roto, clave desconocida, fecha imposible) para
  todo con un mensaje que explica qué corregir (ErrorDatos).
"""

import difflib
import json
import re
import unicodedata
from dataclasses import dataclass
from datetime import date
from pathlib import Path

# Cada dato: (clave en el JSON, opción de la línea de comandos, cómo se nombra en un mensaje).
DATOS = {
    "titulo": ("--titulo", "el título"),
    "autores": ("--autor", "el autor o los autores"),
    "departamento": ("--departamento", "el departamento"),
    "universidad": ("--universidad", "la universidad"),
    "codigo_asignatura": ("--codigo-asignatura", "el código de la asignatura"),
    "asignatura": ("--asignatura", "el nombre de la asignatura"),
    "profesor": ("--profesor", "el profesor o la profesora"),
    "fecha": ("--fecha", "la fecha de entrega"),
}
LISTAS = {"autores", "profesor"}   # admiten un texto o una lista de textos

# Otras formas de escribir la misma clave (ya sin tildes ni mayúsculas).
SINONIMOS = {
    "autor": "autores", "autora": "autores", "autoras": "autores",
    "profesora": "profesor", "profesores": "profesor", "profesoras": "profesor",
    "codigo": "codigo_asignatura", "codigo_de_la_asignatura": "codigo_asignatura",
    "fecha_de_entrega": "fecha", "fecha_entrega": "fecha",
}

MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
FORMATOS_FECHA = "«16/10/2026», «2026-10-16» o «16 de octubre de 2026»"


class _Pares(list):
    """Un objeto JSON leído como lista de pares (clave, valor), para ver las claves repetidas."""


def _atributo(clave):
    """El dato «profesor» (uno o varios) se guarda en DatosPortada.profesores."""
    return "profesores" if clave == "profesor" else clave


class ErrorDatos(Exception):
    """Un dato de portada que no se puede usar. El mensaje, en español, dice qué corregir."""


@dataclass
class DatosPortada:
    titulo: str = None
    autores: tuple = ()
    departamento: str = None
    universidad: str = None
    codigo_asignatura: str = None
    asignatura: str = None
    profesores: tuple = ()
    fecha: date = None

    def valor(self, clave):
        return getattr(self, _atributo(clave))

    def dados(self):
        """Claves de los datos que se han dado, en el orden de DATOS."""
        return [clave for clave in DATOS if self.valor(clave)]

    def faltan(self):
        """Claves de los datos que faltan, en el orden de DATOS."""
        return [clave for clave in DATOS if not self.valor(clave)]


# --- Claves ----------------------------------------------------------------------------

def _clave(texto):
    """«Código asignatura» → «codigo_asignatura»: sin tildes, en minúsculas y con guiones bajos."""
    sin_tildes = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    return re.sub(r"[\s\-]+", "_", sin_tildes.strip().lower())


def _clave_conocida(original):
    clave = _clave(original)
    clave = SINONIMOS.get(clave, clave)
    if clave in DATOS:
        return clave
    parecidas = difflib.get_close_matches(_clave(original), list(DATOS) + list(SINONIMOS), n=1, cutoff=0.6)
    pista = f" ¿Querías decir «{SINONIMOS.get(parecidas[0], parecidas[0])}»?" if parecidas else ""
    raise ErrorDatos(f"«{original}» no es un dato de portada.{pista} "
                     f"Los datos posibles son: {', '.join(DATOS)}.")


# --- Valores ---------------------------------------------------------------------------

def _texto(clave, valor):
    if valor is None:
        return None
    if not isinstance(valor, str):
        raise ErrorDatos(f"El dato «{clave}» tiene que ser un texto entre comillas, por ejemplo "
                         f"\"{clave}\": \"…\".")
    return valor.strip() or None


def _lista(clave, valor):
    if valor is None:
        return ()
    if isinstance(valor, str):
        valor = [valor]
    if not isinstance(valor, list) or not all(isinstance(v, str) for v in valor):
        raise ErrorDatos(f"El dato «{clave}» tiene que ser un texto o una lista de textos, por ejemplo "
                         f"\"{clave}\": [\"Ana Pérez\", \"Luis Gómez\"].")
    return tuple(v.strip() for v in valor if v.strip())


def leer_fecha(valor):
    """
    Convierte «16/10/2026», «2026-10-16» o «16 de octubre de 2026» en una fecha.
    Una fecha imposible («31/02/2026») o con otra forma es un ErrorDatos.
    """
    if valor is None:
        return None
    if not isinstance(valor, str):
        raise ErrorDatos(f"La fecha tiene que ser un texto entre comillas, por ejemplo \"fecha\": \"16/10/2026\".")
    texto = valor.strip()
    if not texto:
        return None
    if m := re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", texto):
        dia, mes, ano = int(m[1]), int(m[2]), int(m[3])
    elif m := re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", texto):
        ano, mes, dia = int(m[1]), int(m[2]), int(m[3])
    elif (m := re.fullmatch(r"(\d{1,2})\s+de\s+(\w+)\s+del?\s+(\d{4})", texto, re.IGNORECASE)) \
            and m[2].lower() in MESES:
        dia, mes, ano = int(m[1]), MESES[m[2].lower()], int(m[3])
    else:
        raise ErrorDatos(f"No se entiende la fecha «{texto}». Escríbela así: {FORMATOS_FECHA}.")
    try:
        return date(ano, mes, dia)
    except ValueError:
        raise ErrorDatos(f"La fecha «{texto}» no existe. Revisa el día y el mes.") from None


def _valor(clave, valor):
    if clave == "fecha":
        return leer_fecha(valor)
    return _lista(clave, valor) if clave in LISTAS else _texto(clave, valor)


# --- Archivo y opciones -------------------------------------------------------------------

def leer_json(ruta):
    """Lee un archivo de datos de portada. Devuelve {clave: valor ya comprobado}."""
    ruta = Path(ruta)
    try:
        # utf-8-sig: el Bloc de notas de Windows añade al principio una marca (BOM) que el JSON no admite.
        texto = ruta.read_bytes().decode("utf-8-sig")
    except FileNotFoundError:
        raise ErrorDatos(f"No se encuentra el archivo de datos «{ruta}». Revisa el nombre y la carpeta.") from None
    except IsADirectoryError:
        raise ErrorDatos(f"«{ruta}» es una carpeta, no un archivo de datos.") from None
    except PermissionError:
        raise ErrorDatos(f"No se puede leer el archivo de datos «{ruta.name}». Puede que otro programa "
                         "lo tenga abierto: ciérralo y vuelve a intentarlo.") from None
    except UnicodeDecodeError:
        raise ErrorDatos(f"El archivo de datos «{ruta.name}» no está guardado en UTF-8, y las tildes se "
                         "leerían mal. En el Bloc de notas: Archivo → Guardar como → Codificación: UTF-8.") from None
    try:
        contenido = json.loads(texto, object_pairs_hook=_Pares)
    except json.JSONDecodeError as e:
        raise ErrorDatos(f"El archivo de datos «{ruta.name}» está mal escrito (línea {e.lineno}, columna "
                         f"{e.colno}). Revisa que cada clave y cada texto vayan entre comillas dobles (\"), "
                         "que los datos estén separados por comas (sin coma después del último) y que todo "
                         "vaya entre llaves { }.") from None
    if not isinstance(contenido, _Pares):
        raise ErrorDatos(f"El archivo de datos «{ruta.name}» tiene que contener los datos entre llaves, "
                         "por ejemplo { \"titulo\": \"…\", \"autores\": [\"…\"] }.")

    datos = {}
    for original, valor in contenido:
        clave = _clave_conocida(original)
        if clave in datos:
            raise ErrorDatos(f"El dato «{clave}» aparece dos veces en «{ruta.name}». Deja solo uno.")
        datos[clave] = _valor(clave, valor)
    return datos


def combinar(ruta_json=None, opciones=None):
    """
    Junta los datos del archivo (si se indica) y los de las opciones sueltas
    ({clave: valor}; None si no se usó la opción), que mandan dato a dato.
    Devuelve DatosPortada, o None si no se ha dado ningún dato de portada.
    """
    opciones = {clave: valor for clave, valor in (opciones or {}).items() if valor is not None}
    if ruta_json is None and not opciones:
        return None
    datos = leer_json(ruta_json) if ruta_json is not None else {}
    for clave, valor in opciones.items():
        if clave not in DATOS:
            raise ValueError(f"dato de portada desconocido: {clave}")   # fallo de programación
        datos[clave] = _valor(clave, valor)
    return DatosPortada(**{_atributo(clave): valor for clave, valor in datos.items()})


def _enumerar(partes):
    return partes[0] if len(partes) == 1 else ", ".join(partes[:-1]) + " y " + partes[-1]


def avisar(datos, informe):
    """Anota en el informe los datos de portada que faltan (no se rellenan)."""
    faltan = datos.faltan()
    if not faltan:
        return
    que = _enumerar([DATOS[c][1] for c in faltan])
    opciones = _enumerar([DATOS[c][0] for c in faltan])
    informe.aviso(f"Faltan datos para la portada: {que}. No se han rellenado: añádelos en tu archivo de "
                  f"datos o con {'la opción' if len(faltan) == 1 else 'las opciones'} {opciones}.")
