"""
Paso 10.5 (adaptado, sin IA): como mucho 3 documentos al día por dirección IP,
contando solo los que se procesan bien, y lo que cuesta cada documento en el registro.
"""

import logging
from datetime import date
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from servidor import app as servidor_app
from servidor import limites
from servidor.limites import Limite, LimiteAlcanzado, mensaje
from servidor.trabajo import memoria_maxima_mb

RAIZ = Path(__file__).resolve().parent.parent
DOCUMENTO = RAIZ / "pruebas" / "01_ensayo_basico.docx"
TEXTO = "El trabajo infantil\nIntroducción\nUn párrafo del trabajo, con su punto final."
WEB = {"Origin": "http://localhost:3000"}


# --- El contador ----------------------------------------------------------------------

def test_tres_al_dia_por_ip():
    limite = Limite(maximo=3)
    for _ in range(3):
        limite.reservar("1.2.3.4")
    with pytest.raises(LimiteAlcanzado, match="Has usado tus 3 documentos de hoy. Podrás procesar más mañana."):
        limite.reservar("1.2.3.4")
    assert limite.restantes("1.2.3.4") == 0
    assert limite.restantes("5.6.7.8") == 3   # cada IP, lo suyo
    limite.reservar("5.6.7.8")


def test_un_documento_que_falla_no_cuenta():
    limite = Limite(maximo=3)
    reserva = limite.reservar("1.2.3.4")
    limite.devolver(reserva)
    assert limite.restantes("1.2.3.4") == 3


def test_al_dia_siguiente_vuelve_a_empezar():
    dia = [date(2026, 10, 20)]
    limite = Limite(maximo=3, hoy=lambda: dia[0])
    reserva = limite.reservar("1.2.3.4")
    huella = reserva[0]
    for _ in range(2):
        limite.reservar("1.2.3.4")
    dia[0] = date(2026, 10, 21)
    assert limite.restantes("1.2.3.4") == 3
    limite.devolver(reserva)   # una reserva de ayer no se devuelve hoy
    assert limite.restantes("1.2.3.4") == 3
    assert limite.reservar("1.2.3.4")[0] != huella   # la clave de las huellas cambia cada día


def test_no_se_guarda_la_ip():
    limite = Limite(maximo=3)
    huella, _ = limite.reservar("1.2.3.4")
    assert "1.2.3.4" not in huella and "1.2.3.4" not in repr(limite.__dict__)


def test_sin_limite():
    limite = Limite(maximo=0)
    for _ in range(10):
        limite.reservar("1.2.3.4")
    assert limite.restantes("1.2.3.4") is None


def test_mensaje_en_singular():
    assert mensaje(1) == "Has usado tu documento de hoy. Podrás procesar más mañana."


# --- En el servidor -------------------------------------------------------------------

@pytest.fixture
def cliente(monkeypatch):
    monkeypatch.setattr(servidor_app, "LIMITE", Limite(maximo=3))
    return TestClient(servidor_app.app)


def subir(cliente, ruta=DOCUMENTO, nombre=None):
    with open(ruta, "rb") as f:
        return cliente.post("/procesar", headers=WEB, files={"archivo": (nombre or ruta.name, f)})


def test_el_cuarto_documento_del_dia_se_rechaza(cliente):
    assert cliente.get("/limite").json() == {"maximo": 3, "restantes": 3}
    restantes = []
    for _ in range(2):
        respuesta = subir(cliente)
        assert respuesta.status_code == 200
        restantes.append(respuesta.json()["limite"]["restantes"])
    texto = cliente.post("/procesar-texto", data={"texto": TEXTO})   # el texto pegado cuenta igual
    assert texto.status_code == 200
    restantes.append(texto.json()["limite"]["restantes"])
    assert restantes == [2, 1, 0]
    cuarto = subir(cliente)
    assert cuarto.status_code == 429
    assert cuarto.json() == {"error": "Has usado tus 3 documentos de hoy. Podrás procesar más mañana."}
    assert cuarto.headers["access-control-allow-origin"] == "http://localhost:3000"   # la web lo puede leer
    assert cliente.get("/limite").json() == {"maximo": 3, "restantes": 0}


def test_los_documentos_con_error_no_cuentan(cliente, tmp_path):
    vacio = tmp_path / "trabajo.docx"
    vacio.write_bytes(b"")
    falso = tmp_path / "falso.docx"
    falso.write_bytes(b"PK no es un zip")
    assert subir(cliente, vacio).status_code == 422
    assert subir(cliente, falso).status_code == 422
    assert cliente.post("/procesar-texto", data={"texto": " "}).status_code == 422
    assert cliente.get("/limite").json()["restantes"] == 3


def test_tiempo_agotado_no_cuenta(cliente, monkeypatch):
    monkeypatch.setattr(servidor_app, "TIEMPO_MAXIMO", 0.001)
    assert subir(cliente).status_code == 504
    assert cliente.get("/limite").json()["restantes"] == 3


# --- Lo que cuesta cada documento -----------------------------------------------------

def test_el_registro_dice_lo_que_ha_costado_sin_el_nombre(cliente, caplog):
    with caplog.at_level(logging.INFO, logger="apa_facil.servidor"):
        assert subir(cliente, nombre="Mi trabajo secreto.docx").status_code == 200
    (linea,) = [r.getMessage() for r in caplog.records if r.name == "apa_facil.servidor"]
    assert linea.startswith("Documento procesado en ") and " s del núcleo; memoria máxima del proceso: " in linea
    assert "secreto" not in linea and "Mi trabajo" not in linea


def test_memoria_maxima():
    memoria = memoria_maxima_mb()
    assert memoria is None or 1 < memoria < 10_000


# --- La IP real detrás del proxy del alojamiento (paso 11.1) ---------------------------

def test_sin_proxies_cuenta_la_ip_de_la_conexion():
    assert limites.ip_cliente("10.0.0.1", "1.2.3.4", proxies=0) == "10.0.0.1"


def test_con_proxy_cuenta_la_ip_que_anade_el_proxy_y_no_la_inventada():
    # El navegador manda «X-Forwarded-For: 6.6.6.6» y el proxy añade la IP real.
    assert limites.ip_cliente("10.0.0.1", "6.6.6.6, 81.2.3.4", proxies=1) == "81.2.3.4"
    assert limites.ip_cliente("10.0.0.1", "6.6.6.6, 81.2.3.4, 172.16.0.9", proxies=2) == "81.2.3.4"


def test_sin_cabecera_cuenta_la_ip_de_la_conexion():
    assert limites.ip_cliente("10.0.0.1", None, proxies=1) == "10.0.0.1"
    assert limites.ip_cliente("10.0.0.1", "81.2.3.4", proxies=2) == "10.0.0.1"
