"""
Paso 10.5 (adaptado: la herramienta no usa IA; decisión del propietario, 2026-10-20):
como mucho APA_LIMITE_DIARIO documentos al día por dirección IP (3 por defecto;
0 = sin límite). Más adelante, quien quiera más tendrá que pagar (fase 12).

- Solo cuentan los documentos que se procesan bien: el turno se reserva antes de
  procesar (para que dos peticiones a la vez no pasen del límite) y se devuelve
  si hay un error.
- El día es el del calendario en UTC (en España, cambia a la 1 o las 2 de la noche).
- Privacidad: no se guarda la IP, sino una huella (HMAC-SHA256 con una clave
  aleatoria que se crea cada día), y solo en memoria. Al cambiar de día se borra
  todo y se cambia la clave: con la huella de ayer no se puede saber nada.
- Va en la memoria del servidor: si se reinicia, el recuento vuelve a empezar, y
  con varias copias del servidor a la vez cada una contaría por su lado (para
  eso harían falta cuentas o una base de datos: fase 12).
"""

import hashlib
import hmac
import os
import secrets
import threading
from datetime import datetime, timezone

MAXIMO = int(os.environ.get("APA_LIMITE_DIARIO", 3))
# Paso 11.1: cuántos proxies de confianza hay delante del servidor (el del alojamiento).
# 0 (en local): la IP es la de la conexión. N: la IP es la N-ésima de X-Forwarded-For
# empezando por la derecha. Las de la izquierda las puede inventar el navegador, las
# de la derecha las añade cada proxy, así que no se pueden falsificar.
PROXIES = int(os.environ.get("APA_PROXIES", 0))


def ip_cliente(conexion, reenviado=None, proxies=PROXIES):
    """La IP de quien hace la petición: `conexion` es la IP que ve el servidor y
    `reenviado`, la cabecera X-Forwarded-For (o None)."""
    if not proxies:
        return conexion or ""
    saltos = [s.strip() for s in (reenviado or "").split(",") if s.strip()]
    if len(saltos) >= proxies:
        return saltos[-proxies]
    return conexion or ""  # falta la cabecera: no ha pasado por el proxy


class LimiteAlcanzado(Exception):
    pass


def mensaje(maximo):
    usados = "tu documento" if maximo == 1 else f"tus {maximo} documentos"
    return f"Has usado {usados} de hoy. Podrás procesar más mañana."


def _hoy():
    return datetime.now(timezone.utc).date()


class Limite:
    def __init__(self, maximo=MAXIMO, hoy=_hoy):
        self.maximo = maximo
        self._hoy = hoy
        self._cerrojo = threading.Lock()
        self._dia, self._clave, self._usados = None, None, {}

    def _al_dia(self):
        """Si ha cambiado el día, se olvida todo y se cambia la clave de las huellas."""
        if self._dia != self._hoy():
            self._dia, self._clave, self._usados = self._hoy(), secrets.token_bytes(32), {}

    def _huella(self, ip):
        return hmac.new(self._clave, (ip or "").encode(), hashlib.sha256).hexdigest()

    def reservar(self, ip):
        """Reserva un documento para esta IP y devuelve su huella; LimiteAlcanzado si ya no le quedan."""
        with self._cerrojo:
            self._al_dia()
            huella = self._huella(ip)
            if self.maximo and self._usados.get(huella, 0) >= self.maximo:
                raise LimiteAlcanzado(mensaje(self.maximo))
            self._usados[huella] = self._usados.get(huella, 0) + 1
            return huella, self._dia

    def devolver(self, reserva):
        """Un documento que no se ha podido procesar no cuenta."""
        huella, dia = reserva
        with self._cerrojo:
            if dia == self._dia and self._usados.get(huella):
                self._usados[huella] -= 1

    def restantes(self, ip):
        """Cuántos documentos le quedan hoy a esta IP (None si no hay límite)."""
        if not self.maximo:
            return None
        with self._cerrojo:
            self._al_dia()
            return max(0, self.maximo - self._usados.get(self._huella(ip), 0))
