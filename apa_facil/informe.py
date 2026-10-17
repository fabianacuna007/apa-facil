"""
El informe de cada documento procesado. Distingue tres cosas:

- cambios: lo que la herramienta modificó;
- avisos: lo que el estudiante debe revisar a mano;
- errores: lo que no se pudo procesar.

Cada entrada recuerda qué regla la produjo. Las reglas no tienen que decirlo:
el pipeline anota en el informe qué regla se está aplicando antes de llamarla.

El informe se puede convertir en texto para la terminal (como_texto) y en JSON
para la web (a_dict, guardar_json). Los mensajes van en español y están
pensados para un estudiante.
"""

import json
from dataclasses import dataclass, field

from apa_facil.documento import ErrorDocumento

VERSION = 1   # versión del formato del JSON: se sube si cambia su estructura

# Nombre legible de cada regla (el módulo de apa_facil/reglas/ que la aplica).
NOMBRES = {
    "margenes": "Márgenes",
    "fuente": "Fuente",
    "interlineado": "Interlineado",
    "alineacion": "Alineación",
    "sangria": "Sangría",
    "paginacion": "Número de página",
}

GRUPOS = [  # (tipo de entrada, clave en el JSON, título en la terminal)
    ("cambio", "cambios", "Lo que hemos cambiado"),
    ("aviso", "avisos", "Lo que debes revisar tú"),
    ("error", "errores", "Lo que no se ha podido hacer"),
]


@dataclass
class Entrada:
    tipo: str        # "cambio", "aviso" o "error"
    regla: str       # p. ej. "paginacion"; None si no viene de una regla
    mensaje: str


@dataclass
class Informe:
    entradas: list = field(default_factory=list)
    regla: str = None        # la regla que se está aplicando (la pone el pipeline)
    archivo: str = None      # nombre del documento original
    copia: str = None        # nombre de la copia generada
    opciones: dict = field(default_factory=dict)
    fecha: str = None        # cuándo se procesó, p. ej. "2026-10-17T11:05:00"

    def cambio(self, mensaje):
        self.entradas.append(Entrada("cambio", self.regla, mensaje))

    def aviso(self, mensaje):
        self.entradas.append(Entrada("aviso", self.regla, mensaje))

    def error(self, mensaje):
        self.entradas.append(Entrada("error", self.regla, mensaje))

    def _mensajes(self, tipo):
        return [e.mensaje for e in self.entradas if e.tipo == tipo]

    @property
    def cambios(self):
        return self._mensajes("cambio")

    @property
    def avisos(self):
        return self._mensajes("aviso")

    @property
    def errores(self):
        return self._mensajes("error")

    # --- Para la web: JSON --------------------------------------------------------

    def a_dict(self):
        def entrada(e):
            return {"regla": e.regla, "nombre": NOMBRES.get(e.regla, "General"), "mensaje": e.mensaje}

        datos = {
            "version": VERSION,
            "archivo": self.archivo,
            "copia": self.copia,
            "fecha": self.fecha,
            "opciones": {"fuente": self.opciones.get("fuente")},
            "resumen": {clave: sum(1 for e in self.entradas if e.tipo == tipo) for tipo, clave, _ in GRUPOS},
        }
        for tipo, clave, _ in GRUPOS:
            datos[clave] = [entrada(e) for e in self.entradas if e.tipo == tipo]
        return datos

    def a_json(self):
        # ensure_ascii=False: las tildes se guardan tal cual ("á", no "á").
        return json.dumps(self.a_dict(), ensure_ascii=False, indent=2)

    def guardar_json(self, ruta):
        try:
            ruta.write_text(self.a_json() + "\n", encoding="utf-8")
        except PermissionError as e:
            raise ErrorDocumento(f"No se puede guardar el informe «{ruta.name}». Puede que otro programa "
                                 "lo tenga abierto: ciérralo y vuelve a intentarlo.") from e

    # --- Para la terminal ------------------------------------------------------------

    def como_texto(self):
        lineas = []
        for tipo, _, titulo in GRUPOS:
            mensajes = self._mensajes(tipo)
            if mensajes:
                lineas.append(f"   {titulo} ({len(mensajes)}):")
                lineas += [f"     · {m}" for m in mensajes]
        if not self.cambios:
            lineas.append("   No ha hecho falta cambiar nada: el formato ya cumplía APA.")
        return "\n".join(lineas)
