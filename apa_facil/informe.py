"""
El informe de cada documento procesado. Distingue tres cosas:

- cambios: lo que la herramienta modificó;
- avisos: lo que el estudiante debe revisar a mano;
- errores: lo que no se pudo procesar.

Los mensajes van en español y están pensados para un estudiante.

De momento es una lista de mensajes por tipo. Se amplía en el paso 2.3
(informe en JSON).
"""

from dataclasses import dataclass, field


@dataclass
class Informe:
    cambios: list = field(default_factory=list)
    avisos: list = field(default_factory=list)
    errores: list = field(default_factory=list)

    def cambio(self, mensaje):
        self.cambios.append(mensaje)

    def aviso(self, mensaje):
        self.avisos.append(mensaje)

    def error(self, mensaje):
        self.errores.append(mensaje)
