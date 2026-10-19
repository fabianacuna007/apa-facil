"""
El informe de cada documento procesado. Distingue tres cosas:

- cambios: lo que la herramienta modificó;
- avisos: lo que el estudiante debe revisar a mano;
- errores: lo que no se pudo procesar.

Además guarda los cambios de contenido previstos (documento.CambioPrevisto):
los párrafos vacíos borrados, los tabuladores quitados del principio y, si se
pide, la numeración y las mayúsculas quitadas de los títulos. Con
ellos, la comprobación de integridad sabe qué diferencias son a propósito. No
van en el JSON: no hacen falta para el estudiante y así no se guarda su texto.

Cada entrada recuerda qué regla la produjo. Las reglas no tienen que decirlo:
el pipeline anota en el informe qué regla se está aplicando antes de llamarla.

El informe se puede convertir en texto para la terminal (como_texto) y en JSON
para la web (a_dict, guardar_json). Los mensajes van en español y están
pensados para un estudiante.
"""

import json
from dataclasses import dataclass, field

from apa_facil.documento import CambioPrevisto, ErrorDocumento

VERSION = 13   # versión del formato del JSON: se sube si cambia su estructura (2 y 3: opciones de los títulos;
              # 4: datos de la portada; 5: --titulo-tal-cual; 6: --titulos-tal-cual en vez de
              # --quitar-mayusculas; 7: --mantener-portada;
              # 8: --quitar-indice; 9: --tablas-tal-cual;
              # 10: --rotulo-tal-cual; 11: --orden-tal-cual;
              # 12: --rotulo-apendice; 13: los casos de los avisos de citas y referencias)

# Nombre legible de cada regla (el módulo de apa_facil/reglas/ que la aplica).
NOMBRES = {
    "portada": "Portada",
    "indice": "Índice",
    "resumen": "Resumen",
    "avisos": "Revisión del texto",
    "limpieza": "Limpieza de formato",
    "margenes": "Márgenes",
    "fuente": "Fuente",
    "interlineado": "Interlineado",
    "alineacion": "Alineación",
    "sangria": "Sangría",
    "titulos": "Títulos",
    "tablas": "Tablas",
    "figuras": "Figuras",
    "citas": "Citas en bloque",
    "referencias": "Referencias",
    "orden": "Orden de las referencias",
    "apendices": "Apéndices",
    "paginacion": "Número de página",
    "revision": "Citas y referencias",
}
MAXIMO_DE_CASOS_EN_LA_TERMINAL = 10   # el JSON los lleva todos

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
    # Paso 7.5: los sitios concretos de un aviso, cada uno un dict con "donde" (en
    # palabras: el apartado y el principio del párrafo), "extracto" (lo escrito),
    # "correccion" (la forma correcta, o None), "explicacion" (o None) y "parrafo"
    # (el número del párrafo en el original, o None en las notas).
    casos: list = None


@dataclass
class Informe:
    entradas: list = field(default_factory=list)
    regla: str = None        # la regla que se está aplicando (la pone el pipeline)
    archivo: str = None      # nombre del documento original
    copia: str = None        # nombre de la copia generada
    opciones: dict = field(default_factory=dict)
    fecha: str = None        # cuándo se procesó, p. ej. "2026-10-17T11:05:00"
    cambios_previstos: list = field(default_factory=list)   # CambioPrevisto

    def cambio(self, mensaje):
        self.entradas.append(Entrada("cambio", self.regla, mensaje))

    def aviso(self, mensaje, casos=None):
        self.entradas.append(Entrada("aviso", self.regla, mensaje, list(casos) if casos else None))

    def error(self, mensaje):
        self.entradas.append(Entrada("error", self.regla, mensaje))

    def cambio_previsto(self, parrafo, accion, quitado="", nuevo="", desde=(), parrafos=(), datos=None):
        """Anota un cambio de contenido hecho a propósito (ver documento.CambioPrevisto)."""
        self.cambios_previstos.append(CambioPrevisto(parrafo, accion, quitado, nuevo, tuple(desde),
                                                     tuple(parrafos), datos))

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
            datos = {"regla": e.regla, "nombre": NOMBRES.get(e.regla, "General"), "mensaje": e.mensaje}
            if e.casos:
                datos["casos"] = [dict(caso) for caso in e.casos]
            return datos

        datos = {
            "version": VERSION,
            "archivo": self.archivo,
            "copia": self.copia,
            "fecha": self.fecha,
            "opciones": {"fuente": self.opciones.get("fuente"),
                         "quitar_numeracion": bool(self.opciones.get("quitar_numeracion")),
                         "titulos_tal_cual": bool(self.opciones.get("titulos_tal_cual")),
                         "quitar_introduccion": bool(self.opciones.get("quitar_introduccion")),
                         "titulo_tal_cual": bool(self.opciones.get("titulo_tal_cual")),
                         "mantener_portada": bool(self.opciones.get("mantener_portada")),
                         "quitar_indice": bool(self.opciones.get("quitar_indice")),
                         "tablas_tal_cual": bool(self.opciones.get("tablas_tal_cual")),
                         "rotulo_tal_cual": bool(self.opciones.get("rotulo_tal_cual")),
                         "orden_tal_cual": bool(self.opciones.get("orden_tal_cual")),
                         "rotulo_apendice": bool(self.opciones.get("rotulo_apendice")),
                         "portada": self._portada()},
            "resumen": {clave: sum(1 for e in self.entradas if e.tipo == tipo) for tipo, clave, _ in GRUPOS},
        }
        for tipo, clave, _ in GRUPOS:
            datos[clave] = [entrada(e) for e in self.entradas if e.tipo == tipo]
        return datos

    def _portada(self):
        # Solo qué datos se han dado y cuáles faltan, no sus valores: son datos personales
        # (nombres) y la web ya los tiene.
        datos = self.opciones.get("portada")
        return {"dados": datos.dados(), "faltan": datos.faltan()} if datos else None

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
            entradas = [e for e in self.entradas if e.tipo == tipo]
            if entradas:
                lineas.append(f"   {titulo} ({len(entradas)}):")
                for e in entradas:
                    lineas.append(f"     · {e.mensaje}")
                    for caso in (e.casos or [])[:MAXIMO_DE_CASOS_EN_LA_TERMINAL]:
                        lineas.append(f"         – {caso_como_texto(caso)}")
                    if e.casos and len(e.casos) > MAXIMO_DE_CASOS_EN_LA_TERMINAL:
                        lineas.append(f"         – y {len(e.casos) - MAXIMO_DE_CASOS_EN_LA_TERMINAL} más "
                                      "(están todos en el informe completo).")
        if not self.cambios:
            lineas.append("   No ha hecho falta cambiar nada: el formato ya cumplía APA.")
        return "\n".join(lineas)


def caso_como_texto(caso):
    """«En «Introducción», en el párrafo que empieza «…»: «(Martínez 2019)» → «(Martínez, 2019)» (…)»."""
    texto = f"{caso['donde']}: «{caso['extracto']}»"
    if caso.get("correccion"):
        texto += f" → «{caso['correccion']}»"
    if caso.get("explicacion"):
        texto += f" ({caso['explicacion']})"
    return texto
