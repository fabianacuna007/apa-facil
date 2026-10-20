"""
Paso 11.3: avisos de los errores inesperados del servidor con Sentry, sin el
contenido de los documentos.

Solo se activa si existe la variable SENTRY_DSN (la clave del proyecto de Sentry,
en el panel de Render). Sin ella no se envía nada a ningún sitio.

De cada error se envía solo:
  - su tipo (KeyError, ValueError…);
  - por dónde ha pasado en el código: archivo, línea, función y las líneas de
    código de alrededor (código del programa, no del documento);
  - la versión del servidor (el commit) y la de Python.
Nunca se envía el mensaje del error (podría llevar texto del documento), las
variables, la petición (cabeceras, IP, formulario, archivo) ni el nombre del
archivo. `limpiar` lo garantiza con una lista blanca: copia solo los campos
permitidos, así que lo que Sentry añada en otra versión tampoco sale.

Los errores del núcleo ocurren en otro proceso (servidor/trabajo.py), que solo
devuelve el tipo y los sitios del código (`trabajo.marcos`); `avisar_fallo_nucleo`
los envía con el mismo formato que un error del propio servidor.
"""

import os

import sentry_sdk

OMITIDO = "(mensaje omitido: podría llevar texto del documento)"
_EVENTO = ("event_id", "timestamp", "platform", "level", "release", "environment", "sdk")
_MARCO = ("filename", "module", "function", "lineno", "in_app", "pre_context", "context_line", "post_context")


def activar(dsn=None, transporte=None):
    """Activa Sentry si hay clave (`dsn` o SENTRY_DSN). Devuelve si ha quedado activo."""
    dsn = os.environ.get("SENTRY_DSN", "") if dsn is None else dsn
    if not dsn:
        return False
    sentry_sdk.init(
        dsn=dsn,
        transport=transporte,
        # Nada automático: ni registros, ni peticiones, ni rendimiento. Solo lo que se envía a mano.
        default_integrations=False,
        auto_enabling_integrations=False,
        send_default_pii=False,
        include_local_variables=False,
        max_request_body_size="never",
        max_breadcrumbs=0,
        before_send=limpiar,
        release=os.environ.get("RENDER_GIT_COMMIT"),   # Render la pone sola
        environment=os.environ.get("APA_ENTORNO", "produccion"),
        server_name="apa-facil-api",
    )
    return True


def limpiar(evento, pista=None):
    """Deja en el evento solo los campos permitidos (lista blanca)."""
    limpio = {clave: evento[clave] for clave in _EVENTO if clave in evento}
    limpio["server_name"] = "apa-facil-api"
    runtime = (evento.get("contexts") or {}).get("runtime")
    if runtime:
        limpio["contexts"] = {"runtime": {k: runtime[k] for k in ("name", "version") if k in runtime}}
    valores = []
    for valor in (evento.get("exception") or {}).get("values") or []:
        marcos = [{k: m[k] for k in _MARCO if k in m}
                  for m in (valor.get("stacktrace") or {}).get("frames") or []]
        valores.append({"type": valor.get("type", "Error"), "value": OMITIDO,
                        "mechanism": {"type": (valor.get("mechanism") or {}).get("type", "generic"),
                                      "handled": True},
                        "stacktrace": {"frames": marcos}})
    if not valores:   # solo se envían errores; cualquier otra cosa se descarta
        return None
    limpio["exception"] = {"values": valores}
    return limpio


def avisar(excepcion):
    """Envía un error del servidor (no hace nada si Sentry no está activo)."""
    sentry_sdk.capture_exception(excepcion)


def avisar_fallo_nucleo(tipo, marcos):
    """Envía un error del núcleo a partir de su tipo y sus sitios [(archivo, línea, función)]."""
    sentry_sdk.capture_event({
        "level": "error",
        "exception": {"values": [{
            "type": tipo,
            "mechanism": {"type": "nucleo"},
            "stacktrace": {"frames": [{"filename": archivo, "lineno": linea, "function": funcion,
                                       "in_app": archivo.startswith(("apa_facil/", "servidor/"))}
                                      for archivo, linea, funcion in marcos]},
        }]},
    })
