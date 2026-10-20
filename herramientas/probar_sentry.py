"""
Paso 11.3: envía a Sentry un error de prueba, con la misma configuración y el
mismo filtro que el servidor (servidor/errores.py), para comprobar la clave.

    python herramientas/probar_sentry.py https://...@....ingest.de.sentry.io/...

En Sentry (Issues) aparece «ErrorDePrueba», con su sitio en el código y el
mensaje omitido.
"""

import sys
from pathlib import Path

import sentry_sdk

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from servidor import errores  # noqa: E402


class ErrorDePrueba(Exception):
    pass


def main():
    if len(sys.argv) != 2:
        sys.exit("Uso: python herramientas/probar_sentry.py <SENTRY_DSN>")
    errores.activar(sys.argv[1])
    try:
        raise ErrorDePrueba("este mensaje no debe llegar a Sentry")
    except ErrorDePrueba as e:
        errores.avisar(e)
    sentry_sdk.flush(10)
    print("Enviado. Búscalo en Sentry › Issues: «ErrorDePrueba», con el mensaje omitido.")


if __name__ == "__main__":
    main()
