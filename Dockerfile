# Servidor de APA Fácil (paso 8.5).
#
#   docker build -t apa-facil .
#   docker run --rm -p 8000:8000 apa-facil
#
# Después: http://localhost:8000/docs. Render y Railway indican el puerto con la
# variable PORT; los límites se cambian con APA_TAMANO_MAXIMO_MB, APA_TIEMPO_MAXIMO
# y APA_SIMULTANEOS (ver servidor/app.py).
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app
COPY requirements-servidor.txt .
RUN pip install -r requirements-servidor.txt

# Solo el núcleo y el servidor: ni pruebas, ni documentos, ni herramientas.
COPY apa_facil/ apa_facil/
COPY servidor/ servidor/

# Sin permisos de administrador: solo puede escribir en su carpeta temporal.
RUN useradd --no-create-home --shell /usr/sbin/nologin apa
USER apa

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s \
    CMD python -c "import os, urllib.request; urllib.request.urlopen(f'http://localhost:{os.environ.get(\"PORT\", 8000)}/salud', timeout=4)"
CMD ["sh", "-c", "exec uvicorn servidor.app:app --host 0.0.0.0 --port ${PORT:-8000} --proxy-headers"]
