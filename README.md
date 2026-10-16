# APA Fácil

Aplica automáticamente el formato APA 7 (versión de estudiante) a trabajos académicos en Word (.docx). Nunca cambia el texto: solo el formato, y siempre en una copia nueva.

El plan de desarrollo está en [ROADMAP.md](ROADMAP.md).

## Instalación (Windows)

Hace falta Python 3.11 o más reciente. Desde la carpeta del proyecto:

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

El primer comando crea un entorno aislado para el proyecto, el segundo lo activa y el tercero instala lo que necesita. A partir de ahí, basta con activar el entorno (segundo comando) cada vez que se abra una terminal nueva.

## Uso

El núcleo (`apa_facil/`) está en construcción. De momento abre el documento y guarda una copia **sin cambios** en `salida/`:

```
python -m apa_facil pruebas\01_ensayo_basico.docx
```

Para aplicar ya el formato APA, se usa el prototipo:

```
python apa.py pruebas\01_ensayo_basico.docx
python apa.py pruebas\01_ensayo_basico.docx --fuente Arial
```

El resultado se guarda en `salida/` con el nombre `nombre_APA.docx`.

## Pruebas automáticas

```
pytest
```
