"""
Paso 2.2: la cadena completa con cada documento de pruebas/.

Cada documento se procesa una sola vez (en una carpeta temporal, nunca en
salida/) y sobre esa copia se comprueba:
1. que se procesa sin errores;
2. que el contenido es idéntico al del original (texto, tablas, imágenes,
   notas y secciones), leyendo la copia desde el disco, salvo los cambios
   previstos que las reglas anotan en el informe (paso 3.5);
3. que el verificador no encuentra ningún incumplimiento (fase 1 y estilos de título), y que se
   genera su informe en JSON (paso 2.3);
4. que el original no se ha modificado;
5. que las excepciones admitidas son exactamente las esperadas. Si un detector
   empezara a "perdonar" párrafos de más, el verificador seguiría diciendo
   "cumple": esta lista lo impide.
"""

import hashlib
import json
from pathlib import Path

import docx
import pytest

from apa_facil import estructura, pipeline
from apa_facil.documento import comparar, radiografia
from herramientas.verificar import ALINEACION, mostrar, verificar

CARPETA = Path(__file__).resolve().parent.parent / "pruebas"
PRUEBAS = sorted(CARPETA.glob("*.docx"))

# Excepciones admitidas que debe tener cada copia: (regla, número de párrafo).
# Al añadir un documento a pruebas/, revisa a mano sus excepciones y apúntalas aquí.
EXCEPCIONES = {
    "01_ensayo_basico": set(),
    "02_informe_tablas": set(),
    "03_trabajo_imagenes_listas": set(),
    "04_formato_caotico": set(),   # sus 3 párrafos con tabulador se limpian desde el paso 3.5
    "05_tfg_completo": {(ALINEACION, 8), (ALINEACION, 9), (ALINEACION, 10)},   # portada: Autora, Tutor, Convocatoria
}


def huella(ruta):
    return hashlib.sha256(Path(ruta).read_bytes()).hexdigest()


@pytest.fixture(scope="module", params=PRUEBAS, ids=lambda p: p.stem)
def procesado(request, tmp_path_factory):
    """Procesa un documento una vez y guarda lo necesario para las comprobaciones."""
    original = request.param
    antes = huella(original)
    salida, informe = pipeline.procesar(original, tmp_path_factory.mktemp(original.stem))
    return {"original": original, "huella_antes": antes, "salida": salida, "informe": informe,
            "verificacion": verificar(salida)}


def test_se_procesa_sin_errores(procesado):
    assert procesado["salida"].exists()
    assert procesado["informe"].errores == []


def test_el_contenido_es_identico(procesado):
    antes = radiografia(docx.Document(str(procesado["original"])))
    despues = radiografia(docx.Document(str(procesado["salida"])))
    diferencias = comparar(antes, despues, procesado["informe"].cambios_previstos)
    assert diferencias == [], "\n".join(diferencias)


def test_cumple_todas_las_reglas(procesado):
    resultado = procesado["verificacion"]
    assert resultado.cumple, mostrar(resultado)


def test_la_copia_lleva_su_estructura(procesado):
    """
    Desde el paso 4.1 cada título lleva su estilo de Word, así que la copia se
    clasifica igual que el original (el verificador solo tiene la copia).
    """
    def titulos(ruta):
        return [(c.tipo, c.nivel, c.texto) for c in estructura.clasificar(docx.Document(str(ruta)))
                if c.tipo in ("titulo", "titulo_referencias", "anexo")]
    assert titulos(procesado["salida"]) == titulos(procesado["original"])


def test_el_original_no_se_toca(procesado):
    assert huella(procesado["original"]) == procesado["huella_antes"]
    assert procesado["salida"].resolve().parent != CARPETA.resolve()


def test_excepciones_esperadas(procesado):
    nombre = procesado["original"].stem
    if nombre not in EXCEPCIONES:
        pytest.fail(f"«{nombre}» no está en EXCEPCIONES: revisa a mano sus excepciones con "
                    f"herramientas/verificar.py y apúntalas en {Path(__file__).name}.")
    halladas = {(h.regla, h.parrafo) for h in procesado["verificacion"].excepciones}
    assert halladas == EXCEPCIONES[nombre], mostrar(procesado["verificacion"])


def test_genera_su_informe(procesado):
    ruta = procesado["salida"].with_suffix(".json")
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    assert datos["archivo"] == procesado["original"].name
    assert datos["resumen"] == {"cambios": len(procesado["informe"].cambios),
                                "avisos": len(procesado["informe"].avisos), "errores": 0}
    assert datos["resumen"]["cambios"] > 0   # ningún original cumplía APA


def test_estan_todos_los_documentos_de_la_lista():
    # Si se borra o se renombra un documento, la lista de excepciones tiene que acompañarlo.
    assert {p.stem for p in PRUEBAS} == set(EXCEPCIONES)
