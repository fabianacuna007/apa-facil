"""
Paso 6.2: las figuras.

Según la web oficial de APA (Figure setup y Sample figures), una figura es todo
lo que no es una tabla: gráficos, esquemas, fotografías, mapas… Tiene la misma
disposición que una tabla: encima, «Figura 1» en negrita y, en la línea
siguiente, el título en cursiva; debajo, la nota con «Nota.» en cursiva. Lo hace
leyendas.py, como para las tablas (con lo que cambia el texto y las decisiones
del propietario: «Gráfico 1» o «Imagen 2» pasan a «Figura 1» y «Figura 2»).

Aquí, una figura es un párrafo con solo una imagen, fuera de la portada y del
índice. No lo son las imágenes dentro de una tabla ni las que van en medio de una
frase. APA no dice cómo alinear la imagen: se deja como está. Si tiene una nota
debajo, la imagen se pega a ella («Conservar con el siguiente»).

No se comprueba el texto de dentro de la imagen (APA: sin serifa, de 8 a 14 pt):
es parte del dibujo.
"""

from apa_facil.documento import w
from apa_facil.estructura import clasificar
from apa_facil.reglas import leyendas
from apa_facil.reglas.leyendas import FIGURA

NUMERO_O_LETRA = FIGURA.numero_o_letra


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    imagenes = {c.elemento: c for c in clasificados if c.tipo == "imagen" and c.elemento.getparent() is not None}
    figuras = leyendas.buscar(doc, clasificados, FIGURA, lambda b: b in imagenes, lambda b: imagenes[b].numero)
    if not figuras:
        return
    leyendas.tratar(doc, opciones, informe, clasificados, figuras, FIGURA)
    for f in figuras:
        if f.notas and f.elemento.tag == w("p"):
            leyendas.formato_de_parrafo(f.elemento, True, sin_sangria=False)   # la imagen, pegada a su nota
