"""
Regla 1.1: márgenes de 2,54 cm en los cuatro lados, en todas las secciones.

- No se toca el tamaño del papel ni la orientación: una sección horizontal
  sigue siendo horizontal.
- El margen de encuadernación se pone a 0, porque Word lo suma a uno de los
  márgenes y entonces ese margen mediría más de 2,54 cm.
- La distancia del encabezado y del pie al borde no se toca (es del paso 1.6).

Al cambiar los márgenes cambia el espacio para el texto. Si una tabla o una
imagen ya no cabe, se avisa, pero no se cambia su tamaño.
"""

from docx.shared import Twips

from apa_facil import documento
from apa_facil.documento import cm, texto_parrafo, w

MARGEN = 1440        # 2,54 cm en twips (exacto: 1 pulgada)
TOLERANCIA = 15      # menos de 0,03 cm: no se nota en la página


def _twips(longitud):
    return None if longitud is None else longitud.twips


def _describir(margenes):
    superior, inferior, izquierdo, derecho = margenes
    if None in margenes:
        return "unos márgenes sin indicar"
    if len(set(margenes)) == 1:
        return f"{cm(superior)} cm en todos"
    return (f"superior {cm(superior)} · inferior {cm(inferior)} · "
            f"izquierdo {cm(izquierdo)} · derecho {cm(derecho)} cm")


def _lista(numeros):
    """[1, 2, 3] → "1, 2 y 3"."""
    textos = [str(n) for n in numeros]
    return textos[0] if len(textos) == 1 else ", ".join(textos[:-1]) + " y " + textos[-1]


def _es_horizontal(seccion):
    ancho, alto = seccion.page_width, seccion.page_height
    return ancho is not None and alto is not None and ancho > alto


def aplicar(doc, opciones, informe, estructura=None):
    cambios = {}   # mensaje → secciones a las que se aplica (para no repetirlo)
    for numero, seccion in enumerate(doc.sections, start=1):
        antes = tuple(_twips(m) for m in (seccion.top_margin, seccion.bottom_margin,
                                          seccion.left_margin, seccion.right_margin))
        encuadernacion = _twips(seccion.gutter) or 0
        if antes == (MARGEN,) * 4 and encuadernacion == 0:
            continue

        seccion.top_margin = seccion.bottom_margin = Twips(MARGEN)
        seccion.left_margin = seccion.right_margin = Twips(MARGEN)
        seccion.gutter = Twips(0)

        if antes == (MARGEN,) * 4:
            mensaje = "márgenes de 2,54 cm en los cuatro lados (ya lo eran)"
        else:
            mensaje = f"márgenes cambiados a 2,54 cm en los cuatro lados (antes: {_describir(antes)})"
        if encuadernacion:
            mensaje += (f". Se ha quitado el margen de encuadernación de {cm(encuadernacion)} cm, "
                        "que hacía más ancho uno de los márgenes")
        clave = (mensaje, _es_horizontal(seccion))
        cambios.setdefault(clave, []).append(numero)

    for (mensaje, horizontal), numeros in cambios.items():
        if len(doc.sections) == 1:
            informe.cambio(f"{mensaje[0].upper()}{mensaje[1:]}.")
            continue
        donde = ("Sección " if len(numeros) == 1 else "Secciones ") + _lista(numeros)
        if horizontal:
            donde += " (página horizontal, se mantiene así)"
        informe.cambio(f"{donde}: {mensaje}.")

    _avisar_de_lo_que_no_cabe(doc, informe)


def _avisar_de_lo_que_no_cabe(doc, informe):
    """Tablas e imágenes más anchas que el espacio que dejan los márgenes."""
    contador = {}
    for seccion, bloques in zip(doc.sections, documento.bloques_por_seccion(doc)):
        if seccion.page_width is None:
            continue
        disponible = seccion.page_width.twips - seccion.left_margin.twips - seccion.right_margin.twips

        elementos = [("tabla", documento.ancho_tabla(t), t) for t in documento.tablas_de(bloques)]
        elementos += [(tipo, ancho, None) for tipo, ancho in documento.dibujos_de(bloques)]
        for tipo, ancho, tabla in elementos:
            contador[tipo] = contador.get(tipo, 0) + 1
            if ancho is None or ancho <= disponible + TOLERANCIA:
                continue
            nombre = f"{tipo} {contador[tipo]}"
            if tabla is not None:
                nombre += _primera_fila(tabla)
                remedio = ("APA recomienda ponerla en una página horizontal: en Word, selecciónala entera y, en "
                           "la pestaña «Disposición», abre «Márgenes» → «Márgenes personalizados», elige "
                           "«Horizontal» y, en «Aplicar a», «Texto seleccionado». Si cabe estrechándola, también "
                           "puedes hacer clic en la tabla y usar «Autoajustar» → «Autoajustar a la ventana».")
            else:
                remedio = "En Word, redúcela arrastrando una de sus esquinas."
            informe.aviso(f"{nombre[0].upper()}{nombre[1:]} mide {cm(ancho)} cm de ancho y entre los "
                          f"márgenes solo caben {cm(disponible)} cm: se saldrá por la derecha. {remedio}")


def _primera_fila(tabla):
    """Un trozo del texto de la primera fila, para que el estudiante sepa qué tabla es."""
    fila = tabla.find(w("tr"))
    if fila is None:
        return ""
    texto = " | ".join(t for t in (texto_parrafo(p).strip() for p in fila.iter(w("p"))) if t)
    if not texto:
        return ""
    return f" («{texto[:40]}…»)" if len(texto) > 40 else f" («{texto}»)"
