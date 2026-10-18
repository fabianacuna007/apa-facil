"""
Paso 5.5: el resumen y las palabras clave.

Según la web oficial de APA (Paragraph alignment and indentation y la guía
Abstract and Keywords Guide):
- el rótulo «Resumen», en la primera línea de la segunda página, centrado y en
  negrita (lo hacen los pasos 4.1, 4.2 y 4.6, y el 5.3 lo deja en página nueva);
- el texto, sin sangría (paso 3.4), en un solo párrafo de 250 palabras como
  mucho: si no, solo se avisa;
- las palabras clave, en la línea siguiente, con sangría (paso 3.4): el rótulo
  «Palabras clave:» en cursiva y sin negrita, y las palabras sin cursiva ni
  negrita, en minúscula salvo los nombres propios (si no, solo se avisa: pasarlas
  a minúscula podría estropear un nombre propio), separadas por comas, de tres a
  cinco y sin punto al final. El punto final se quita (decisión del propietario),
  salvo que sea el de una abreviatura («EE. UU.», «etc.»);
- el texto del trabajo empieza en una página nueva («start each section on a new
  page»).

El ROADMAP decía que las palabras clave iban en cursiva; según APA, solo su rótulo.
"""

import copy
import re

from apa_facil.documento import PALABRAS_CLAVE, Estilos, texto_parrafo, w
from apa_facil.estructura import clasificar
from apa_facil.reglas import titulos
from apa_facil.reglas.fuente import _runs_propios

MAXIMO_DE_PALABRAS = 250
PALABRAS_CLAVE_MINIMAS, PALABRAS_CLAVE_MAXIMAS = 3, 5
ABREVIATURA_FINAL = re.compile(r"(\b\w{1,4}\.\s*){2,}$|\betc\.$|\b[A-ZÁÉÍÓÚÑ]\.$")


def _extracto(texto, largo=40):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo - 1] + "…"


def _activado(estilos, r, p, etiqueta):
    """¿Se ve activado en el fragmento un ajuste como la negrita ("b") o la cursiva ("i")?"""
    for capa in estilos._capas_texto(r, p):
        el = capa.find(w(etiqueta)) if capa is not None else None
        if el is not None:
            return el.get(w("val")) not in ("0", "false", "off")
    return False


def _poner(r, etiqueta, activado):
    """Pone en el formato directo del fragmento un ajuste (y su versión para otros alfabetos) encendido o apagado."""
    rpr = r.get_or_add_rPr()
    for nombre in (etiqueta, etiqueta + "Cs"):
        el = getattr(rpr, f"get_or_add_{nombre}")()
        if activado:
            el.attrib.pop(w("val"), None)
        else:
            el.set(w("val"), "0")


def _partir(r, posicion):
    """Parte el fragmento en dos en `posicion` (contada en sus letras), sin cambiar el texto. Devuelve el segundo."""
    textos = list(r.iter(w("t")))
    hasta = 0
    for t in textos:
        largo = len(t.text or "")
        if hasta + largo > posicion:
            corte = posicion - hasta
            segundo = copy.deepcopy(r)
            segundos_t = list(segundo.iter(w("t")))
            indice = textos.index(t)
            for viejo in textos[indice + 1:]:          # en el primero, solo hasta el corte
                viejo.getparent().remove(viejo)
            for nuevo in segundos_t[:indice]:          # en el segundo, desde el corte
                nuevo.getparent().remove(nuevo)
            primero_texto, segundo_texto = (t.text or "")[:corte], (t.text or "")[corte:]
            t.text, segundos_t[indice].text = primero_texto, segundo_texto
            for pieza in (t, segundos_t[indice]):
                pieza.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.addnext(segundo)
            return segundo
        hasta += largo
    return None


def _formato_de_palabras_clave(doc, p, final_del_rotulo):
    """Rótulo en cursiva y sin negrita; las palabras, sin cursiva ni negrita. Devuelve si ha cambiado algo."""
    estilos = Estilos(doc)
    desde = 0
    for r in _runs_propios(p):   # primero, que el rótulo acabe donde acaba un fragmento
        largo = sum(len(t.text or "") for t in r.iter(w("t")))
        if desde < final_del_rotulo < desde + largo:
            _partir(r, final_del_rotulo - desde)
            break
        desde += largo
    desde, cambiado = 0, False
    for r in _runs_propios(p):
        largo = sum(len(t.text or "") for t in r.iter(w("t")))
        if not largo:
            continue
        en_el_rotulo = desde < final_del_rotulo
        for etiqueta, quiere in (("i", en_el_rotulo), ("b", False)):
            if _activado(estilos, r, p, etiqueta) != quiere:
                _poner(r, etiqueta, quiere)
                cambiado = True
        desde += largo
    return cambiado


def _quitar_punto_final(p):
    """Quita el punto del final del párrafo (del último fragmento con texto)."""
    for t in reversed([t for r in _runs_propios(p) for t in r.iter(w("t"))]):
        if t.text:
            if t.text.endswith("."):
                t.text = t.text[:-1]
                return True
            return False
    return False


def _palabras_clave(doc, informe, c, clasificados):
    texto = texto_parrafo(c.elemento)
    rotulo = PALABRAS_CLAVE.match(texto)
    lista = [palabra.strip() for palabra in texto[rotulo.end():].split(",") if palabra.strip()]
    if _formato_de_palabras_clave(doc, c.elemento, rotulo.end(2) if rotulo.group(2) else rotulo.end(1)):
        informe.cambio("Resumen: el rótulo «Palabras clave:» va ahora en cursiva y sin negrita, y las palabras clave "
                       "sin cursiva ni negrita, como pide APA.")
    if not rotulo.group(2):
        informe.aviso("El rótulo de las palabras clave no lleva dos puntos: APA lo escribe «Palabras clave:». "
                      "Añádelos tú.")
    if texto.endswith("."):
        if ABREVIATURA_FINAL.search(texto[rotulo.end():]):
            informe.aviso("Las palabras clave terminan en un punto que parece de una abreviatura "
                          f"(«{_extracto(lista[-1] if lista else texto, 20)}»): APA no pone punto al final, pero se "
                          "ha dejado. Revísalo tú.")
        elif _quitar_punto_final(c.elemento):
            informe.cambio_previsto(c.numero, "quitar_punto_final", quitado=".")
            informe.cambio("Resumen: se ha quitado el punto del final de las palabras clave, como pide APA.")
    if lista and not PALABRAS_CLAVE_MINIMAS <= len(lista) <= PALABRAS_CLAVE_MAXIMAS:
        informe.aviso(f"Hay {len(lista)} palabra{'s' if len(lista) > 1 else ''} clave: APA recomienda de "
                      f"{PALABRAS_CLAVE_MINIMAS} a {PALABRAS_CLAVE_MAXIMAS}.")
    formas = titulos._formas_del_texto(clasificados)
    con_mayuscula = [palabra for palabra in lista if palabra[:1].isupper() and not palabra.isupper()
                     and palabra.split()[0].lower() not in formas]
    if con_mayuscula:
        informe.aviso(f"Algunas palabras clave empiezan con mayúscula ({', '.join(f'«{x}»' for x in con_mayuscula[:3])}"
                      f"{'…' if len(con_mayuscula) > 3 else ''}). APA las escribe en minúscula, salvo los nombres "
                      "propios: cámbialas tú si no lo son.")


def _empieza_pagina_despues(doc, informe, ultimo):
    """El texto del trabajo empieza en una página nueva después del resumen."""
    from apa_facil.reglas import portada   # aquí: portada usa titulos, como este módulo
    siguiente = next((s for s in ultimo.itersiblings() if s.tag in (w("p"), w("tbl"), w("sdt"))), None)
    if portada._termina_en_salto(ultimo) or siguiente is None:
        return
    if siguiente.tag != w("p"):
        informe.aviso("Después del resumen hay una tabla o un índice, así que no se ha podido hacer que lo que sigue "
                      "empiece en una página nueva: pon tú un salto de página al final del resumen.")
        return
    if not titulos._salto_de_pagina_antes(Estilos(doc), siguiente):
        portada._empezar_pagina(siguiente, ultimo)
        informe.cambio("Resumen: lo que sigue al resumen empieza ahora en una página nueva, como pide APA.")


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    resumen = [c for c in clasificados if c.zona == "resumen" and c.elemento.getparent() is not None]
    if not resumen:
        return
    clave = next((c for c in resumen if PALABRAS_CLAVE.match(c.texto)), None)
    texto = [c for c in resumen if c.tipo == "cuerpo" and c is not clave]
    palabras = sum(len(c.texto.split()) for c in texto)
    if len(texto) > 1:
        informe.aviso(f"El resumen tiene {len(texto)} párrafos: APA lo escribe en uno solo. Únelos tú.")
    if palabras > MAXIMO_DE_PALABRAS:
        informe.aviso(f"El resumen tiene {palabras} palabras: APA pone un máximo de {MAXIMO_DE_PALABRAS}, salvo que "
                      "tu profesor pida otra cosa.")
    if clave is not None:
        _palabras_clave(doc, informe, clave, clasificados)
    _empieza_pagina_despues(doc, informe, resumen[-1].elemento)
