"""
Regla 1.4: texto alineado a la izquierda, sin justificar y sin guiones automáticos.

- Justificado (en todas sus variantes) → a la izquierda: estilos, cuerpo,
  tablas y notas al pie.
- Centrado: no se toca (títulos, portada, imágenes, leyendas, tablas). Los
  pasos posteriores deciden qué debe ir centrado.
- Alineado a la derecha → a la izquierda, salvo en lo que parece una portada
  (antes del primer salto de página o de sección y antes del primer párrafo
  largo): eso se deja y se avisa, porque la portada se trata en el paso 5.3.
- Los guiones automáticos se desactivan. Los guiones opcionales del texto solo
  se avisan: quitarlos modificaría el texto.
- Encabezados y pies no se tocan (paso 1.6).
"""

from apa_facil import documento
from apa_facil.documento import Estilos, texto_parrafo, w

JUSTIFICADOS = {"both", "distribute", "lowKashida", "mediumKashida", "highKashida", "thaiDistribute"}
A_LA_DERECHA = {"right", "end"}
PALABRAS_DE_PARRAFO_LARGO = 40   # el mismo umbral que APA usa para las citas en bloque
PARTES_QUE_NO_SE_TOCAN = ("/word/header", "/word/footer")
PARTES_DE_NOTAS = ("/word/footnotes.xml", "/word/endnotes.xml")


def _alinear(ppr, valor):
    ppr.get_or_add_jc().set(w("val"), valor)


def _jc_propio(ppr):
    jc = ppr.find(w("jc")) if ppr is not None else None
    return jc.get(w("val")) if jc is not None else None


def _salto_antes(estilos, p):
    """¿Empieza el párrafo en una página nueva ("salto de página anterior")?"""
    for capa in estilos._capas_parrafo(p):
        el = capa.find(w("pageBreakBefore")) if capa is not None else None
        if el is not None:
            return el.get(w("val")) not in ("0", "false", "off")
    return False


def _salto_en(p):
    """¿Hay un salto de página o de sección dentro del párrafo o justo después?"""
    if any(b.get(w("type")) == "page" for b in p.iter(w("br"))):
        return True
    if next(p.iter(w("lastRenderedPageBreak")), None) is not None:
        return True
    seccion = p.find(f"{w('pPr')}/{w('sectPr')}")
    if seccion is not None:
        tipo = seccion.find(w("type"))
        return tipo is None or tipo.get(w("val")) != "continuous"
    return False


def parece_portada(parrafos, estilos):
    """Los párrafos del principio que parecen una portada: hasta el primer salto o párrafo largo."""
    portada = []
    for p in parrafos:
        if _salto_antes(estilos, p) or len(texto_parrafo(p).split()) >= PALABRAS_DE_PARRAFO_LARGO:
            break
        portada.append(p)
        if _salto_en(p):
            break
    return portada


def _dentro_de(p, *etiquetas):
    return next(p.iterancestors(*etiquetas), None) is not None


def aplicar(doc, opciones, informe):
    estilos = Estilos(doc)
    cuerpo = [p for p in doc.element.body.iter(w("p")) if not _dentro_de(p, w("tc"), w("txbxContent"))]
    portada = set(parece_portada(cuerpo, estilos))

    # 1. Antes de cambiar nada: qué hay justificado y a la derecha.
    justificados = sum(1 for p in doc.element.body.iter(w("p"))
                       if (estilos.parrafo(p, "jc", "val") or "left") in JUSTIFICADOS)
    a_la_derecha = [p for p in cuerpo if (estilos.parrafo(p, "jc", "val") or "left") in A_LA_DERECHA]
    respetados = [p for p in a_la_derecha if p in portada]
    a_mover = [p for p in a_la_derecha if p not in portada]

    def contar_notas(nombre, raiz):
        nonlocal justificados
        if nombre.startswith(PARTES_DE_NOTAS):
            justificados += sum(1 for p in raiz.iter(w("p"))
                                if (estilos.parrafo(p, "jc", "val") or "left") in JUSTIFICADOS)
    documento.editar_partes_de_texto(doc, contar_notas)

    # 2. Encabezados y pies: se fija su alineación actual para que no cambie con los estilos.
    def fijar(nombre, raiz):
        if nombre.startswith(PARTES_QUE_NO_SE_TOCAN):
            for p in raiz.iter(w("p")):
                _alinear(p.get_or_add_pPr(), estilos.parrafo(p, "jc", "val") or "left")
    documento.editar_partes_de_texto(doc, fijar)

    # 3. Justificado → izquierda en los valores por defecto, los estilos y el formato directo.
    defecto = documento.parrafo_por_defecto(doc)
    if _jc_propio(defecto) in JUSTIFICADOS:
        _alinear(defecto, "left")
    for estilo in doc.styles.element.iter(w("style")):
        if _jc_propio(estilo.find(w("pPr"))) in JUSTIFICADOS:
            _alinear(estilo.find(w("pPr")), "left")

    def en_parte(nombre, raiz):
        if nombre.startswith(PARTES_QUE_NO_SE_TOCAN):
            return
        for p in raiz.iter(w("p")):
            if _jc_propio(p.find(w("pPr"))) in JUSTIFICADOS:
                _alinear(p.find(w("pPr")), "left")
    documento.editar_partes_de_texto(doc, en_parte)

    # 4. Derecha → izquierda, salvo en la portada.
    for p in a_mover:
        _alinear(p.get_or_add_pPr(), "left")

    # 5. Guiones.
    guiones_automaticos = _desactivar_guiones(doc)
    guiones_opcionales = sum(1 for _ in doc.element.body.iter(w("softHyphen")))

    _informar(informe, justificados, a_mover, respetados, guiones_automaticos, guiones_opcionales)


def _desactivar_guiones(doc):
    """Quita los guiones automáticos del documento. Devuelve si estaban activados."""
    ajustes = doc.settings.element
    guiones = ajustes.find(w("autoHyphenation"))
    if guiones is None:
        return False
    activados = guiones.get(w("val")) not in ("0", "false", "off")
    ajustes.remove(guiones)
    return activados


def _extracto(p, largo=35):
    texto = " ".join(texto_parrafo(p).split())
    return texto if len(texto) <= largo else texto[:largo] + "…"


def _parrafos(n):
    return "1 párrafo" if n == 1 else f"{n} párrafos"


def _informar(informe, justificados, a_mover, respetados, guiones_automaticos, guiones_opcionales):
    if justificados:
        verbo = "pasa a estar alineado" if justificados == 1 else "pasan a estar alineados"
        informe.cambio(f"Alineación: {_parrafos(justificados)} "
                       f"{'justificado' if justificados == 1 else 'justificados'} {verbo} a la izquierda "
                       "(APA no admite texto justificado).")
    if a_mover:
        verbo = "pasa" if len(a_mover) == 1 else "pasan"
        informe.cambio(f"Alineación: {_parrafos(len(a_mover))} alineado{'s' if len(a_mover) > 1 else ''} a la "
                       f"derecha {verbo} a la izquierda («{_extracto(a_mover[0])}»"
                       f"{' y otros' if len(a_mover) > 1 else ''}).")
    if respetados:
        varios = len(respetados) > 1
        informe.aviso(f"Al principio del documento hay {_parrafos(len(respetados))} alineado{'s' if varios else ''} "
                      f"a la derecha («{_extracto(respetados[0])}»). Parece{'n' if varios else ''} parte de la "
                      f"portada de tu universidad y se ha{'n' if varios else ''} dejado como estaba"
                      f"{'n' if varios else ''}. En APA, la portada va centrada.")
    if guiones_automaticos:
        informe.cambio("Se han desactivado los guiones automáticos a final de línea (APA no los admite).")
    if guiones_opcionales:
        cuantos = "1 guion opcional" if guiones_opcionales == 1 else f"{guiones_opcionales} guiones opcionales"
        informe.aviso(f"El texto tiene {cuantos}, que parten palabras a final de línea. APA no los admite: "
                      "búscalos en Word (Ctrl+B, «Más», «Especial», «Guion opcional») y quítalos.")
