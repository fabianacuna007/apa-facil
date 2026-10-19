"""
Paso 6.7: los apéndices.

Según la web oficial de APA (Appendices setup):
- cada apéndice empieza en una página nueva, al final del trabajo;
- arriba, el rótulo centrado y en negrita: «Apéndice» si hay uno solo, y
  «Apéndice A», «Apéndice B»… si hay varios;
- en la línea siguiente, su título, también centrado y en negrita (con las
  mayúsculas de APA, como los demás títulos: paso 4.6);
- su texto, a doble espacio (lo hace el paso 1.3);
- sus tablas y figuras se numeran con la letra del apéndice, sin punto
  («Tabla A1»): cambiar el número cambia el texto, así que solo se avisa.

El rótulo lleva el estilo «Título 1» de Word (centrado y en negrita, en el panel
de navegación); el título va centrado y en negrita, sin estilo de título, para
que el índice no lo repita. Si el rótulo y el título van en la misma línea
(«Anexo I. Guion de las entrevistas»), se separan y desaparece lo que había en
medio («. »), como en las leyendas de las tablas (6.1). El párrafo vacío que solo
llevaba el salto de página antes del apéndice se borra (como en el 6.4).

«Anexo» no se cambia por «Apéndice» por defecto (ROADMAP y decisión del
propietario: muchas universidades piden «Anexo»): se avisa, y se cambia con
--rotulo-apendice.
"""

from apa_facil.documento import LEYENDA_CON_NUMERO, ROTULO_DE_APENDICE, Estilos, texto_parrafo, w
from apa_facil.estructura import clasificar
from apa_facil.reglas import leyendas, referencias, titulos
from apa_facil.reglas.leyendas import _lista, _n, _plural, _xml

LETRAS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def apendices(clasificados):
    """Cada apéndice: (el Parrafo del rótulo, el Parrafo del título o None), en orden."""
    vivos = [c for c in clasificados if c.elemento.getparent() is not None and c.tipo != "vacio"]
    resultado = []
    for n, c in enumerate(vivos):
        if c.tipo != "anexo" or not ROTULO_DE_APENDICE.match(texto_parrafo(c.elemento).strip()):
            continue
        siguiente = vivos[n + 1] if n + 1 < len(vivos) else None
        titulo = siguiente if siguiente is not None and siguiente.tipo == "anexo" and \
            not ROTULO_DE_APENDICE.match(texto_parrafo(siguiente.elemento).strip()) else None
        resultado.append((c, titulo))
    return resultado


def rotulo_apa(n, cuantos):
    return "Apéndice" if cuantos == 1 else f"Apéndice {LETRAS[n]}"


def _formato_del_titulo(doc, p):
    """Centrado, en negrita y sin cursiva, sin sangría y pegado a lo que sigue; sin estilo de título."""
    antes = _xml(p)
    ppr = p.get_or_add_pPr()
    estilo = ppr.find(w("pStyle"))
    if estilo is not None:
        ppr.remove(estilo)
    for etiqueta in ("outlineLvl", "numPr", "pageBreakBefore"):   # (el salto, copiado del rótulo al separarlos)
        for el in ppr.findall(w(etiqueta)):
            ppr.remove(el)
    leyendas.formato_de_parrafo(p, True)
    ppr.get_or_add_jc().set(w("val"), "center")
    leyendas._formato_de_texto(doc, p, 0, True)
    return _xml(p) != antes


def aplicar(doc, opciones, informe, estructura=None):
    clasificados = estructura if estructura is not None else clasificar(doc)
    todos = apendices(clasificados)
    if not todos:
        return
    id_estilo = referencias._estilo_de_nivel_1(doc)
    separados, renombrados, sin_renombrar, sin_titulo, mayusculas, dudosas, paginas = [], [], [], [], [], [], []
    con_formato = 0

    for n, (rotulo, titulo) in enumerate(todos):
        p = rotulo.elemento
        hecho = referencias._pagina_nueva(doc, informe, rotulo, clasificados)
        if hecho:
            paginas.append(hecho)

        # El rótulo y el título, en dos líneas.
        texto = texto_parrafo(p).strip()
        apendice = ROTULO_DE_APENDICE.match(texto_parrafo(p))
        parrafo_del_titulo = titulo.elemento if titulo is not None else None
        if apendice and apendice["titulo"] and leyendas._se_puede_cambiar(p, apendice.end("separador")):
            parrafo_del_titulo = leyendas._dividir(p, apendice.end("rotulo"), apendice.end("separador"))
            if not opciones.get("titulos_tal_cual"):
                dudosas += titulos._poner_mayusculas(
                    parrafo_del_titulo, lambda x: titulos.mayusculas_apa(x, clasificados)) or []
            informe.cambio_previsto(rotulo.numero, "dividir_apendice", quitado=apendice["separador"],
                                    nuevo=texto_parrafo(parrafo_del_titulo))
            separados.append((texto, apendice["rotulo"], texto_parrafo(parrafo_del_titulo)))
        elif titulo is not None and not opciones.get("titulos_tal_cual"):
            antes = texto_parrafo(titulo.elemento)
            dudosas += titulos._poner_mayusculas(titulo.elemento, lambda x: titulos.mayusculas_apa(x, clasificados)) or []
            if texto_parrafo(titulo.elemento) != antes:
                informe.cambio_previsto(titulo.numero, "mayusculas_apa", nuevo=texto_parrafo(titulo.elemento))
                mayusculas.append((antes, texto_parrafo(titulo.elemento)))
        elif titulo is None and not (apendice and apendice["titulo"]):
            sin_titulo.append(texto)

        # La palabra de APA: «Apéndice» (o «Apéndice A»…), solo si se pide.
        actual, esperado = texto_parrafo(p).strip(), rotulo_apa(n, len(todos))
        if actual != esperado:
            if not opciones.get("rotulo_apendice"):
                sin_renombrar.append((actual, esperado))
            elif referencias._cambiar_texto(p, esperado):
                informe.cambio_previsto(rotulo.numero, "rotulo_de_apendice", nuevo=esperado)
                renombrados.append((actual, esperado))
            else:
                sin_renombrar.append((actual, esperado))

        # El formato: el rótulo, «Título 1»; el título, centrado y en negrita.
        antes = _xml(p)
        p.get_or_add_pPr().get_or_add_pStyle().set(w("val"), id_estilo)
        titulos._formato_a_mano(Estilos(doc), type("Rotulo", (), {"elemento": p, "nivel": 1})())
        cambiado = _xml(p) != antes
        if parrafo_del_titulo is not None:
            cambiado |= _formato_del_titulo(doc, parrafo_del_titulo)
        con_formato += cambiado

    _informar(informe, todos, separados, renombrados, mayusculas, dudosas, paginas, con_formato)
    _avisar(informe, clasificados, todos, sin_renombrar, sin_titulo)


def _informar(informe, todos, separados, renombrados, mayusculas, dudosas, paginas, con_formato):
    if paginas:
        informe.cambio(f"Apéndices: {_plural(len(paginas), 'apéndice')} empieza{_n(len(paginas))} ahora en una página "
                       "nueva, como pide APA"
                       + (" (se ha quitado el párrafo vacío que hacía el salto, para que no quede una línea en blanco "
                          "arriba)." if "borrado" in paginas else "."))
    if separados:
        texto, rotulo, titulo = separados[0]
        informe.cambio(f"Apéndices: el rótulo y el título van ahora en líneas distintas, como pide APA («{texto}» → "
                       f"«{rotulo}» y, debajo, «{titulo}»). El título lleva las mayúsculas de APA (si prefieres dejarlo "
                       "como estaba, usa la opción --titulos-tal-cual).")
    if mayusculas:
        informe.cambio(f"Apéndices: el título {_lista([d for _, d in mayusculas])} lleva ahora mayúscula inicial en las "
                       "palabras principales, como pide APA.")
    titulos.avisar_dudosas(informe, dudosas, "los títulos de los apéndices")
    if renombrados:
        informe.cambio("Apéndices: " + ", ".join(f"«{a}» → «{d}»" for a, d in renombrados)
                       + ", como pide APA (lo has pedido con la opción --rotulo-apendice). Si el texto los nombra "
                         "(«ver anexo I»), cámbialo tú.")
    if con_formato:
        informe.cambio(f"Apéndices: el rótulo y el título de {_plural(con_formato, 'apéndice')} van centrados y en "
                       "negrita, como pide APA; el rótulo, con el estilo «Título 1» de Word.")


def _avisar(informe, clasificados, todos, sin_renombrar, sin_titulo):
    if sin_renombrar:
        informe.aviso("APA llama «Apéndice» a los anexos («Apéndice A», «Apéndice B»… si hay varios): "
                      + ", ".join(f"«{a}» sería «{d}»" for a, d in sin_renombrar)
                      + ". Se ha dejado como estaba porque muchas universidades piden «Anexo»; si la tuya sigue APA, "
                        "usa la opción --rotulo-apendice.")
    if sin_titulo:
        informe.aviso(f"{_lista(sin_titulo)} no tiene{_n(len(sin_titulo))} título: APA lo pone en la línea siguiente "
                      "al rótulo, centrado y en negrita. Escríbelo tú.")
    leyendas_sin_letra = [texto_parrafo(c.elemento) for c in clasificados
                          if c.zona == "anexos" and c.tipo in ("leyenda_tabla", "leyenda_figura")
                          and c.elemento.getparent() is not None
                          and (lambda m: m is not None and m["numero"].split()[-1][0].isdigit())(
                              LEYENDA_CON_NUMERO.match(texto_parrafo(c.elemento)))]
    if leyendas_sin_letra:
        informe.aviso(f"En los apéndices, APA numera las tablas y figuras con la letra del apéndice y sin punto "
                      f"(«Tabla A1», «Figura B2»): {_lista(leyendas_sin_letra)}. Cámbialo tú.")
    primero = todos[0][0]
    despues = [c for c in clasificados[clasificados.index(primero) + 1:]
               if c.tipo == "titulo_referencias" and c.elemento.getparent() is not None]
    if despues:
        informe.aviso("Los apéndices van al final del trabajo, después de las referencias: mueve tú la lista de "
                      "referencias delante de los apéndices.")
