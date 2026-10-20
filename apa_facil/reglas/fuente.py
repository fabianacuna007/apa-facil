"""
Regla 1.2: una sola fuente para todo el documento, con su tamaño.

Cómo se elige (decisión del propietario del proyecto):
- se mira qué fuente tiene más letras en el cuerpo del documento;
- si es una de las recomendadas por APA, se mantiene y solo se ajusta el tamaño;
- si no, se cambia por la recomendada más parecida;
- la opción "fuente" (paso 1.7) manda sobre todo lo anterior.

No se tocan: el código de programación (APA pide una fuente monoespaciada),
los símbolos y las ecuaciones, ni el tamaño de las notas al pie.
"""

import re
from collections import Counter

from apa_facil import documento
from apa_facil.documento import Estilos, en_ecuacion, poner_fuente, poner_tamano, texto_parrafo, w

# Fuentes recomendadas por APA (web oficial, 2026) y su tamaño en puntos.
RECOMENDADAS = {
    "Aptos": 12,
    "Calibri": 11,
    "Arial": 11,
    "Lucida Sans Unicode": 10,
    "Times New Roman": 12,
    "Georgia": 11,
    "Computer Modern": 10,
}

# Nombres con que aparece Computer Modern (la fuente de LaTeX) en Word.
COMPUTER_MODERN = ("computer modern", "cmu serif", "cmu ", "latin modern", "lm roman")

# Fuente no recomendada → la recomendada más parecida.
PARECIDAS = {
    **dict.fromkeys(["segoe ui", "candara", "corbel", "gill sans", "gill sans mt", "open sans",
                     "roboto", "lato", "noto sans", "carlito", "source sans pro", "comic sans ms"],
                    "Calibri"),
    **dict.fromkeys(["helvetica", "liberation sans", "arimo", "verdana", "tahoma", "trebuchet ms",
                     "century gothic", "montserrat", "franklin gothic", "franklin gothic book",
                     "franklin gothic medium", "futura"], "Arial"),
    **dict.fromkeys(["lucida sans", "lucida grande"], "Lucida Sans Unicode"),
    **dict.fromkeys(["cambria", "garamond", "eb garamond", "book antiqua", "palatino",
                     "palatino linotype", "baskerville", "baskerville old face", "liberation serif",
                     "tinos", "times"], "Times New Roman"),
    **dict.fromkeys(["constantia", "century schoolbook", "bookman old style", "century",
                     "rockwell"], "Georgia"),
}

MONOESPACIADAS = ("courier", "consolas", "lucida console", "mono", "menlo", "monaco",
                  "cascadia", "source code", "fira code")
SIMBOLOS = {"symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings", "cambria math",
            "mt extra", "marlett", "segoe ui symbol", "segoe ui emoji", "zapf dingbats"}
TAMANO_CODIGO = 10

ESTILOS_DE_CODIGO = ("code", "código", "codigo", "preformat", "formato previo")
ESTILOS_DE_NOTAS = {"footnote text", "footnote reference", "endnote text", "endnote reference",
                    "annotation text", "annotation reference"}
PARTES_DE_NOTAS = ("/word/footnotes.xml", "/word/endnotes.xml")

# Señales de que un párrafo en fuente monoespaciada es código de programación.
PARECE_CODIGO = re.compile(
    r"[;{}]\s*$"                                            # termina en ; { }
    r"|^\s*(def|class|import|from|return|print|function|var|let|const|public|private"
    r"|static|void|int|for|while|if|elif|else)\b.*[:({=]"   # palabras de programación
    r"|==|!=|<=|>=|=>|->|\+\+|&&|\|\|"                      # operadores
    r"|^\s*(//|/\*|#include|#!)"                            # comentarios y directivas
    r"|\w\s+=\s+\S"                                         # asignación: x = 5
    r"|\w\([^)]*\)\s*;?\s*$"                                # llamada al final: print(x)
    r"|^ {2,}\S",                                           # sangría hecha con espacios
    re.MULTILINE)


# --- Elegir la fuente ---------------------------------------------------------

def familia(nombre):
    """La fuente recomendada a la que pertenece `nombre` (Calibri Light → Calibri), o None."""
    bajo = nombre.lower()
    if bajo.startswith(COMPUTER_MODERN):
        return "Computer Modern"
    for recomendada in RECOMENDADAS:
        if bajo == recomendada.lower() or bajo.startswith(recomendada.lower() + " "):
            return recomendada
    return None


def es_monoespaciada(nombre):
    return any(m in nombre.lower() for m in MONOESPACIADAS)


def es_de_simbolos(nombre):
    return nombre.lower() in SIMBOLOS


def mas_parecida(nombre):
    """(fuente recomendada más parecida, ¿la conocíamos?)."""
    bajo = nombre.lower()
    if bajo in PARECIDAS:
        return PARECIDAS[bajo], True
    if es_monoespaciada(nombre):
        return "Times New Roman", True
    return ("Arial" if "sans" in bajo else "Times New Roman"), False


def _nombres_de_estilo(estilos, p):
    return [estilos.nombre(s.get(w("styleId"))).lower() for s in estilos.cadena(estilos.estilo_de(p))]


def _runs_propios(p):
    """Los fragmentos de texto del párrafo (no los de cuadros de texto que lleve dentro)."""
    return [r for r in p.iter(w("r")) if next(r.iterancestors(w("p")), None) is p]


def _letras(r):
    return sum(len(t.text or "") for t in r.iter(w("t")))


def es_codigo(p, estilos):
    """¿Es un párrafo de código de programación?"""
    if any(c in nombre for nombre in _nombres_de_estilo(estilos, p) for c in ESTILOS_DE_CODIGO):
        return True
    letras = Counter()
    for r in _runs_propios(p):
        letras[es_monoespaciada(estilos.fuente(r, p))] += _letras(r)
    if letras[True] == 0 or letras[True] < letras[False]:
        return False
    return bool(PARECE_CODIGO.search(texto_parrafo(p)))


def _analizar(doc, estilos, intactos=frozenset()):
    """Letras por familia de fuente y por tamaño en el cuerpo, y párrafos de código."""
    familias, tamanos, codigo = Counter(), Counter(), []
    for numero, p in enumerate(doc.element.body.iter(w("p")), start=1):
        if p in intactos:   # la portada que se mantiene (reglas/portada.py)
            continue
        if es_codigo(p, estilos):
            codigo.append((numero, p))
            continue
        for r in _runs_propios(p):
            letras = _letras(r)
            fuente = estilos.fuente(r, p)
            if not letras or en_ecuacion(r) or es_de_simbolos(fuente):
                continue
            familias[familia(fuente) or fuente] += letras
            tamanos[estilos.tamano(r, p)] += letras
    return familias, tamanos, codigo


def _elegir(principal, opciones, informe):
    """(fuente que se aplica, tamaño, cómo se ha elegido)."""
    pedida = opciones.get("fuente")
    if pedida:
        elegida = familia(pedida)
        if elegida:
            return elegida, RECOMENDADAS[elegida], "pedida"
        informe.error(f"«{pedida}» no es una de las fuentes recomendadas por APA "
                      f"({_lista(RECOMENDADAS)}). Se ha elegido la fuente automáticamente.")
    if principal is None:   # documento sin texto
        return "Times New Roman", RECOMENDADAS["Times New Roman"], "por defecto"
    if principal in RECOMENDADAS:
        return principal, RECOMENDADAS[principal], "mantenida"
    destino, conocida = mas_parecida(principal)
    return destino, RECOMENDADAS[destino], "parecida" if conocida else "desconocida"


# --- Aplicarla ------------------------------------------------------------------

def _aplicar_en_estilos(doc, nombre, puntos):
    poner_fuente(documento.propiedades_por_defecto(doc), nombre)
    poner_tamano(documento.propiedades_por_defecto(doc), puntos)
    for estilo in doc.styles.element.iter(w("style")):
        nombre_estilo = (estilo.find(w("name")).get(w("val")) if estilo.find(w("name")) is not None
                         else "").lower()
        if any(c in nombre_estilo for c in ESTILOS_DE_CODIGO):
            continue
        for rpr in estilo.iter(w("rPr")):   # incluye el formato condicional de los estilos de tabla
            _cambiar(rpr, nombre, None if nombre_estilo in ESTILOS_DE_NOTAS else puntos)


def _cambiar(rpr, nombre, puntos):
    """Cambia la fuente y el tamaño solo donde ya estaban indicados (si no, se heredan)."""
    fuentes = rpr.find(w("rFonts"))
    if fuentes is not None and not es_de_simbolos(fuentes.get(w("ascii")) or ""):
        poner_fuente(rpr, nombre)
    if puntos is not None and rpr.find(w("sz")) is not None:
        poner_tamano(rpr, puntos)


def _aplicar_en_numeracion(doc, nombre, puntos):
    """Los números y viñetas de las listas (las viñetas de símbolos se respetan)."""
    try:
        numeracion = doc.part.numbering_part
    except (KeyError, NotImplementedError):   # el documento no tiene listas
        return
    for rpr in numeracion.element.iter(w("rPr")):
        _cambiar(rpr, nombre, puntos)


def _aplicar_en_texto(doc, estilos, nombre, puntos, intactos=frozenset()):
    def en_parte(nombre_parte, raiz):
        es_nota = nombre_parte.startswith(PARTES_DE_NOTAS)
        tamano = None if es_nota else puntos
        for p in raiz.iter(w("p")):
            if p in intactos:
                continue
            codigo = es_codigo(p, estilos)
            marca = p.find(f"{w('pPr')}/{w('rPr')}")   # formato de la marca de párrafo
            if marca is not None and not codigo:
                _cambiar(marca, nombre, tamano)
            for r in _runs_propios(p):
                if en_ecuacion(r) or r.find(w("sym")) is not None or es_de_simbolos(estilos.fuente(r, p)):
                    continue
                if codigo:
                    poner_tamano(r.get_or_add_rPr(), TAMANO_CODIGO)
                elif r.find(w("rPr")) is not None:
                    _cambiar(r.find(w("rPr")), nombre, tamano)
    documento.editar_partes_de_texto(doc, en_parte)


# --- La regla ---------------------------------------------------------------------

def aplicar(doc, opciones, informe, estructura=None):
    estilos = Estilos(doc)
    intactos = getattr(estructura, "intactos", frozenset())
    familias, tamanos, codigo = _analizar(doc, estilos, intactos)
    principal = familias.most_common(1)[0][0] if familias else None
    nombre, puntos, como = _elegir(principal, opciones, informe)

    # Computer Modern se llama de otra forma en Word (CMU Serif…): se respeta su nombre real.
    nombre_en_word = nombre
    if nombre == "Computer Modern" and principal == "Computer Modern":
        nombre_en_word = _nombre_original(doc, estilos)

    ya_cumplia = set(familias) <= {nombre} and set(tamanos) <= {puntos}
    _aplicar_en_estilos(doc, nombre_en_word, puntos)
    _aplicar_en_numeracion(doc, nombre_en_word, puntos)
    _aplicar_en_texto(doc, estilos, nombre_en_word, puntos, intactos)

    if not ya_cumplia:
        informe.cambio(_mensaje(familias, tamanos, principal, nombre, puntos, como))
    if como == "desconocida":
        informe.aviso(f"No conocemos la fuente «{principal}», así que se ha cambiado por {nombre} {puntos} pt. "
                      f"Si prefieres otra de las recomendadas por APA ({_lista(RECOMENDADAS)}), "
                      "puedes elegirla al procesar el documento.")
    for numero, p in codigo:
        fuente = next((estilos.fuente(r, p) for r in _runs_propios(p) if _letras(r)), "monoespaciada")
        informe.aviso(f"El párrafo {numero} («{_extracto(texto_parrafo(p))}») se ha dejado en {fuente} "
                      f"{TAMANO_CODIGO} pt porque parece código de programación, y APA pide para el "
                      "código una fuente monoespaciada. Si no es código, cámbialo a mano.")


def _nombre_original(doc, estilos):
    for p in doc.element.body.iter(w("p")):
        for r in _runs_propios(p):
            fuente = estilos.fuente(r, p)
            if _letras(r) and familia(fuente) == "Computer Modern":
                return fuente
    return "CMU Serif"


def _mensaje(familias, tamanos, principal, nombre, puntos, como):
    destino = f"{nombre} {_pt(puntos)} pt"
    if como == "pedida":
        return f"Fuente: se ha aplicado {destino} en todo el documento, como has pedido."
    if len(familias) > 1:
        lista = _lista(f for f, _ in familias.most_common())
        mensaje = f"Tu trabajo mezclaba {len(familias)} fuentes ({lista}). Se ha unificado todo en {destino}"
        if como == "mantenida":
            return mensaje + ", la fuente recomendada por APA que más usabas."
        return (mensaje + f": la que más usabas, {principal}, se ha cambiado por la fuente recomendada "
                "por APA más parecida.")
    if como == "mantenida":
        antes = _lista(_pt(t) for t in sorted(tamanos))
        return (f"Fuente: se mantiene {nombre}, recomendada por APA, y todo el texto pasa a "
                f"{_pt(puntos)} pt (antes había texto a {antes} pt).")
    if como == "por defecto":
        return f"Fuente: se ha aplicado {destino}."
    return f"Fuente: {principal} se ha cambiado por {destino}, la fuente recomendada por APA más parecida."


def _pt(puntos):
    return f"{puntos:g}".replace(".", ",")


def _lista(elementos):
    """["a", "b", "c"] → "a, b y c"."""
    elementos = list(elementos)
    if len(elementos) <= 1:
        return "".join(elementos)
    return ", ".join(elementos[:-1]) + " y " + elementos[-1]


def _extracto(texto, largo=40):
    texto = " ".join(texto.split())
    return texto if len(texto) <= largo else texto[:largo] + "…"
