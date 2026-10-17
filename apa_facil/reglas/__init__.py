"""
Las reglas APA, una por archivo. Cada regla es una función

    aplicar(doc, opciones, informe, estructura=None)

que recibe el documento, las opciones del usuario, el informe donde anota
lo que hizo y la estructura del documento: la clasificación de sus párrafos
(estructura.clasificar), hecha sobre el original antes de aplicar ninguna
regla. Las reglas que no la necesitan la ignoran; si una regla que la
necesita se llama sin ella (por ejemplo, en sus pruebas), clasifica el
documento tal como lo recibe.
"""
