"""
Clasificación correcta de cada párrafo de los 5 documentos de pruebas/,
escrita a mano: es la referencia con la que se mide el clasificador de la
fase 3 (paso 3.1 y siguientes).

Cada párrafo del cuerpo (sin contar los de dentro de tablas) es una tupla
(tipo, nivel, zona). El nivel solo se indica en los títulos. El comentario
lleva la posición del párrafo y el principio de su texto.

Si se añade un documento a pruebas/, hay que añadir aquí su clasificación.
"""

ESPERADO = {
    "01_ensayo_basico.docx": [
        ("portada", None, "portada"),                 #  0 La Revolución Industrial y sus efectos en el
        ("portada", None, "portada"),                 #  1 Alumno: Javier Moreno Díaz
        ("portada", None, "portada"),                 #  2 Asignatura: Historia Contemporánea
        ("portada", None, "portada"),                 #  3 Profesora: Elena Sáez
        ("portada", None, "portada"),                 #  4 Curso 2025-2026
        ("titulo", 1, "cuerpo"),                      #  5 Introducción
        ("cuerpo", None, "cuerpo"),                   #  6 La Revolución Industrial supuso una transform
        ("cuerpo", None, "cuerpo"),                   #  7 En este ensayo se analizan tres cambios princ
        ("titulo", 1, "cuerpo"),                      #  8 1. Cambios en la organización del trabajo
        ("cuerpo", None, "cuerpo"),                   #  9 El taller artesanal se caracterizaba por una
        ("cuerpo", None, "cuerpo"),                   # 10 La mecanización también cambió el valor de la
        ("titulo", 1, "cuerpo"),                      # 11 2. El nacimiento de la clase obrera
        ("cuerpo", None, "cuerpo"),                   # 12 La concentración de trabajadores en las ciuda
        ("cuerpo", None, "cuerpo"),                   # 13 Con el tiempo, estas asociaciones evolucionar
        ("titulo", 1, "cuerpo"),                      # 14 3. Las primeras leyes laborales
        ("cuerpo", None, "cuerpo"),                   # 15 A lo largo del siglo XIX, varios gobiernos eu
        ("titulo", 1, "cuerpo"),                      # 16 Conclusión
        ("cuerpo", None, "cuerpo"),                   # 17 La Revolución Industrial no solo transformó l
        ("titulo_referencias", None, "referencias"),  # 18 BIBLIOGRAFÍA
        ("referencia", None, "referencias"),          # 19 Ramírez Soto, M. (2015). Fábricas y ciudades:
        ("referencia", None, "referencias"),          # 20 Ortega, L. y Blanco, R. (2018). El tiempo de
        ("referencia", None, "referencias"),          # 21 Lefèvre, A. 2009. Industrialisation et travai
        ("referencia", None, "referencias"),          # 22 https://www.historiaejemplo.es/revolucion-ind
    ],
    "02_informe_tablas.docx": [
        ("portada", None, "portada"),                 #  0 Hábitos de sueño y rendimiento académico en e
        ("portada", None, "portada"),                 #  1 Sara Delgado Ruiz · Psicología de la Salud ·
        ("titulo", 1, "resumen"),                     #  2 Resumen
        ("cuerpo", None, "resumen"),                  #  3 Este informe analiza la relación entre las ho
        ("titulo", 1, "cuerpo"),                      #  4 Introducción
        ("cuerpo", None, "cuerpo"),                   #  5 El sueño cumple funciones esenciales en la co
        ("cuerpo", None, "cuerpo"),                   #  6 Según la revisión de Castro et al. (2019), en
        ("titulo", 1, "cuerpo"),                      #  7 Método
        ("titulo", 2, "cuerpo"),                      #  8 Participantes
        ("cuerpo", None, "cuerpo"),                   #  9 Participaron 180 estudiantes (112 mujeres y 6
        ("titulo", 2, "cuerpo"),                      # 10 Instrumentos
        ("cuerpo", None, "cuerpo"),                   # 11 Se utilizó el Índice de Calidad de Sueño de P
        ("titulo", 2, "cuerpo"),                      # 12 Procedimiento
        ("cuerpo", None, "cuerpo"),                   # 13 Los cuestionarios se aplicaron en horario de
        ("titulo", 1, "cuerpo"),                      # 14 Resultados
        ("cuerpo", None, "cuerpo"),                   # 15 En la tabla 1 se presentan los estadísticos d
        ("leyenda_tabla", None, "cuerpo"),            # 16 Tabla 1: Estadísticos descriptivos de las var
        ("cuerpo", None, "cuerpo"),                   # 17 La tabla 2 muestra el rendimiento académico s
        ("leyenda_tabla", None, "cuerpo"),            # 18 Tabla 2: Rendimiento académico según las hora
        ("titulo", 1, "cuerpo"),                      # 19 Discusión
        ("cuerpo", None, "cuerpo"),                   # 20 Los resultados coinciden con estudios previos
        ("titulo_referencias", None, "referencias"),  # 21 Referencias
        ("referencia", None, "referencias"),          # 22 Castro, P., Llorente, M. y Benítez, J. (2019)
        ("referencia", None, "referencias"),          # 23 Navarro, A. y Gil, S. (2020). Sueño, memoria
    ],
    "03_trabajo_imagenes_listas.docx": [
        ("portada", None, "portada"),                 #  0 El consumo de energía en los hogares: diferen
        ("portada", None, "portada"),                 #  1 Trabajo en grupo: Marta Pino, Daniel Luengo y
        ("titulo", 1, "cuerpo"),                      #  2 1. Introducción
        ("cuerpo", None, "cuerpo"),                   #  3 El consumo de energía en el ámbito doméstico
        ("cuerpo", None, "cuerpo"),                   #  4 En este trabajo nos centramos en tres factore
        ("lista", None, "cuerpo"),                    #  5 El tipo de vivienda y su aislamiento térmico.
        ("lista", None, "cuerpo"),                    #  6 Los sistemas de calefacción y refrigeración d
        ("lista", None, "cuerpo"),                    #  7 Los hábitos de uso de los electrodomésticos.
        ("titulo", 1, "cuerpo"),                      #  8 2. Datos y resultados
        ("cuerpo", None, "cuerpo"),                   #  9 A partir de los datos de facturación de 60 ho
        ("imagen", None, "cuerpo"),                   # 10 (sin texto)
        ("leyenda_figura", None, "cuerpo"),           # 11 Gráfico 1 – Consumo eléctrico medio mensual (
        ("titulo", 2, "cuerpo"),                      # 12 2.1. Fuentes de energía
        ("cuerpo", None, "cuerpo"),                   # 13 Los hogares estudiados combinan distintas fue
        ("imagen", None, "cuerpo"),                   # 14 (sin texto)
        ("leyenda_figura", None, "cuerpo"),           # 15 Imagen 2: fuentes de energía de los hogares (
        ("titulo", 1, "cuerpo"),                      # 16 3. Propuestas de ahorro
        ("cuerpo", None, "cuerpo"),                   # 17 A partir de los resultados, proponemos las si
        ("lista", None, "cuerpo"),                    # 18 Ajustar la temperatura de la calefacción y de
        ("lista", None, "cuerpo"),                    # 19 Sustituir la iluminación por tecnología LED.
        ("lista", None, "cuerpo"),                    # 20 Mejorar el aislamiento de la vivienda:
        ("lista", None, "cuerpo"),                    # 21 ventanas de doble acristalamiento;
        ("lista", None, "cuerpo"),                    # 22 aislamiento de la cubierta.
        ("lista", None, "cuerpo"),                    # 23 Instalar placas solares para autoconsumo.
        ("titulo", 1, "cuerpo"),                      # 24 4. Conclusiones
        ("cuerpo", None, "cuerpo"),                   # 25 Las diferencias de consumo entre el medio urb
        ("titulo_referencias", None, "referencias"),  # 26 Bibliografía
        ("referencia", None, "referencias"),          # 27 Herrero, C. (2021). Energía y vivienda en Esp
        ("referencia", None, "referencias"),          # 28 Campos, F., & Vidal, N. (2022). Autoconsumo s
    ],
    "04_formato_caotico.docx": [
        ("portada", None, "portada"),                 #  0 REDES SOCIALES Y AUTOESTIMA EN LA ADOLESCENCI
        ("vacio", None, "portada"),                   #  1 (sin texto)
        ("vacio", None, "portada"),                   #  2 (sin texto)
        ("portada", None, "portada"),                 #  3 Trabajo de Psicología del Desarrollo Nombre:
        ("vacio", None, "portada"),                   #  4 (sin texto)
        ("titulo", 1, "cuerpo"),                      #  5 ¿Qué son las redes sociales?
        ("cuerpo", None, "cuerpo"),                   #  6 Las redes sociales son plataformas digitales
        ("cuerpo", None, "cuerpo"),                   #  7 Durante esta etapa, la identidad todavía se e
        ("vacio", None, "cuerpo"),                    #  8 (sin texto)
        ("titulo", 1, "cuerpo"),                      #  9 Efectos sobre la autoestima
        ("cuerpo", None, "cuerpo"),                   # 10 Los estudios consultados señalan efectos tant
        ("cuerpo", None, "cuerpo"),                   # 11 Por otro lado, las redes también pueden ofrec
        ("vacio", None, "cuerpo"),                    # 12 (sin texto)
        ("titulo", 2, "cuerpo"),                      # 13 Factores de riesgo:
        ("lista", None, "cuerpo"),                    # 14 • Comparación social con imágenes retocadas
        ("lista", None, "cuerpo"),                    # 15 • Búsqueda de "me gusta" como medida de valor
        ("lista", None, "cuerpo"),                    # 16 - Ciberacoso
        ("lista", None, "cuerpo"),                    # 17 - Uso nocturno que reduce las horas de sueño
        ("vacio", None, "cuerpo"),                    # 18 (sin texto)
        ("cuerpo", None, "cuerpo"),                   # 19 Más información en: www.example.org/adolescen
        ("cuerpo", None, "cuerpo"),                   # 20 La tabla de la página siguiente resume los da
        ("vacio", None, "cuerpo"),                    # 21 (sin texto)
        ("leyenda_tabla", None, "cuerpo"),            # 22 Datos de uso por red social
        ("vacio", None, "cuerpo"),                    # 23 (sin texto)
        ("titulo", 1, "cuerpo"),                      # 24 Conclusión
        ("cuerpo", None, "cuerpo"),                   # 25 Las redes sociales no son buenas ni malas en
        ("vacio", None, "cuerpo"),                    # 26 (sin texto)
        ("vacio", None, "cuerpo"),                    # 27 (sin texto)
        ("vacio", None, "cuerpo"),                    # 28 (sin texto)
        ("titulo_referencias", None, "referencias"),  # 29 Webgrafía:
        ("referencia", None, "referencias"),          # 30 Martínez, J. Redes sociales y adolescencia. 2
        ("referencia", None, "referencias"),          # 31 Pérez, L.; Romero, A.; Saiz, T.; Moya, P. (20
        ("referencia", None, "referencias"),          # 32 www.example.org/datos-uso-redes-2024
    ],
    "05_tfg_completo.docx": [
        ("imagen", None, "portada"),                  #  0 (sin texto)
        ("portada", None, "portada"),                 #  1 UNIVERSIDAD AUTÓNOMA DEL GUADIANA
        ("portada", None, "portada"),                 #  2 Facultad de Ciencias Sociales y Humanas
        ("portada", None, "portada"),                 #  3 Grado en Geografía y Ordenación del Territori
        ("portada", None, "portada"),                 #  4 TURISMO RURAL Y DESPOBLACIÓN EN LA ESPAÑA INT
        ("portada", None, "portada"),                 #  5 estudio de cuatro municipios
        ("portada", None, "portada"),                 #  6 TRABAJO FIN DE GRADO
        ("portada", None, "portada"),                 #  7 Autora: Lucía Fernández Prieto
        ("portada", None, "portada"),                 #  8 Tutor: Dr. Andrés Villalobos Ruiz
        ("portada", None, "portada"),                 #  9 Convocatoria: junio de 2026
        ("vacio", None, "portada"),                   # 10 (sin texto)
        ("titulo", 1, "resumen"),                     # 11 RESUMEN
        ("cuerpo", None, "resumen"),                  # 12 Este trabajo analiza el papel del turismo rur
        ("cuerpo", None, "resumen"),                  # 13 Palabras clave: turismo rural, despoblación,
        ("vacio", None, "resumen"),                   # 14 (sin texto)
        ("indice", None, "indice"),                   # 15 (sin texto)
        ("indice", None, "indice"),                   # 16 (sin texto)
        ("vacio", None, "indice"),                    # 17 (sin texto)
        ("titulo", 1, "cuerpo"),                      # 18 1. Introducción
        ("cuerpo", None, "cuerpo"),                   # 19 La despoblación del medio rural es uno de los
        ("cuerpo", None, "cuerpo"),                   # 20 En las últimas décadas, el turismo rural se h
        ("titulo", 2, "cuerpo"),                      # 21 1.1. Objetivos
        ("titulo", 3, "cuerpo"),                      # 22 1.1.1. Objetivo general
        ("cuerpo", None, "cuerpo"),                   # 23 Analizar la relación entre el desarrollo del
        ("titulo", 3, "cuerpo"),                      # 24 1.1.2. Objetivos específicos
        ("titulo", 4, "cuerpo"),                      # 25 a) Evolución demográfica
        ("cuerpo", None, "cuerpo"),                   # 26 Describir la evolución de la población de los
        ("titulo", 4, "cuerpo"),                      # 27 b) Percepción de los empresarios
        ("cuerpo", None, "cuerpo"),                   # 28 Conocer la valoración que hacen los empresari
        ("titulo", 1, "cuerpo"),                      # 29 2. Marco teórico
        ("cuerpo", None, "cuerpo"),                   # 30 La literatura sobre desarrollo rural distingu
        ("cita_bloque", None, "cuerpo"),              # 31 "El turismo rural puede complementar las rent
        ("cuerpo", None, "cuerpo"),                   # 32 Por su parte, Herranz (2015) señala que los e
        ("titulo", 1, "cuerpo"),                      # 33 3. Metodología
        ("cuerpo", None, "cuerpo"),                   # 34 Se seleccionaron cuatro municipios con distin
        ("titulo", 1, "cuerpo"),                      # 35 4. Resultados
        ("titulo", 2, "cuerpo"),                      # 36 4.1. Evolución de la población
        ("cuerpo", None, "cuerpo"),                   # 37 La tabla siguiente muestra la evolución de la
        ("leyenda_tabla", None, "cuerpo"),            # 38 Tabla 1. Evolución de la población (2001-2024
        ("leyenda_tabla", None, "cuerpo"),            # 39 Fuente: elaboración propia a partir del padró
        ("cuerpo", None, "cuerpo"),                   # 40 Todos los municipios pierden población, pero
        ("titulo", 2, "cuerpo"),                      # 41 4.2. Percepción de los empresarios
        ("cuerpo", None, "cuerpo"),                   # 42 Las personas entrevistadas coinciden en que e
        ("titulo", 1, "cuerpo"),                      # 43 5. Conclusiones
        ("cuerpo", None, "cuerpo"),                   # 44 El turismo rural contribuye a mantener activi
        ("vacio", None, "cuerpo"),                    # 45 (sin texto)
        ("titulo_referencias", None, "referencias"),  # 46 BIBLIOGRAFÍA
        ("referencia", None, "referencias"),          # 47 Molina, R. (2017). Turismo y territorio: lími
        ("referencia", None, "referencias"),          # 48 Collantes, F. y Pinilla, V. (2011). ¿Lugares
        ("referencia", None, "referencias"),          # 49 Consejo Económico de Ejemplo (2022). Informe
        ("referencia", None, "referencias"),          # 50 Herranz, M. Turismo rural y servicios público
        ("referencia", None, "referencias"),          # 51 Sáenz Gil, P. (2020). Nuevos pobladores en el
        ("referencia", None, "referencias"),          # 52 Torres, E. (2019). Agroturismo y empleo femen
        ("vacio", None, "referencias"),               # 53 (sin texto)
        ("anexo", None, "anexos"),                    # 54 Anexo I. Guion de las entrevistas
        ("lista", None, "anexos"),                    # 55 1. ¿Desde cuándo funciona su negocio y por qu
        ("lista", None, "anexos"),                    # 56 2. ¿Cómo ha evolucionado el número de visitan
        ("lista", None, "anexos"),                    # 57 3. ¿Cuántas personas trabajan en el negocio y
        ("lista", None, "anexos"),                    # 58 4. ¿Qué dificultades encuentra para mantener
        ("lista", None, "anexos"),                    # 59 5. ¿Cree que el turismo ayuda a que la gente
    ],
}
