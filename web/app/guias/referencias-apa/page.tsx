import Guia, { Apartado, Ejemplo, Lista, Referencia, metadatosDe } from "@/componentes/Guia";

const SLUG = "referencias-apa";
export const metadata = metadatosDe(SLUG);

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        La lista de referencias recoge, al final del trabajo, todas las obras que citas en el texto, con los datos
        necesarios para encontrarlas. Solo van las que citas: no es una bibliografía de todo lo que has leído.
      </p>

      <Apartado id="formato" titulo="Formato de la lista">
        <Lista>
          <li>Empieza en una página nueva, después del texto.</li>
          <li>
            Arriba, el rótulo <strong>Referencias</strong>, centrado y en negrita («Referencia» si solo hay una). No
            «Bibliografía», «Obras citadas» ni «Webgrafía».
          </li>
          <li>Cada referencia es un párrafo, alineado a la izquierda.</li>
          <li>
            Por orden alfabético de la primera palabra de cada referencia, normalmente el apellido del primer autor. Si
            la obra no tiene autor, cuenta el título.
          </li>
          <li>
            Con <strong>sangría francesa</strong> de 1,27 cm: la primera línea pegada al margen y las demás con
            sangría. En Word: <strong>Párrafo</strong> › <strong>Especial</strong> › <strong>Francesa</strong>,{" "}
            <strong>1,27 cm</strong>.
          </li>
          <li>A doble espacio, sin espacio extra entre referencias.</li>
          <li>
            Cada obra citada en el texto tiene su referencia, y cada referencia se cita al menos una vez. Las
            comunicaciones personales (una conversación, un correo) solo se citan en el texto.
          </li>
        </Lista>
      </Apartado>

      <Apartado id="elementos" titulo="Los cuatro elementos de una referencia">
        <p>Casi todas las referencias tienen la misma estructura: <strong>autor, fecha, título y fuente</strong>.</p>
        <Lista>
          <li>
            <strong>Autor:</strong> apellidos, coma e iniciales («Romero, L. A.»), con un espacio entre las iniciales.
            Se escriben hasta 20 autores, separados por comas, con «y» antes del último (APA, en inglés, usa «&amp;»).
            Si el autor es una organización, su nombre completo, sin siglas.
          </li>
          <li>
            <strong>Fecha:</strong> el año entre paréntesis y un punto: «(2021).». Si no tiene fecha, «(s. f.).».
          </li>
          <li>
            <strong>Título:</strong> con mayúscula solo en la primera palabra (y en los nombres propios). Los libros y
            las páginas web van en cursiva; los artículos y los capítulos, no.
          </li>
          <li>
            <strong>Fuente:</strong> dónde está la obra: la revista, la editorial, el libro que la contiene o la web. Al
            final, el DOI o la dirección web.
          </li>
        </Lista>
        <p>
          El DOI se escribe como enlace: «https://doi.org/…». Si una obra tiene DOI, se pone siempre, aunque la hayas
          leído en papel. No se pone la ciudad de la editorial ni el ISBN.
        </p>
      </Apartado>

      <Apartado id="ejemplos" titulo="Ejemplos">
        <p>Todos los ejemplos son inventados (también los DOI y las direcciones web).</p>
        <Ejemplo titulo="Artículo de revista">
          <Referencia>
            Romero, L. A., y Vidal, M. (2021). Horas de sueño y rendimiento en estudiantes universitarios.{" "}
            <em>Revista de Psicología Educativa</em>, <em>14</em>(2), 45–60. https://doi.org/10.0000/rpe.2021.045
          </Referencia>
        </Ejemplo>
        <p>
          El nombre de la revista y el volumen van en cursiva; el número, entre paréntesis y sin cursiva; después, las
          páginas. Si el artículo no tiene DOI y lo has sacado de una base de datos académica, la referencia acaba en
          las páginas.
        </p>
        <Ejemplo titulo="Libro">
          <Referencia>
            Ortega, D. (2019). <em>El sueño en la adolescencia: Causas y consecuencias</em> (2.ª ed.). Editorial Ejemplo.
          </Referencia>
        </Ejemplo>
        <p>La edición, si no es la primera, va entre paréntesis detrás del título y sin cursiva.</p>
        <Ejemplo titulo="Capítulo de un libro con editores">
          <Referencia>
            Gil, P. (2020). Estrategias de estudio en la universidad. En L. Navarro y R. Díaz (Eds.),{" "}
            <em>Aprender a aprender</em> (pp. 33–58). Editorial Ejemplo.
          </Referencia>
        </Ejemplo>
        <p>
          Esta forma es solo para libros con editores, en los que cada capítulo tiene su autor. Si todo el libro es del
          mismo autor, la referencia es la del libro entero.
        </p>
        <Ejemplo titulo="Página web">
          <Referencia>
            Fundación Ejemplo para la Salud. (2023, 5 de marzo). <em>Consejos para dormir mejor en época de
            exámenes</em>. https://www.ejemplo.org/consejos-sueno
          </Referencia>
        </Ejemplo>
        <p>
          Si la página es de una organización y no se indica otro autor, el autor es la organización. Pon la fecha más
          concreta que aparezca.
        </p>
      </Apartado>

      <Apartado id="errores" titulo="Errores frecuentes">
        <Lista>
          <li>Titular la lista «Bibliografía» en lugar de «Referencias».</li>
          <li>Poner la ciudad de la editorial («Madrid: Editorial Ejemplo»): en APA 7 ya no se pone.</li>
          <li>Escribir el DOI como «doi:10.…» en lugar de «https://doi.org/10.…».</li>
          <li>Separar los autores con punto y coma o usar «&amp;» en un trabajo en español.</li>
          <li>Dejar referencias que no se citan en el texto, o citas sin su referencia.</li>
        </Lista>
      </Apartado>
    </Guia>
  );
}
