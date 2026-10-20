import Guia, { Apartado, Ejemplo, Lista, metadatosDe } from "@/componentes/Guia";

const SLUG = "portada-apa";
export const metadata = metadatosDe(SLUG);

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        Todo trabajo en formato APA empieza con una portada. APA tiene dos versiones, la de estudiante y la profesional
        (la de los artículos para publicar). Si eres estudiante, usa la de estudiante, salvo que tu profesor o tu
        universidad pidan otra.
      </p>

      <Apartado id="elementos" titulo="Qué lleva la portada de estudiante">
        <p>Estos elementos, en este orden y cada uno en su línea:</p>
        <Lista numerada>
          <li>
            <strong>El título del trabajo</strong>, en negrita y centrado, entre tres y cuatro líneas por debajo del
            margen superior. Va con mayúscula en las palabras principales. Si tiene subtítulo, puede ir en otra línea.
            No hay un largo máximo, pero conviene que sea concreto y que lleve las palabras clave del tema.
          </li>
          <li><strong>Una línea en blanco.</strong></li>
          <li>
            <strong>Los autores</strong>, centrados. Con dos autores, se separan con «y»; con tres o más, con comas y
            una «y» antes del último.
          </li>
          <li>
            <strong>La afiliación</strong>: el departamento o la facultad y la universidad, separados por una coma.
          </li>
          <li>
            <strong>La asignatura</strong>: el código, dos puntos y el nombre, tal como aparecen en los materiales del
            curso.
          </li>
          <li><strong>El profesor</strong>, escrito como aparece en los materiales del curso.</li>
          <li><strong>La fecha de entrega</strong>, en el formato de tu país (por ejemplo, «20 de octubre de 2026»).</li>
        </Lista>
        <p>
          Todo va centrado y a doble espacio, y la portada lleva el número de página 1 arriba a la derecha, como el
          resto del trabajo. La versión de estudiante no lleva titulillo ni nota del autor.
        </p>
      </Apartado>

      <Apartado id="ejemplo" titulo="Ejemplo de portada">
        <Ejemplo titulo="Ejemplo (datos inventados)">
          <div className="flex flex-col items-center gap-1 text-center">
            <span className="self-end text-sm">1</span>
            <span aria-hidden className="h-10" />
            <strong>Hábitos de Sueño y Rendimiento Académico en Estudiantes de Primer Curso</strong>
            <span aria-hidden className="h-4" />
            <span>Lucía Martín Gómez y Daniel Ortega Ruiz</span>
            <span>Facultad de Psicología, Universidad de Ejemplo</span>
            <span>PSI 104: Psicología del Aprendizaje</span>
            <span>Dra. Elena Navarro</span>
            <span>20 de octubre de 2026</span>
          </div>
        </Ejemplo>
      </Apartado>

      <Apartado id="word" titulo="Cómo hacerla en Word">
        <Lista numerada>
          <li>En la primera página, pulsa <strong>Intro</strong> tres o cuatro veces antes de escribir el título.</li>
          <li>Escribe el título, céntralo y ponlo en negrita.</li>
          <li>Deja una línea en blanco y escribe cada dato en su línea, centrado y sin negrita.</li>
          <li>
            Al final de la fecha, inserta un salto de página (<strong>Insertar</strong> › <strong>Salto de página</strong>)
            para que el texto empiece en la página 2.
          </li>
        </Lista>
      </Apartado>

      <Apartado id="universidad" titulo="¿Y si tu universidad tiene su propia portada?">
        <p>
          Muchas universidades españolas e hispanoamericanas exigen su propia portada en los TFG, los TFM y otros
          trabajos, con el logo y unos datos concretos. En ese caso, usa la de tu universidad: el resto del trabajo
          puede seguir el formato APA igualmente.
        </p>
        <p>
          APA Fácil hace las dos cosas: si rellenas los datos de la portada, pone la portada APA de estudiante; si no,
          o si marcas «Mantener la portada de mi universidad», deja la tuya tal como está.
        </p>
      </Apartado>
    </Guia>
  );
}
