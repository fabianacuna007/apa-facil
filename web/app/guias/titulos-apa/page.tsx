import Guia, { Apartado, Ejemplo, Lista, metadatosDe } from "@/componentes/Guia";

const SLUG = "titulos-apa";
export const metadata = metadatosDe(SLUG);

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        APA tiene cinco niveles de título para organizar las secciones de un trabajo. El nivel 1 es el principal, el
        nivel 2 es un apartado dentro del 1, el 3 un apartado dentro del 2, y así hasta el 5. Usa solo los niveles que
        necesites: un trabajo corto puede no tener ninguno.
      </p>

      <Apartado id="niveles" titulo="Los cinco niveles">
        <Lista>
          <li><strong>Nivel 1:</strong> centrado y en negrita. El texto empieza en un párrafo nuevo.</li>
          <li><strong>Nivel 2:</strong> a la izquierda y en negrita. El texto empieza en un párrafo nuevo.</li>
          <li><strong>Nivel 3:</strong> a la izquierda, en negrita y cursiva. El texto empieza en un párrafo nuevo.</li>
          <li>
            <strong>Nivel 4:</strong> con sangría, en negrita y terminado en punto. El texto sigue en la misma línea.
          </li>
          <li>
            <strong>Nivel 5:</strong> con sangría, en negrita y cursiva, y terminado en punto. El texto sigue en la
            misma línea.
          </li>
        </Lista>
        <p>En todos, las palabras principales van con mayúscula inicial.</p>
        <Ejemplo titulo="Cómo se ven (ejemplo inventado)">
          <p className="text-center font-bold">Método</p>
          <p className="indent-8">Para este estudio se encuestó a estudiantes de tres facultades…</p>
          <p className="font-bold">Participantes</p>
          <p className="indent-8">Participaron 120 estudiantes de primer curso…</p>
          <p className="font-bold italic">Criterios de Selección</p>
          <p className="indent-8">Se incluyó a quienes estaban matriculados a tiempo completo…</p>
          <p className="indent-8">
            <strong>Estudiantes de Ciencias.</strong> El grupo de ciencias tenía más horas de laboratorio…
          </p>
          <p className="indent-8">
            <strong className="italic">Estudiantes de Primer Año.</strong> En este grupo, la mitad vivía fuera de casa…
          </p>
        </Ejemplo>
      </Apartado>

      <Apartado id="errores" titulo="Errores frecuentes">
        <Lista>
          <li><strong>Numerar los títulos.</strong> APA no usa números ni letras delante («1.», «2.3.», «a)»).</li>
          <li>
            <strong>Empezar con «Introducción».</strong> La introducción no lleva título: el título del trabajo, al
            principio de la primera página de texto, hace de título de nivel 1. Dentro de la introducción sí puede
            haber apartados de nivel 2.
          </li>
          <li><strong>Dejar un solo apartado dentro de una sección.</strong> Si hay un nivel 2, que haya al menos dos.</li>
          <li>
            <strong>Dejar líneas en blanco antes o después de un título</strong>, aunque el título caiga al final de una
            página. Los títulos van a doble espacio, como el resto.
          </li>
          <li>
            <strong>Escribirlos enteros en mayúsculas.</strong> APA solo pone mayúscula al principio de las palabras
            principales.
          </li>
        </Lista>
      </Apartado>

      <Apartado id="mayusculas" titulo="¿Mayúsculas en todas las palabras?">
        <p>
          APA, que es una norma en inglés, pide mayúscula inicial en las palabras principales de los títulos
          («Criterios de Selección»), pero no en artículos, preposiciones ni conjunciones cortas. En español es más
          habitual escribir con mayúscula solo la primera palabra («Criterios de selección»), y algunos profesores lo
          prefieren así. Si el tuyo lo pide, haz caso.
        </p>
      </Apartado>

      <Apartado id="word" titulo="Cómo hacerlo en Word">
        <p>
          Usa los estilos de Word (<strong>Título 1</strong>, <strong>Título 2</strong>…, en la pestaña{" "}
          <strong>Inicio</strong>): además del formato, sirven para el panel de navegación, para el índice automático y
          para quien lee con un lector de pantalla. Para que sigan el formato APA, haz clic derecho en cada estilo ›{" "}
          <strong>Modificar</strong> y cambia la alineación, la negrita y la cursiva; quita también el color y el
          espacio antes y después.
        </p>
        <p>
          Los niveles 4 y 5 van en la misma línea que su texto: escribe el título y el principio del párrafo, selecciona
          solo el título y aplícale el estilo.
        </p>
      </Apartado>
    </Guia>
  );
}
