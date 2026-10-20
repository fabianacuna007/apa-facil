import Guia, { Apartado, Lista, metadatosDe } from "@/componentes/Guia";

const SLUG = "formato-apa-word";
export const metadata = metadatosDe(SLUG);

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        La 7.ª edición de las normas APA tiene una versión para estudiantes, más sencilla que la de las publicaciones
        científicas. Estas son las reglas de formato que se aplican a todo el trabajo y cómo ponerlas en Microsoft
        Word. Los nombres de los menús son los de Word en español.
      </p>

      <Apartado id="resumen" titulo="El formato APA 7 en una lista">
        <Lista>
          <li><strong>Márgenes:</strong> 2,54 cm (una pulgada) en los cuatro lados.</li>
          <li><strong>Fuente:</strong> la misma en todo el trabajo, por ejemplo Calibri 11, Arial 11, Times New Roman 12, Georgia 11 o Aptos 12.</li>
          <li><strong>Interlineado:</strong> doble en todo el trabajo, sin espacio extra antes ni después de los párrafos.</li>
          <li><strong>Alineación:</strong> a la izquierda, sin justificar.</li>
          <li><strong>Sangría:</strong> 1,27 cm en la primera línea de cada párrafo.</li>
          <li><strong>Número de página:</strong> arriba a la derecha, en todas las páginas, empezando por la portada.</li>
        </Lista>
      </Apartado>

      <Apartado id="margenes" titulo="Márgenes">
        <p>
          APA pide márgenes de 2,54 cm arriba, abajo, a la izquierda y a la derecha. En Word: pestaña{" "}
          <strong>Disposición</strong> › <strong>Márgenes</strong> › <strong>Normal</strong>. Comprueba que el margen
          de encuadernación sea 0 (está en <strong>Márgenes personalizados</strong>), porque Word lo suma a uno de los
          márgenes.
        </p>
        <p>
          En un TFG, un TFM o una tesis, tu universidad puede pedir otros márgenes (por ejemplo, uno más ancho a la
          izquierda para encuadernar). En ese caso, mandan los de tu universidad.
        </p>
      </Apartado>

      <Apartado id="fuente" titulo="Fuente y tamaño">
        <p>
          APA admite varias fuentes legibles, entre ellas Aptos 12, Calibri 11, Arial 11, Lucida Sans Unicode 10, Times
          New Roman 12, Georgia 11 y Computer Modern 10 (la de LaTeX). Lo importante es usar <strong>la misma en todo el
          trabajo</strong>, también en los títulos y en las tablas. Hay tres excepciones:
        </p>
        <Lista>
          <li>el texto de dentro de las figuras (los ejes de un gráfico, por ejemplo), en una fuente sin serifa de 8 a 14 puntos;</li>
          <li>el código de programación, en una fuente monoespaciada, como Courier New 10;</li>
          <li>las notas al pie, que pueden quedarse con el tamaño que pone Word.</li>
        </Lista>
        <p>
          La forma más cómoda de cambiarla en Word es en el estilo <strong>Normal</strong>: en la pestaña{" "}
          <strong>Inicio</strong>, haz clic derecho en el estilo <strong>Normal</strong> › <strong>Modificar</strong>,
          elige la fuente y el tamaño, y pulsa <strong>Aceptar</strong>. Así cambia todo el texto que no tenga un
          formato puesto a mano.
        </p>
      </Apartado>

      <Apartado id="interlineado" titulo="Interlineado doble">
        <p>
          Todo el trabajo va a doble espacio: el resumen, el texto, las citas en bloque, los números, títulos y notas de
          las tablas y figuras, y la lista de referencias. Tampoco se deja espacio extra antes ni después de los
          párrafos. Las excepciones son el interior de las tablas (puede ser sencillo, 1,5 o doble) y las notas al pie.
        </p>
        <p>
          En Word: selecciona todo (<strong>Ctrl + E</strong>), y en <strong>Inicio</strong> abre el cuadro de{" "}
          <strong>Párrafo</strong> (la flechita de la esquina). En <strong>Interlineado</strong> elige{" "}
          <strong>Doble</strong>, y pon <strong>0 pto</strong> en <strong>Espaciado anterior</strong> y{" "}
          <strong>posterior</strong>.
        </p>
      </Apartado>

      <Apartado id="alineacion" titulo="Alineación a la izquierda">
        <p>
          El texto va alineado a la izquierda, con el borde derecho irregular. No se justifica, y no se cortan palabras
          con guiones a final de línea. Solo se centran la portada, los rótulos de las secciones (como «Referencias») y
          los títulos de nivel 1.
        </p>
      </Apartado>

      <Apartado id="sangria" titulo="Sangría de primera línea">
        <p>
          La primera línea de cada párrafo lleva 1,27 cm (media pulgada) de sangría. No la hagas con la barra
          espaciadora: en Word, en el cuadro de <strong>Párrafo</strong>, elige <strong>Especial</strong> ›{" "}
          <strong>Primera línea</strong> y <strong>1,27 cm</strong>.
        </p>
        <p>No llevan esa sangría:</p>
        <Lista>
          <li>la portada, que va centrada;</li>
          <li>los títulos de nivel 1 a 3 y los rótulos de las secciones;</li>
          <li>la primera línea del resumen;</li>
          <li>los números, títulos y notas de las tablas y figuras, que van pegados al margen;</li>
          <li>las citas en bloque, que llevan 1,27 cm de sangría en todas sus líneas;</li>
          <li>las referencias, que llevan sangría francesa (la primera línea pegada al margen y las demás con 1,27 cm).</li>
        </Lista>
      </Apartado>

      <Apartado id="numero-de-pagina" titulo="Número de página">
        <p>
          En un trabajo de estudiante, el encabezado solo lleva el número de página: arriba a la derecha, en todas las
          páginas, y la portada es la página 1. El «titulillo» (el título abreviado arriba a la izquierda) es de los
          artículos para publicar, y solo se pone si lo pide tu profesor.
        </p>
        <p>
          En Word: <strong>Insertar</strong> › <strong>Número de página</strong> › <strong>Principio de página</strong>{" "}
          › <strong>Número sin formato 3</strong>. Si tu trabajo tiene secciones, comprueba que la numeración no se
          reinicia en ninguna.
        </p>
      </Apartado>

      <Apartado id="orden" titulo="Orden de las partes">
        <p>Las partes de un trabajo APA van en este orden, y cada una empieza en una página nueva:</p>
        <Lista numerada>
          <li>portada;</li>
          <li>resumen (en los trabajos de estudiante, solo si lo piden);</li>
          <li>texto;</li>
          <li>referencias;</li>
          <li>notas al pie, si no van al pie de cada página;</li>
          <li>tablas y figuras, si no van dentro del texto;</li>
          <li>apéndices.</li>
        </Lista>
        <p>
          Las tablas y figuras pueden ir dentro del texto, después de la primera vez que se mencionan, o cada una en su
          página después de las referencias.
        </p>
      </Apartado>
    </Guia>
  );
}
