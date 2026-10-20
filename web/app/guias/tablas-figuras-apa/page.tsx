import Guia, { Apartado, Ejemplo, Lista, metadatosDe } from "@/componentes/Guia";

const SLUG = "tablas-figuras-apa";
export const metadata = metadatosDe(SLUG);

const FILAS = [
  ["Ciencias", "42", "6,8", "7,1"],
  ["Humanidades", "38", "7,3", "6,9"],
  ["Ingeniería", "40", "6,1", "6,7"],
];

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        En APA, una tabla es lo que se presenta en filas y columnas, y una figura es cualquier otra imagen: gráficos,
        diagramas, mapas, fotografías o dibujos. Las dos se preparan de la misma forma: número arriba, título debajo,
        el contenido y, si hace falta, una nota.
      </p>

      <Apartado id="partes" titulo="Las partes de una tabla o figura">
        <Lista>
          <li>
            <strong>Número:</strong> «Tabla 1», «Figura 1»…, en negrita y encima. Se numeran en el orden en que se
            mencionan en el texto.
          </li>
          <li>
            <strong>Título:</strong> en la línea siguiente, breve y descriptivo, en cursiva y con mayúscula en las
            palabras principales.
          </li>
          <li>
            <strong>Contenido:</strong> la tabla en sí, o la imagen de la figura.
          </li>
          <li>
            <strong>Nota:</strong> debajo, empezando con <em>Nota.</em> en cursiva, solo si hace falta explicar algo:
            abreviaturas, de dónde se ha sacado el contenido, qué significan los asteriscos…
          </li>
        </Lista>
        <p>
          El número, el título y la nota van pegados al margen izquierdo y a doble espacio. Usa la misma fuente que en
          el resto del trabajo.
        </p>
      </Apartado>

      <Apartado id="tablas" titulo="Cómo debe ser una tabla">
        <Lista>
          <li>
            Hazla con la herramienta de tablas de Word (<strong>Insertar</strong> › <strong>Tabla</strong>), nunca con
            tabuladores o espacios.
          </li>
          <li>Todas las columnas llevan encabezado, también la primera. Los encabezados van centrados.</li>
          <li>
            La primera columna va alineada a la izquierda; las demás, centradas (o a la izquierda si tienen mucho texto).
          </li>
          <li>El interior puede ir a espacio sencillo, 1,5 o doble, lo que se lea mejor.</li>
          <li>
            <strong>Bordes:</strong> solo una línea arriba, otra debajo de los encabezados y otra abajo. Sin líneas
            verticales y sin recuadrar cada celda.
          </li>
          <li>
            Si ocupa más de una página, la fila de encabezados se repite en cada una (en Word: selecciónala ›{" "}
            <strong>Disposición</strong> › <strong>Repetir filas de título</strong>). Si es demasiado ancha, ponla en una
            página horizontal.
          </li>
        </Lista>
        <Ejemplo titulo="Ejemplo (datos inventados)">
          <p className="font-bold">Tabla 1</p>
          <p className="italic">Horas de Sueño por Facultad Durante la Época de Exámenes</p>
          <table className="w-full border-y border-tinta text-sm">
            <thead>
              <tr className="border-b border-tinta">
                {["Facultad", "n", "Horas entre semana", "Horas en exámenes"].map((t) => (
                  <th key={t} scope="col" className="px-2 py-1 text-center font-normal">{t === "n" ? <em>n</em> : t}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {FILAS.map(([facultad, ...datos]) => (
                <tr key={facultad}>
                  <td className="px-2 py-0.5 text-left">{facultad}</td>
                  {datos.map((d, i) => <td key={i} className="px-2 py-0.5 text-center">{d}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
          <p><em>Nota.</em> Datos de una encuesta a estudiantes de primer curso.</p>
        </Ejemplo>
      </Apartado>

      <Apartado id="figuras" titulo="Cómo debe ser una figura">
        <Lista>
          <li>La imagen tiene que verse nítida, con los ejes y los elementos bien rotulados y las unidades indicadas.</li>
          <li>El texto de dentro de la imagen (ejes, etiquetas) va en una fuente sin serifa de 8 a 14 puntos.</li>
          <li>Si tiene leyenda (la clave de colores o símbolos), va dentro de la figura.</li>
        </Lista>
        <Ejemplo titulo="Cómo se coloca (ejemplo inventado)">
          <p className="font-bold">Figura 1</p>
          <p className="italic">Evolución de las Horas de Sueño a lo Largo del Curso</p>
          <p className="rounded border border-dashed border-borde-fuerte py-6 text-center text-sm text-tinta-tenue">
            [aquí va el gráfico]
          </p>
          <p><em>Nota.</em> Media semanal de los 120 participantes.</p>
        </Ejemplo>
      </Apartado>

      <Apartado id="donde" titulo="Dónde se colocan">
        <p>Hay dos opciones; elige una para todo el trabajo:</p>
        <Lista>
          <li>
            <strong>Dentro del texto</strong>, después del párrafo en que se mencionan por primera vez. Si comparten
            página con texto, van arriba o abajo de la página, no en medio, separadas del texto por una línea en blanco.
          </li>
          <li><strong>Al final</strong>, cada una en su página, después de las referencias (primero las tablas y luego las figuras).</li>
        </Lista>
        <p>
          En los dos casos, cada tabla y figura se menciona en el texto por su número: «como muestra la Tabla 1».
        </p>
      </Apartado>
    </Guia>
  );
}
