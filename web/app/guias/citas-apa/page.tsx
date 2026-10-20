import Guia, { Apartado, Ejemplo, Lista, metadatosDe } from "@/componentes/Guia";

const SLUG = "citas-apa";
export const metadata = metadatosDe(SLUG);

const AUTORES = [
  ["Un autor", "(Romero, 2021)", "Romero (2021)"],
  ["Dos autores", "(Romero y Vidal, 2021)", "Romero y Vidal (2021)"],
  ["Tres o más", "(Romero et al., 2021)", "Romero et al. (2021)"],
  ["Organización con siglas (primera vez)", "(Instituto Nacional de Estadística [INE], 2022)",
   "Instituto Nacional de Estadística (INE, 2022)"],
  ["Organización con siglas (después)", "(INE, 2022)", "INE (2022)"],
];

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        APA usa el sistema autor-fecha: cada vez que usas una idea o unas palabras de otra persona, pones en el texto
        el apellido del autor y el año, y la obra completa va en la lista de referencias del final. Cada cita tiene
        que tener su referencia, y cada referencia, al menos una cita.
      </p>

      <Apartado id="tipos" titulo="Cita entre paréntesis o cita narrativa">
        <p>Hay dos formas de citar, y las dos son correctas:</p>
        <Lista>
          <li>
            <strong>Entre paréntesis:</strong> el autor y el año, separados por una coma, dentro de un paréntesis, en
            mitad o al final de la frase.
          </li>
          <li>
            <strong>Narrativa:</strong> el apellido forma parte de la frase, y el año va entre paréntesis justo
            detrás.
          </li>
        </Lista>
        <Ejemplo titulo="Ejemplo (cita inventada)">
          <p>Dormir menos de seis horas empeora la memoria a corto plazo (Romero, 2021).</p>
          <p>Romero (2021) comprobó que dormir menos de seis horas empeora la memoria a corto plazo.</p>
        </Ejemplo>
        <p>
          Si el paréntesis lleva más texto, el año va entre comas: «(véase Romero, 2021, para más detalles)». Y si en un
          paréntesis van juntos un texto y una cita, se separan con punto y coma, sin paréntesis dentro de paréntesis.
        </p>
      </Apartado>

      <Apartado id="autores" titulo="Cuántos autores se escriben">
        <p>
          Con uno o dos autores, se escriben siempre todos. Con tres o más, solo el primero seguido de «et al.», desde
          la primera cita.
        </p>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[32rem] border-y border-tinta text-left text-sm">
            <thead>
              <tr className="border-b border-tinta">
                <th scope="col" className="px-2 py-1.5 font-semibold text-tinta">Autores</th>
                <th scope="col" className="px-2 py-1.5 font-semibold text-tinta">Entre paréntesis</th>
                <th scope="col" className="px-2 py-1.5 font-semibold text-tinta">Narrativa</th>
              </tr>
            </thead>
            <tbody>
              {AUTORES.map(([tipo, parentesis, narrativa]) => (
                <tr key={tipo}>
                  <td className="px-2 py-1.5">{tipo}</td>
                  <td className="px-2 py-1.5 text-tinta">{parentesis}</td>
                  <td className="px-2 py-1.5 text-tinta">{narrativa}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p>
          En inglés, APA usa «&amp;» entre autores dentro del paréntesis; en un trabajo en español se escribe «y» en los
          dos casos. Las siglas de una organización se presentan solo la primera vez, y después se usan siempre.
        </p>
      </Apartado>

      <Apartado id="fechas" titulo="El año">
        <Lista>
          <li>Solo el año, aunque en la referencia haya también el día y el mes.</li>
          <li>El mismo año que en la referencia.</li>
          <li>Si la obra no tiene fecha: «s. f.» (sin fecha), por ejemplo «(Ministerio de Educación, s. f.)».</li>
          <li>Si está aceptada pero aún no publicada: «en prensa».</li>
          <li>
            Si citas dos obras del mismo autor y el mismo año, añade una letra al año en la cita y en la referencia:
            «(Romero, 2021a, 2021b)».
          </li>
        </Lista>
      </Apartado>

      <Apartado id="varias" titulo="Varias obras en la misma cita">
        <p>
          Se ordenan alfabéticamente, como en la lista de referencias, y se separan con punto y coma. Si son del mismo
          autor, el apellido se escribe una vez y los años van por orden.
        </p>
        <Ejemplo titulo="Ejemplo">
          <p>(Gil, 2019; Romero et al., 2021; Vidal y Ortega, 2018)</p>
          <p>(Romero, 2015, 2019, 2021)</p>
        </Ejemplo>
      </Apartado>

      <Apartado id="repetir" titulo="Citar varias veces la misma obra">
        <p>
          Cada cita va completa: APA no usa «ibid.» ni «op. cit.». En las citas narrativas, si nombras al mismo autor
          varias veces en un párrafo, el año solo hace falta la primera vez (y otra vez en cada párrafo nuevo).
        </p>
      </Apartado>

      <Apartado id="textuales" titulo="Citas textuales: cortas y en bloque">
        <p>
          Cuando copias las palabras exactas de una fuente, la cita lleva además la página: «p. 45» si es una, «pp.
          45–46» si son varias. Si la obra no tiene páginas, indica otra ubicación, como el párrafo.
        </p>
        <p>
          <strong>Menos de 40 palabras:</strong> entre comillas, dentro de tu propio texto, con la cita completa en la
          misma frase.
        </p>
        <Ejemplo titulo="Ejemplo (cita inventada)">
          <p>Según Romero (2021), «el sueño es la primera víctima de la época de exámenes» (p. 45).</p>
          <p>Los estudiantes recortan horas de descanso porque «el sueño es la primera víctima de la época de exámenes» (Romero, 2021, p. 45).</p>
        </Ejemplo>
        <p><strong>40 palabras o más:</strong> en un bloque aparte, con estas reglas:</p>
        <Lista>
          <li>sin comillas;</li>
          <li>empieza en una línea nueva, con 1,27 cm de sangría a la izquierda en todas sus líneas;</li>
          <li>a doble espacio, sin espacio extra antes ni después;</li>
          <li>la cita va después del punto final del bloque, y detrás del paréntesis no se pone punto.</li>
        </Lista>
        <Ejemplo titulo="Ejemplo (cita inventada)">
          <p>Romero (2021) describe así lo que ocurre en los meses de exámenes:</p>
          <p className="pl-8">
            Durante las semanas previas a los exámenes, la mayoría de los estudiantes reduce sus horas de sueño para
            ganar tiempo de estudio. Sin embargo, esa estrategia tiene el efecto contrario al buscado, porque el
            cansancio acumulado dificulta la concentración y empeora el recuerdo de lo estudiado. (p. 47)
          </p>
        </Ejemplo>
      </Apartado>
    </Guia>
  );
}
