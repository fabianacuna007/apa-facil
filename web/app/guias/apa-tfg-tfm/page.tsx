import Link from "next/link";
import Guia, { Apartado, Lista, metadatosDe } from "@/componentes/Guia";

const SLUG = "apa-tfg-tfm";
export const metadata = metadatosDe(SLUG);

const enlace = "font-medium text-primario underline underline-offset-4";

export default function Pagina() {
  return (
    <Guia slug={SLUG}>
      <p>
        Muchas universidades piden «normas APA» en el trabajo de fin de grado (TFG) o de máster (TFM). Pero APA está
        pensada para artículos y trabajos de clase, y la propia APA dice que las tesis y los trabajos de este tipo
        pueden necesitar partes que ella no recoge, y que en ese caso se siguen las normas de la universidad. En la
        práctica, casi siempre se combinan las dos cosas.
      </p>

      <Apartado id="regla" titulo="La regla de oro">
        <p>
          <strong>Lo que diga la guía de tu facultad manda sobre APA.</strong> APA se usa para todo lo que tu
          universidad no fija: normalmente, el formato del texto, los títulos, las tablas y figuras, las citas y las
          referencias. Busca la guía o el reglamento del TFG o TFM de tu titulación antes de empezar a dar formato.
        </p>
      </Apartado>

      <Apartado id="apa" titulo="Lo que suele seguir APA">
        <Lista>
          <li>
            El formato del texto: fuente, interlineado doble, alineación a la izquierda y sangría (ver{" "}
            <Link href="/guias/formato-apa-word" className={enlace}>formato APA 7 en Word</Link>).
          </li>
          <li>Los títulos de los apartados (ver <Link href="/guias/titulos-apa" className={enlace}>títulos y apartados</Link>).</li>
          <li>Las tablas y figuras (ver <Link href="/guias/tablas-figuras-apa" className={enlace}>tablas y figuras</Link>).</li>
          <li>
            Las citas y la lista de referencias (ver <Link href="/guias/citas-apa" className={enlace}>citas en el texto</Link>{" "}
            y <Link href="/guias/referencias-apa" className={enlace}>lista de referencias</Link>).
          </li>
        </Lista>
      </Apartado>

      <Apartado id="universidad" titulo="Lo que suele fijar la universidad">
        <Lista>
          <li>
            <strong>La portada:</strong> casi todas tienen su modelo, con el logo, la titulación, el tutor y la
            convocatoria. Se usa ese en lugar de la portada APA de estudiante.
          </li>
          <li>
            <strong>El índice:</strong> APA no lo incluye en los trabajos de estudiante, pero casi todos los TFG y TFM
            lo llevan. Hazlo con el índice automático de Word (<strong>Referencias</strong> › <strong>Tabla de
            contenido</strong>), que se actualiza solo y usa los estilos de título.
          </li>
          <li>
            <strong>El resumen y las palabras clave</strong>, a menudo también en inglés (<em>abstract</em> y{" "}
            <em>keywords</em>), con un número máximo de palabras.
          </li>
          <li>
            <strong>Los márgenes:</strong> APA pide 2,54 cm, pero algunas universidades piden más a la izquierda para
            encuadernar.
          </li>
          <li>
            <strong>La numeración de los apartados</strong> («1. Introducción», «2.1. Objetivos»): APA no numera los
            títulos, pero muchas guías de TFG lo piden.
          </li>
          <li>
            <strong>Los anexos:</strong> APA los llama «apéndices» y los rotula «Apéndice A», «Apéndice B»…; muchas
            universidades usan «Anexo I», «Anexo II»… Usa el rótulo que pida la tuya.
          </li>
          <li><strong>La extensión</strong>, normalmente en páginas o en palabras.</li>
        </Lista>
      </Apartado>

      <Apartado id="orden" titulo="Un orden de partes habitual">
        <p>
          Combinando APA y lo que suelen pedir las universidades, un TFG o TFM suele quedar así (comprueba el orden de
          la tuya):
        </p>
        <Lista numerada>
          <li>portada de la universidad;</li>
          <li>resumen y palabras clave (y su versión en inglés);</li>
          <li>índice (y, si hay muchas, índice de tablas y figuras);</li>
          <li>texto: introducción, objetivos, marco teórico, método, resultados, discusión y conclusiones;</li>
          <li>referencias;</li>
          <li>anexos.</li>
        </Lista>
      </Apartado>

      <Apartado id="apa-facil" titulo="Cómo te ayuda APA Fácil">
        <p>
          APA Fácil aplica el formato APA 7 de estudiante sin tocar tu texto, y tiene opciones pensadas para estos
          casos: si no rellenas los datos de la portada (o marcas «Mantener la portada de mi universidad»), tu portada
          se queda tal como está; el índice se mantiene, salvo que marques «Quitar el índice»; y los títulos numerados
          se dejan como están y se avisa, para que decidas tú. Al terminar, te da un informe con todo lo que ha
          cambiado y lo que debes revisar.
        </p>
      </Apartado>
    </Guia>
  );
}
