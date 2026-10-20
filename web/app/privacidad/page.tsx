import type { Metadata } from "next";
import PaginaLegal, { Apartado } from "@/componentes/PaginaLegal";
import { CORREO_CONTACTO, RESPONSABLE } from "@/lib/legal";

export const metadata: Metadata = {
  title: "Privacidad",
  description: "Qué datos trata APA Fácil, para qué y durante cuánto tiempo: tu trabajo se borra en cuanto se procesa.",
  alternates: { canonical: "/privacidad" },
};

// Pasos 9.8, 11.2 y 11.3. Describe lo que hace de verdad el servidor (pasos 8.2, 8.3 y 10.5), el despliegue
// del 11.1 (Vercel, Cloudflare y Render), los avisos de errores (Sentry, servidor/errores.py) y la analítica
// (Vercel Web Analytics): si cambia algo de eso, hay que cambiar esto.
// Plazos comprobados en la documentación de cada proveedor el 20 de octubre de 2026.
export default function Privacidad() {
  const enlace = "font-medium text-primario underline underline-offset-4";
  return (
    <PaginaLegal titulo="Política de privacidad" actualizada="20 de octubre de 2026">
      <p>
        APA Fácil está pensada para que tu trabajo siga siendo solo tuyo. Esta página explica quién trata tus datos,
        cuáles, para qué y durante cuánto tiempo.
      </p>
      <Apartado titulo="Responsable">
        <p>
          {RESPONSABLE}. Para cualquier pregunta sobre tus datos, escribe a{" "}
          <a href={`mailto:${CORREO_CONTACTO}`} className={enlace}>{CORREO_CONTACTO}</a>.
        </p>
      </Apartado>
      <Apartado titulo="Qué datos tratamos">
        <ul className="list-disc space-y-1 pl-5">
          <li>El documento de Word que subes o el texto que pegas.</li>
          <li>Los datos de portada que escribas (título, autores, universidad, asignatura, profesor y fecha).</li>
          <li>Las opciones que elijas.</li>
          <li>
            Datos técnicos de cada petición: la dirección IP desde la que llega, la página pedida, el navegador y si ha
            ido bien (por ejemplo, «POST /procesar 200»). Nunca el nombre del archivo ni su contenido.
          </li>
          <li>
            Estadísticas de visitas: qué páginas se ven, desde qué país, con qué navegador y sistema, y desde qué web
            se llega. Son anónimas: no te identifican ni siguen lo que haces en otras webs.
          </li>
          <li>
            Si el servidor falla, el tipo de error y el sitio del programa donde ha ocurrido. Nunca tu documento, su
            nombre, tu IP ni nada de lo que has escrito.
          </li>
        </ul>
        <p>No hay cuentas de usuario: no te pedimos ni correo ni contraseña.</p>
        <p>
          Para el límite de documentos al día contamos cuántos procesas desde tu dirección IP. No guardamos la IP:
          solo una huella cifrada con una clave que cambia cada día, en la memoria del servidor, y se borra al
          acabar el día.
        </p>
      </Apartado>
      <Apartado titulo="Para qué">
        <p>
          Solo para aplicar el formato APA a tu trabajo y devolverte la copia y el informe. No usamos tu trabajo para
          nada más: ni para entrenar sistemas de inteligencia artificial, ni para publicidad, ni lo compartimos con
          nadie aparte de los proveedores que se indican abajo. La herramienta no usa inteligencia artificial: tu
          trabajo no se envía a ningún proveedor de IA.
        </p>
        <p>
          La base legal es la prestación del servicio que nos pides (artículo 6.1.b del Reglamento General de
          Protección de Datos). Los datos técnicos, el límite diario, los avisos de errores y las estadísticas de
          visitas se basan en nuestro interés legítimo en mantener el servicio seguro, corregir sus fallos y saber
          cuánto se usa (artículo 6.1.f).
        </p>
      </Apartado>
      <Apartado titulo="Proveedores que tratan tus datos">
        <p>Para funcionar, APA Fácil usa cuatro proveedores, que tratan los datos solo por cuenta nuestra:</p>
        <ul className="list-disc space-y-2 pl-5">
          <li>
            <strong className="text-tinta">Vercel Inc.</strong> (Estados Unidos) aloja la web, es decir, las páginas que
            ves, y cuenta las visitas. Recibe los datos técnicos de cada visita, pero no tu trabajo: el documento va
            directamente de tu navegador al servidor.
          </li>
          <li>
            <strong className="text-tinta">Cloudflare, Inc.</strong> (Estados Unidos) es la red por la que pasan las
            peticiones al servidor antes de llegar a Render. Transmite tu trabajo cifrado y los datos técnicos, sin
            guardar el trabajo.
          </li>
          <li>
            <strong className="text-tinta">Render Services, Inc.</strong> (Estados Unidos) aloja el servidor que da
            formato a tu trabajo, en un centro de datos de Fráncfort (Alemania, Unión Europea).
          </li>
          <li>
            <strong className="text-tinta">Functional Software, Inc.</strong> (Sentry, Estados Unidos) nos avisa de los
            errores del servidor, con los datos guardados en Fráncfort. Solo recibe el tipo de error y el sitio del
            programa donde ha ocurrido; antes de enviarlo, el servidor le quita todo lo demás.
          </li>
        </ul>
        <p>
          Las cuatro empresas son estadounidenses, así que puede haber transferencias de datos fuera de la Unión
          Europea. Vercel, Cloudflare y Sentry están adheridas al Marco de Privacidad de Datos UE-EE. UU., que la Comisión
          Europea considera adecuado, y Render se compromete en su contrato con las cláusulas contractuales tipo de
          la Comisión Europea.
        </p>
      </Apartado>
      <Apartado titulo="Cuánto tiempo">
        <p>
          Tu documento se guarda en una carpeta temporal del servidor solo mientras se procesa (unos segundos), y la
          carpeta se borra en cuanto termina, también si hay un error o se pasa del tiempo máximo. No guardamos copias
          ni del original ni del resultado. La copia corregida y el informe se envían a tu navegador, que es donde se
          descargan; al cerrar la página, desaparecen.
        </p>
        <p>
          Los datos técnicos quedan en los registros de los proveedores: 7 días en Render y 1 hora en Vercel.
          Cloudflare los conserva según su propia política de privacidad. Los avisos de errores se guardan 30 días en
          Sentry. Para contar las visitas, Vercel calcula una huella a partir de cada petición que descarta a las 24
          horas; las estadísticas que quedan son anónimas.
        </p>
      </Apartado>
      <Apartado titulo="Cookies">
        <p>
          No usamos cookies ni guardamos nada en tu navegador, así que no hay nada que aceptar. Las estadísticas de
          visitas (Vercel Web Analytics) funcionan sin cookies, y las fuentes de letra se sirven desde nuestra propia
          web, sin llamadas a Google.
        </p>
      </Apartado>
      <Apartado titulo="Tus derechos">
        <p>
          Puedes pedir el acceso, la rectificación o la supresión de tus datos, la limitación de su tratamiento, y
          oponerte a él, escribiendo a{" "}
          <a href={`mailto:${CORREO_CONTACTO}`} className={enlace}>{CORREO_CONTACTO}</a>. Como no guardamos tus
          documentos, en la práctica no hay nada que borrar después de procesarlos. Si crees que no hemos atendido
          bien tu petición, puedes reclamar ante la Agencia Española de Protección de Datos (aepd.es).
        </p>
      </Apartado>
    </PaginaLegal>
  );
}
