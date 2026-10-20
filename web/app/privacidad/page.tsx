import type { Metadata } from "next";
import PaginaLegal, { Apartado } from "@/componentes/PaginaLegal";

export const metadata: Metadata = { title: "Privacidad" };

// Paso 9.8. Describe lo que hace de verdad el servidor (pasos 8.2 y 8.3): si cambia, hay que cambiar esto.
export default function Privacidad() {
  return (
    <PaginaLegal titulo="Política de privacidad" actualizada="19 de octubre de 2026">
      <p>
        APA Fácil está pensada para que tu trabajo siga siendo solo tuyo. Esta página explica qué datos
        tratamos, para qué y durante cuánto tiempo.
      </p>
      <Apartado titulo="Qué datos tratamos">
        <ul className="list-disc space-y-1 pl-5">
          <li>El documento de Word que subes o el texto que pegas.</li>
          <li>Los datos de portada que escribas (título, autores, universidad, asignatura, profesor y fecha).</li>
          <li>Las opciones que elijas.</li>
          <li>
            Datos técnicos de cada petición: la dirección IP desde la que llega, la página pedida y si ha ido
            bien (por ejemplo, «POST /procesar 200»). Nunca el nombre del archivo ni su contenido.
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
          Solo para aplicar el formato APA a tu trabajo y devolverte la copia y el informe. No usamos tu trabajo
          para nada más: ni para entrenar sistemas de inteligencia artificial, ni para publicidad, ni lo
          compartimos con nadie. La base legal es la prestación del servicio que nos pides (artículo 6.1.b del
          Reglamento General de Protección de Datos).
        </p>
      </Apartado>
      <Apartado titulo="Cuánto tiempo">
        <p>
          Tu documento se guarda en una carpeta temporal del servidor solo mientras se procesa, y la carpeta se
          borra en cuanto termina, también si hay un error o se pasa del tiempo máximo. No guardamos copias ni
          del original ni del resultado. La copia corregida y el informe se envían a tu navegador, que es donde
          se descargan; al cerrar la página, desaparecen.
        </p>
        <p>Los datos técnicos de las peticiones se guardan en los registros del servidor durante el tiempo que fije el proveedor de alojamiento.</p>
      </Apartado>
      <Apartado titulo="Cookies">
        <p>No usamos cookies ni herramientas de analítica.</p>
      </Apartado>
      <Apartado titulo="Tus derechos">
        <p>
          Puedes pedir el acceso, la rectificación o la supresión de tus datos, y oponerte a su tratamiento.
          Como no guardamos tus documentos, en la práctica no hay nada que borrar después de procesarlos. También
          puedes reclamar ante la Agencia Española de Protección de Datos (aepd.es).
        </p>
      </Apartado>
    </PaginaLegal>
  );
}
