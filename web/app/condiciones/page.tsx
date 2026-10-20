import type { Metadata } from "next";
import Link from "next/link";
import PaginaLegal, { Apartado } from "@/componentes/PaginaLegal";

export const metadata: Metadata = { title: "Condiciones de uso" };

export default function Condiciones() {
  return (
    <PaginaLegal titulo="Condiciones de uso" actualizada="19 de octubre de 2026">
      <Apartado titulo="Qué es APA Fácil">
        <p>
          Una herramienta gratuita que aplica a tu trabajo en Word el formato de la 7.ª edición de las normas
          APA para trabajos de estudiante, y te avisa de lo que debes revisar a mano. No hace falta registrarse.
        </p>
      </Apartado>
      <Apartado titulo="Lo que hace con tu texto">
        <p>
          Cambia el formato (márgenes, fuente, interlineado, sangrías, títulos, portada, tablas, referencias…),
          no lo que escribes. Los pocos retoques de texto que pide APA, como las mayúsculas de los títulos o
          cambiar «Bibliografía» por «Referencias», se indican en el informe y casi todos se pueden desactivar en
          «Más opciones». Nunca inventa datos: si falta algo en la portada o en una referencia, te lo señala.
        </p>
      </Apartado>
      <Apartado titulo="Revisa siempre el resultado">
        <p>
          La herramienta se ofrece tal cual, sin garantía de que el resultado cumpla todo lo que pida tu
          universidad o tu profesor, que pueden tener normas propias. Revisa la copia y el informe antes de
          entregar el trabajo: la responsabilidad del trabajo que entregas es tuya.
        </p>
      </Apartado>
      <Apartado titulo="Tu trabajo es tuyo">
        <p>
          No adquirimos ningún derecho sobre lo que subes y no lo guardamos (ver la{" "}
          <Link href="/privacidad" className="font-medium text-primario underline underline-offset-4">política de privacidad</Link>).
          Sube solo documentos que tengas derecho a usar.
        </p>
      </Apartado>
      <Apartado titulo="Uso razonable">
        <p>
          Se admiten documentos de Word (.docx) de hasta 10 MB, o hasta 500.000 caracteres de texto pegado, y un
          número limitado de documentos al día desde cada conexión (ahora, 3; solo cuentan los que se procesan
          bien). No intentes saltarte el límite, sobrecargar el servicio ni enviar archivos dañinos.
        </p>
      </Apartado>
      <Apartado titulo="Sin relación con la APA">
        <p>
          APA Fácil no está relacionada con la American Psychological Association ni cuenta con su aval. «APA» se
          usa solo para nombrar las normas que aplica.
        </p>
      </Apartado>
    </PaginaLegal>
  );
}
