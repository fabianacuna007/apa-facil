import { GUIAS } from "@/lib/guias";

// Paso 13.1: datos del sitio para Google y las redes (títulos, dirección canónica, mapa del sitio).
// Al comprar el dominio (paso 13.3), basta con poner NEXT_PUBLIC_SITE_URL en Vercel (o el alojamiento nuevo).
export const URL_DEL_SITIO = (process.env.NEXT_PUBLIC_SITE_URL || "https://apa-facil-nine.vercel.app").replace(/\/$/, "");
export const NOMBRE_DEL_SITIO = "APA Fácil";
export const DESCRIPCION_DEL_SITIO =
  "Pon tu trabajo en formato APA 7 automáticamente y gratis: sube tu Word y descárgalo con márgenes, fuente, " +
  "interlineado, títulos, portada, tablas y referencias según las normas APA de estudiante. Sin registro.";

// Código de la etiqueta HTML de Google Search Console (Configuración › Verificación de la propiedad).
// Propiedad «Prefijo de URL» de https://apa-facil-nine.vercel.app (dada de alta por el propietario, 2026-10-20).
export const VERIFICACION_DE_GOOGLE = "zID9VcRisf-jHTVD5RYgtn5JNvbkPfP9sJ30Gvnhk8I";

// Las páginas que se dan a conocer a los buscadores (sitemap.xml), de la más a la menos importante.
export const PAGINAS: { ruta: string; prioridad: number }[] = [
  { ruta: "/", prioridad: 1 },
  { ruta: "/formatear", prioridad: 0.9 },
  { ruta: "/guias", prioridad: 0.8 },
  ...GUIAS.map((g) => ({ ruta: `/guias/${g.slug}`, prioridad: 0.7 })),   // paso 13.2
  { ruta: "/privacidad", prioridad: 0.2 },
  { ruta: "/condiciones", prioridad: 0.2 },
];
