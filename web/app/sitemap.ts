import type { MetadataRoute } from "next";
import { PAGINAS, URL_DEL_SITIO } from "@/lib/sitio";

// Paso 13.1: el mapa del sitio que se envía a Google Search Console.
export default function sitemap(): MetadataRoute.Sitemap {
  return PAGINAS.map(({ ruta, prioridad }) => ({
    url: `${URL_DEL_SITIO}${ruta === "/" ? "" : ruta}`,
    changeFrequency: "monthly",
    priority: prioridad,
  }));
}
