import type { MetadataRoute } from "next";
import { URL_DEL_SITIO } from "@/lib/sitio";

// Paso 13.1: todo se puede rastrear; el mapa del sitio dice qué páginas hay.
export default function robots(): MetadataRoute.Robots {
  return {
    rules: { userAgent: "*", allow: "/" },
    sitemap: `${URL_DEL_SITIO}/sitemap.xml`,
  };
}
