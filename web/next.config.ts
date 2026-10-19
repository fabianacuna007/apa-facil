import type { NextConfig } from "next";

// Tailwind se carga con PostCSS (postcss.config.mjs), como indica la guía de Next.js.
// El cargador de Turbopack que traía la plantilla no recargaba globals.css al editarlo.
const nextConfig: NextConfig = {
  cacheComponents: true,
  partialPrefetching: true,
};

export default nextConfig;
