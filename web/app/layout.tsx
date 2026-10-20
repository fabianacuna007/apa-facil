import type { Metadata, Viewport } from "next";
import { Analytics } from "@vercel/analytics/next";
import { Inter, Plus_Jakarta_Sans } from "next/font/google";
import Cabecera from "@/componentes/Cabecera";
import Pie from "@/componentes/Pie";
import { DESCRIPCION_DEL_SITIO, NOMBRE_DEL_SITIO, URL_DEL_SITIO, VERIFICACION_DE_GOOGLE } from "@/lib/sitio";
import "./globals.css";

// Paso 9.2: Inter para el texto y Plus Jakarta Sans, más rotunda, para los títulos.
const inter = Inter({ variable: "--font-inter", subsets: ["latin"] });
const jakarta = Plus_Jakarta_Sans({ variable: "--font-jakarta", subsets: ["latin"], weight: ["600", "700", "800"] });

// Paso 13.1: lo que leen Google y las redes. Cada página pone su título, su descripción y su dirección canónica.
export const metadata: Metadata = {
  metadataBase: new URL(URL_DEL_SITIO),
  title: { default: "Formato APA 7 automático y gratis para tu trabajo en Word · APA Fácil", template: "%s · APA Fácil" },
  description: DESCRIPCION_DEL_SITIO,
  applicationName: NOMBRE_DEL_SITIO,
  openGraph: { type: "website", locale: "es_ES", siteName: NOMBRE_DEL_SITIO },
  twitter: { card: "summary_large_image" },
  ...(VERIFICACION_DE_GOOGLE ? { verification: { google: VERIFICACION_DE_GOOGLE } } : {}),
};

export const viewport: Viewport = { themeColor: "#f6f8fb" };

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="es" className={`${inter.variable} ${jakarta.variable} h-full antialiased`}>
      <body className="flex min-h-full flex-col font-sans">
        <a
          href="#contenido"
          className="sr-only z-20 rounded-lg bg-primario px-4 py-2 font-semibold text-white focus:not-sr-only focus:fixed focus:left-4 focus:top-3"
        >
          Saltar al contenido
        </a>
        <Cabecera />
        <div id="contenido" className="flex flex-1 flex-col">{children}</div>
        <Pie />
        {/* Paso 11.3: visitas por página, sin cookies (ver /privacidad). Solo envía datos en Vercel. */}
        <Analytics />
      </body>
    </html>
  );
}
