import type { Metadata, Viewport } from "next";
import { Inter, Plus_Jakarta_Sans } from "next/font/google";
import Cabecera from "@/componentes/Cabecera";
import Pie from "@/componentes/Pie";
import "./globals.css";

// Paso 9.2: Inter para el texto y Plus Jakarta Sans, más rotunda, para los títulos.
const inter = Inter({ variable: "--font-inter", subsets: ["latin"] });
const jakarta = Plus_Jakarta_Sans({ variable: "--font-jakarta", subsets: ["latin"], weight: ["600", "700", "800"] });

export const metadata: Metadata = {
  title: { default: "APA Fácil: formato APA 7 para tu trabajo en Word", template: "%s · APA Fácil" },
  description: "Sube tu trabajo en Word y descárgalo con el formato APA 7 de estudiante: márgenes, fuente, " +
    "interlineado, títulos, portada, tablas y referencias. Sin cambiar tu texto.",
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
      </body>
    </html>
  );
}
