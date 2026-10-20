import type { Metadata } from "next";
import Link from "next/link";
import { ArrowRight } from "lucide-react";
import { Tarjeta } from "@/componentes/ui";
import { GUIAS } from "@/lib/guias";

// Paso 13.2: el índice de las guías.
export const metadata: Metadata = {
  title: "Guías de APA 7 en español",
  description: "Guías breves y con ejemplos para aplicar las normas APA 7 de estudiante: formato en Word, portada, " +
    "títulos, tablas y figuras, citas, referencias y APA en el TFG y el TFM.",
  alternates: { canonical: "/guias" },
};

export default function Guias() {
  return (
    <main className="mx-auto flex w-full max-w-4xl flex-col gap-8 px-4 py-12 sm:px-6 sm:py-16">
      <header className="flex flex-col gap-3">
        <h1 className="font-titulo text-3xl font-extrabold tracking-tight sm:text-4xl">Guías de APA 7 en español</h1>
        <p className="max-w-2xl text-lg text-tinta-suave">
          Todo lo que pide la versión de estudiante de las normas APA, explicado en corto y con ejemplos. Basado en la
          web oficial de APA.
        </p>
      </header>
      <ul className="grid gap-4 sm:grid-cols-2">
        {GUIAS.map((g) => (
          <li key={g.slug}>
            <Link href={`/guias/${g.slug}`} className="group block h-full rounded-2xl">
              <Tarjeta className="flex h-full flex-col gap-2 p-6 transition-colors group-hover:border-primario">
                <h2 className="font-titulo text-lg font-bold text-tinta">{g.corto}</h2>
                <p className="text-sm text-tinta-suave">{g.descripcion}</p>
                <span className="mt-auto flex items-center gap-1 pt-2 text-sm font-semibold text-primario">
                  Leer la guía <ArrowRight className="size-4" aria-hidden />
                </span>
              </Tarjeta>
            </Link>
          </li>
        ))}
      </ul>
    </main>
  );
}
