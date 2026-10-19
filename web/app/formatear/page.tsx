import type { Metadata } from "next";
import { Suspense } from "react";
import Herramienta from "@/componentes/Herramienta";

export const metadata: Metadata = { title: "Formatear mi trabajo" };

export default function Formatear() {
  return (
    <main className="mx-auto flex w-full max-w-6xl flex-col gap-8 px-4 py-10 sm:px-6 sm:py-14">
      <div className="flex flex-col gap-2 text-center">
        <h1 className="font-titulo text-3xl font-extrabold tracking-tight sm:text-5xl">Formatear mi trabajo</h1>
        <p className="mx-auto max-w-2xl text-tinta-suave">
          Sube tu trabajo, rellena la portada si quieres una nueva y descárgalo con el formato APA 7 de estudiante.
        </p>
      </div>
      {/* Herramienta lee ?modo=texto de la dirección, así que va dentro de Suspense. */}
      <Suspense>
        <Herramienta />
      </Suspense>
    </main>
  );
}
