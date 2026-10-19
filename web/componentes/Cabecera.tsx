import Link from "next/link";
import { GraduationCap, ShieldCheck } from "lucide-react";
import { BotonEnlace, Pastilla } from "@/componentes/ui";

/** Barra superior de todas las páginas, como la de la imagen de referencia. */
export default function Cabecera() {
  return (
    <header className="sticky top-0 z-10 border-b border-borde bg-superficie/90 backdrop-blur">
      <div className="mx-auto flex h-16 max-w-6xl items-center gap-3 px-4 sm:px-6">
        <Link href="/" className="flex shrink-0 items-center gap-2.5 rounded-lg font-titulo text-lg font-bold whitespace-nowrap text-tinta">
          <span className="grid size-8 place-items-center rounded-lg bg-primario text-white" aria-hidden>
            <GraduationCap className="size-5" />
          </span>
          APA Fácil
        </Link>
        <span className="hidden sm:block">
          <Pastilla tono="exito">
            <ShieldCheck className="size-3.5" aria-hidden />
            No guardamos tu trabajo
          </Pastilla>
        </span>
        <nav aria-label="Principal" className="ml-auto">
          <BotonEnlace href="/formatear" tamano="pequeno" className="whitespace-nowrap">
            <span className="sm:hidden">Empezar</span>
            <span className="hidden sm:inline">Formatear mi trabajo</span>
          </BotonEnlace>
        </nav>
      </div>
    </header>
  );
}
