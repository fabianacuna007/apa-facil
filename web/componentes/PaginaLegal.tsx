import { Tarjeta } from "@/componentes/ui";

/** Formato común de las páginas legales (pasos 9.8 y 11.2). */
export default function PaginaLegal({ titulo, actualizada, children }: {
  titulo: string; actualizada: string; children: React.ReactNode;
}) {
  return (
    <main className="mx-auto flex w-full max-w-3xl flex-col gap-6 px-4 py-12 sm:px-6 sm:py-16">
      <div className="flex flex-col gap-2">
        <h1 className="font-titulo text-3xl font-extrabold tracking-tight sm:text-4xl">{titulo}</h1>
        <p className="text-sm text-tinta-suave">Última actualización: {actualizada}</p>
      </div>
      <Tarjeta className="p-6 sm:p-8">
        <div className="flex flex-col gap-4 text-[15px] leading-relaxed text-tinta-suave">{children}</div>
      </Tarjeta>
    </main>
  );
}

export function Apartado({ titulo, children }: { titulo: string; children: React.ReactNode }) {
  return (
    <section className="flex flex-col gap-2">
      <h2 className="mt-2 font-titulo text-lg font-bold text-tinta">{titulo}</h2>
      {children}
    </section>
  );
}
