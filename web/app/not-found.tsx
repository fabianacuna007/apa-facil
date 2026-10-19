import { BotonEnlace } from "@/componentes/ui";

export default function NoEncontrada() {
  return (
    <main className="mx-auto flex w-full max-w-xl flex-1 flex-col items-center justify-center gap-4 px-4 py-20 text-center">
      <p className="font-titulo text-6xl font-extrabold text-primario">404</p>
      <h1 className="font-titulo text-2xl font-bold">Esta página no existe</h1>
      <p className="text-tinta-suave">Puede que el enlace esté mal escrito o que la página se haya movido.</p>
      <BotonEnlace href="/">Volver al inicio</BotonEnlace>
    </main>
  );
}
