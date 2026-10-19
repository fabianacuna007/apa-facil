import EstadoServidor from "@/componentes/EstadoServidor";

// Página provisional del paso 9.1: la página de inicio de verdad es el paso 9.3.
export default function Inicio() {
  return (
    <main className="mx-auto flex w-full max-w-2xl flex-1 flex-col justify-center gap-4 px-4 py-16">
      <h1 className="text-3xl font-semibold tracking-tight">APA Fácil</h1>
      <p>Aplica el formato APA 7 de estudiante a tu trabajo en Word.</p>
      <EstadoServidor />
    </main>
  );
}
