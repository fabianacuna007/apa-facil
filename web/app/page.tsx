import { FileText } from "lucide-react";
import { Boton, Campo, Pastilla, Tarjeta, claseEntrada } from "@/componentes/ui";

// Muestra de la estética del paso 9.2. La página de inicio de verdad es el paso 9.3.
export default function Inicio() {
  return (
    <main className="mx-auto flex w-full max-w-2xl flex-col gap-6 px-4 py-16">
      <h1 className="skew-y-3 font-titulo text-4xl font-extrabold tracking-tight">APA Fácil</h1>
      <p className="text-tinta-suave">Aplica el formato APA 7 de estudiante a tu trabajo en Word.</p>
      <Tarjeta className="flex flex-col gap-4 p-6">
        <div className="flex gap-2"><Pastilla tono="exito">Hecho</Pastilla><Pastilla tono="aviso">Revisar</Pastilla><Pastilla tono="error">Error</Pastilla></div>
        <Campo id="t" etiqueta="Título del trabajo" ayuda="Tal cual lo quieres en la portada." obligatorio>
          <input id="t" className={claseEntrada} placeholder="p. ej., La revolución industrial en España" />
        </Campo>
        <div className="flex gap-3"><Boton><FileText className="size-4" />Primario</Boton><Boton variante="secundario">Secundario</Boton></div>
      </Tarjeta>
    </main>
  );
}
