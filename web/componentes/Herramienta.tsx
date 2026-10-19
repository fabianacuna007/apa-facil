"use client";

import { useState } from "react";
import { Download, LoaderCircle, Sparkles } from "lucide-react";
import SubidaArchivo from "@/componentes/SubidaArchivo";
import { Boton, Tarjeta } from "@/componentes/ui";
import { archivoDescargable, ErrorApi, procesar, type Resultado } from "@/lib/api";

type Estado =
  | { fase: "inicial" }
  | { fase: "procesando" }
  | { fase: "hecho"; resultado: Resultado; enlace: string }
  | { fase: "error"; mensaje: string };

/** La herramienta: el trabajo, la portada y las opciones a la izquierda, y el resultado a la derecha. */
export default function Herramienta() {
  const [archivo, setArchivo] = useState<File | null>(null);
  const [estado, setEstado] = useState<Estado>({ fase: "inicial" });
  const procesando = estado.fase === "procesando";

  async function enviar(e: React.FormEvent) {
    e.preventDefault();
    if (!archivo) return;
    if (estado.fase === "hecho") URL.revokeObjectURL(estado.enlace);
    setEstado({ fase: "procesando" });
    try {
      const resultado = await procesar(archivo);
      setEstado({ fase: "hecho", resultado, enlace: URL.createObjectURL(archivoDescargable(resultado.documento)) });
    } catch (error) {
      setEstado({ fase: "error", mensaje: error instanceof ErrorApi ? error.message : "Ha habido un error inesperado." });
    }
  }

  return (
    <Tarjeta className="grid overflow-hidden lg:grid-cols-2">
      <form onSubmit={enviar} className="flex flex-col gap-6 p-5 sm:p-7 lg:border-r lg:border-borde">
        <h2 className="font-titulo text-xl font-bold">Tu trabajo</h2>
        <SubidaArchivo archivo={archivo} onChange={setArchivo} desactivada={procesando} />
        <Boton type="submit" tamano="grande" disabled={!archivo || procesando}>
          {procesando
            ? <><LoaderCircle className="size-5 animate-spin" aria-hidden /> Aplicando el formato…</>
            : <><Sparkles className="size-5" aria-hidden /> Aplicar formato APA</>}
        </Boton>
      </form>
      <section aria-live="polite" className="bg-hundido/60 p-5 sm:p-7">
        <h2 className="font-titulo text-xl font-bold">Resultado</h2>
        {estado.fase === "hecho" && (
          <a href={estado.enlace} download={estado.resultado.documento.nombre} className="mt-4 inline-flex items-center gap-2 font-semibold text-primario">
            <Download className="size-4" aria-hidden /> Descargar {estado.resultado.documento.nombre}
          </a>
        )}
        {estado.fase === "error" && <p role="alert" className="mt-4 text-error">{estado.mensaje}</p>}
      </section>
    </Tarjeta>
  );
}
