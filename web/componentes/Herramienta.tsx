"use client";

import { useState } from "react";
import { useSearchParams } from "next/navigation";
import { ClipboardPaste, FileUp, LoaderCircle, Sparkles } from "lucide-react";
import OpcionesFormato from "@/componentes/OpcionesFormato";
import Portada from "@/componentes/Portada";
import Resultado, { type Estado } from "@/componentes/Resultado";
import SubidaArchivo from "@/componentes/SubidaArchivo";
import TextoPegado, { MAXIMO_CARACTERES } from "@/componentes/TextoPegado";
import { Boton, Pestanas, Tarjeta } from "@/componentes/ui";
import { archivoDescargable, ErrorApi, procesar, procesarTexto, type DatosPortada, type Opciones } from "@/lib/api";

type Modo = "archivo" | "texto";

/** La herramienta: el trabajo, la portada y las opciones a la izquierda, y el resultado a la derecha. */
export default function Herramienta() {
  // «Pegar texto» en la página de inicio lleva a /formatear?modo=texto.
  const [modo, setModo] = useState<Modo>(useSearchParams().get("modo") === "texto" ? "texto" : "archivo");
  const [archivo, setArchivo] = useState<File | null>(null);
  const [texto, setTexto] = useState("");
  const [portada, setPortada] = useState<DatosPortada>({});
  const [opciones, setOpciones] = useState<Opciones>({});
  const [estado, setEstado] = useState<Estado>({ fase: "inicial" });
  const procesando = estado.fase === "procesando";
  const listo = modo === "archivo" ? archivo !== null : texto.trim() !== "" && texto.length <= MAXIMO_CARACTERES;

  /** «Formatear otro trabajo»: se quita el archivo y el resultado; la portada y las opciones se quedan. */
  function otro() {
    if (estado.fase === "hecho") URL.revokeObjectURL(estado.enlace);
    setArchivo(null);
    setTexto("");
    setEstado({ fase: "inicial" });
    document.getElementById("seccion-1")?.scrollIntoView({ behavior: "smooth" });
  }

  async function enviar(e: React.FormEvent) {
    e.preventDefault();
    if (!listo || procesando) return;
    if (estado.fase === "hecho") URL.revokeObjectURL(estado.enlace);
    setEstado({ fase: "procesando", nombre: modo === "archivo" && archivo ? archivo.name : "tu texto" });
    try {
      const resultado = modo === "archivo" && archivo
        ? await procesar(archivo, opciones, portada)
        : await procesarTexto(texto, opciones, portada);
      setEstado({ fase: "hecho", resultado, enlace: URL.createObjectURL(archivoDescargable(resultado.documento)) });
    } catch (error) {
      setEstado({ fase: "error", mensaje: error instanceof ErrorApi ? error.message : "Ha habido un error inesperado." });
    }
  }

  return (
    <Tarjeta className="grid overflow-hidden lg:grid-cols-2">
      <form onSubmit={enviar} className="flex flex-col gap-6 p-5 sm:p-7 lg:border-r lg:border-borde">
        <Seccion numero={1} titulo="Tu trabajo">
          <Pestanas
            etiqueta="Cómo quieres darnos tu trabajo"
            idPanel="panel-trabajo"
            valor={modo}
            onChange={setModo}
            opciones={[
              { valor: "archivo", texto: "Subir Word", icono: <FileUp className="size-5" aria-hidden /> },
              { valor: "texto", texto: "Pegar texto", icono: <ClipboardPaste className="size-5" aria-hidden /> },
            ]}
          />
          <div id="panel-trabajo" role="tabpanel" aria-labelledby={`pestana-${modo}`}>
            {modo === "archivo"
              ? <SubidaArchivo archivo={archivo} onChange={setArchivo} desactivada={procesando} />
              : <TextoPegado texto={texto} onChange={setTexto} desactivada={procesando} />}
          </div>
        </Seccion>
        <Seccion
          numero={2}
          titulo="Portada"
          detalle={modo === "archivo"
            ? "Opcional. Con estos datos creamos la portada APA de estudiante; lo que dejes vacío no se rellena. Sin datos, se mantiene la portada que tenga tu trabajo."
            : "Rellénala para que tu trabajo lleve la portada APA de estudiante; lo que dejes vacío no se rellena. Sin datos, la primera página llevará solo el título."}
        >
          <Portada datos={portada} onChange={setPortada} desactivada={procesando} />
        </Seccion>
        <Seccion numero={3} titulo="Opciones">
          <OpcionesFormato opciones={opciones} onChange={setOpciones} desactivada={procesando} />
        </Seccion>
        <Boton type="submit" tamano="grande" disabled={!listo || procesando}>
          {procesando
            ? <><LoaderCircle className="size-5 animate-spin" aria-hidden /> Aplicando el formato…</>
            : <><Sparkles className="size-5" aria-hidden /> Aplicar formato APA</>}
        </Boton>
      </form>
      <div className="bg-hundido/60 p-5 sm:p-7">
        <div className="lg:sticky lg:top-24">
          <Resultado estado={estado} onOtro={otro} />
        </div>
      </div>
    </Tarjeta>
  );
}

function Seccion({ numero, titulo, detalle, children }: {
  numero: number; titulo: string; detalle?: string; children: React.ReactNode;
}) {
  return (
    <section aria-labelledby={`seccion-${numero}`} className="flex flex-col gap-4">
      <div className="flex flex-col gap-1">
        <h2 id={`seccion-${numero}`} className="flex items-center gap-2.5 font-titulo text-lg font-bold">
          <span className="grid size-7 place-items-center rounded-full bg-primario-suave text-sm text-primario" aria-hidden>
            {numero}
          </span>
          {titulo}
        </h2>
        {detalle && <p className="text-sm text-tinta-suave">{detalle}</p>}
      </div>
      {children}
    </section>
  );
}
