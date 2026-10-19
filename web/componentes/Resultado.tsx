"use client";

import { useEffect, useRef } from "react";
import { ChevronDown, CircleAlert, CircleCheck, Download, FileText, LoaderCircle, RotateCcw, TriangleAlert } from "lucide-react";
import { Boton, Pastilla } from "@/componentes/ui";
import type { Caso, Entrada, Resultado as DatosResultado } from "@/lib/api";
import { mensajeParaLaWeb } from "@/lib/opciones";

export type Estado =
  | { fase: "inicial" }
  | { fase: "procesando"; nombre: string }
  | { fase: "hecho"; resultado: DatosResultado; enlace: string }
  | { fase: "error"; mensaje: string };

/**
 * Paso 9.6: el panel de la derecha. Mientras se procesa, una espera; al terminar,
 * el botón de descarga y el informe, en dos grupos: «Lo que debes revisar tú»
 * (avisos y errores, abierto) y «Lo que hemos cambiado» (cerrado, por si se quiere ver).
 */
export default function Resultado({ estado, onOtro }: { estado: Estado; onOtro: () => void }) {
  const titulo = useRef<HTMLHeadingElement>(null);

  // Al terminar, el foco va al resultado: en el móvil queda debajo del formulario.
  useEffect(() => {
    if (estado.fase === "hecho" || estado.fase === "error") {
      titulo.current?.focus();
      titulo.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }, [estado.fase]);

  return (
    <div className="flex flex-col gap-5">
      <h2 ref={titulo} tabIndex={-1} className="scroll-mt-24 font-titulo text-lg font-bold focus:outline-none">
        Resultado
      </h2>
      <div aria-live="polite" className="flex flex-col gap-5">
        {estado.fase === "inicial" && <Vacio />}
        {estado.fase === "procesando" && <Procesando nombre={estado.nombre} />}
        {estado.fase === "error" && <Fallo mensaje={estado.mensaje} />}
        {estado.fase === "hecho" && <Hecho resultado={estado.resultado} enlace={estado.enlace} onOtro={onOtro} />}
      </div>
    </div>
  );
}

function Vacio() {
  return (
    <div className="flex min-h-72 flex-col items-center justify-center gap-3 rounded-xl border border-dashed border-borde-fuerte px-6 py-12 text-center">
      <FileText className="size-10 text-tinta-tenue" strokeWidth={1.5} aria-hidden />
      <p className="font-semibold text-tinta-suave">Aquí aparecerá tu trabajo en formato APA</p>
      <p className="max-w-xs text-sm text-tinta-suave">
        Sube tu trabajo y pulsa «Aplicar formato APA». Te daremos una copia nueva y un informe de lo que debes revisar.
      </p>
    </div>
  );
}

function Procesando({ nombre }: { nombre: string }) {
  return (
    <div role="status" className="flex min-h-72 flex-col items-center justify-center gap-3 rounded-xl border border-borde bg-superficie px-6 py-12 text-center">
      <LoaderCircle className="size-10 animate-spin text-primario" aria-hidden />
      <p className="font-semibold">Aplicando el formato APA a «{nombre}»…</p>
      <p className="max-w-xs text-sm text-tinta-suave">
        Suele tardar unos segundos; un trabajo largo, hasta un minuto. No cierres esta página.
      </p>
    </div>
  );
}

function Fallo({ mensaje }: { mensaje: string }) {
  return (
    <div role="alert" className="flex gap-3 rounded-xl border border-error-borde bg-error-suave p-4 text-error">
      <CircleAlert className="mt-0.5 size-5 shrink-0" aria-hidden />
      <div className="flex flex-col gap-1">
        <p className="font-semibold">No se ha podido procesar el trabajo</p>
        <p className="text-sm">{mensaje}</p>
      </div>
    </div>
  );
}

function Hecho({ resultado, enlace, onOtro }: { resultado: DatosResultado; enlace: string; onOtro: () => void }) {
  const { informe, documento } = resultado;
  const revisar = [...informe.errores.map((e) => ({ ...e, error: true })), ...informe.avisos];
  return (
    <>
      <div className="flex flex-col gap-4 rounded-xl border border-exito-borde bg-exito-suave p-5">
        <div className="flex gap-3">
          <CircleCheck className="mt-0.5 size-6 shrink-0 text-exito" aria-hidden />
          <div className="flex min-w-0 flex-col gap-1">
            <p className="font-titulo text-lg font-bold">Tu trabajo ya tiene el formato APA</p>
            <p className="text-sm text-tinta-suave">
              Ábrelo en Word y repasa los avisos de abajo antes de entregarlo.
            </p>
          </div>
        </div>
        <a
          href={enlace}
          download={documento.nombre}
          className="inline-flex h-12 items-center justify-center gap-2 rounded-xl bg-primario px-5 font-semibold text-white shadow-sm transition-colors hover:bg-primario-hover"
        >
          <Download className="size-5" aria-hidden />
          <span className="truncate">Descargar {documento.nombre}</span>
        </a>
        <div className="flex flex-wrap gap-2">
          <Pastilla tono="exito">{contar(informe.resumen.cambios, "cambio", "cambios")}</Pastilla>
          <Pastilla tono={revisar.length ? "aviso" : "exito"}>
            {revisar.length ? contar(revisar.length, "cosa que revisar", "cosas que revisar") : "Nada que revisar"}
          </Pastilla>
        </div>
      </div>

      <Grupo
        titulo="Lo que debes revisar tú"
        icono={<TriangleAlert className="size-5 text-aviso" aria-hidden />}
        entradas={revisar}
        abierto
        vacio="No hemos encontrado nada que tengas que revisar a mano."
      />
      <Grupo
        titulo="Lo que hemos cambiado"
        icono={<CircleCheck className="size-5 text-exito" aria-hidden />}
        entradas={informe.cambios}
        vacio="Tu trabajo ya cumplía el formato: no ha hecho falta cambiar nada."
      />

      <Boton variante="secundario" onClick={onOtro} className="self-start">
        <RotateCcw className="size-4" aria-hidden /> Formatear otro trabajo
      </Boton>
    </>
  );
}

function contar(n: number, uno: string, varios: string) {
  return `${n} ${n === 1 ? uno : varios}`;
}

/** Un grupo del informe, desplegable, con las entradas ordenadas por la regla de la que vienen. */
function Grupo({ titulo, icono, entradas, abierto, vacio }: {
  titulo: string; icono: React.ReactNode; entradas: (Entrada & { error?: boolean })[]; abierto?: boolean; vacio: string;
}) {
  const porRegla = new Map<string, (Entrada & { error?: boolean })[]>();
  for (const e of entradas) porRegla.set(e.nombre, [...(porRegla.get(e.nombre) ?? []), e]);

  return (
    <details open={abierto} className="group rounded-xl border border-borde bg-superficie">
      <summary className="flex cursor-pointer list-none items-center gap-2.5 rounded-xl px-4 py-3.5 hover:bg-hundido">
        {icono}
        <span className="font-titulo font-bold">{titulo}</span>
        <span className="text-sm text-tinta-suave">({entradas.length})</span>
        <ChevronDown className="ml-auto size-4 transition-transform group-open:rotate-180" aria-hidden />
      </summary>
      <div className="flex flex-col gap-5 border-t border-borde px-4 py-4">
        {entradas.length === 0 && <p className="text-sm text-tinta-suave">{vacio}</p>}
        {[...porRegla].map(([regla, lista]) => (
          <section key={regla} className="flex flex-col gap-2">
            <h3 className="text-xs font-bold tracking-wide text-tinta-suave uppercase">{regla}</h3>
            <ul className="flex flex-col gap-2">
              {lista.map((e, i) => <Linea key={i} entrada={e} />)}
            </ul>
          </section>
        ))}
      </div>
    </details>
  );
}

function Linea({ entrada }: { entrada: Entrada & { error?: boolean } }) {
  return (
    <li className={[
      "flex flex-col gap-2 rounded-lg border-l-4 py-1.5 pl-3 text-sm",
      entrada.error ? "border-error" : "border-borde",
    ].join(" ")}>
      {entrada.error && <Pastilla tono="error" className="self-start">No se ha podido hacer</Pastilla>}
      <p>{mensajeParaLaWeb(entrada.mensaje)}</p>
      {entrada.casos && entrada.casos.length > 0 && (
        <ul className="flex flex-col gap-2">
          {entrada.casos.map((c, i) => <CasoConcreto key={i} caso={c} />)}
        </ul>
      )}
    </li>
  );
}

function CasoConcreto({ caso }: { caso: Caso }) {
  return (
    <li className="flex flex-col gap-1 rounded-lg bg-hundido px-3 py-2">
      <p className="text-tinta-suave">{caso.donde}</p>
      {caso.extracto && (
        <p>Has escrito: <span className="rounded bg-aviso-suave px-1 font-medium text-aviso">{caso.extracto}</span></p>
      )}
      {caso.correccion && (
        <p>Corrección: <span className="rounded bg-exito-suave px-1 font-medium text-exito">{caso.correccion}</span></p>
      )}
      {caso.explicacion && <p className="text-tinta-suave">{caso.explicacion}</p>}
    </li>
  );
}
