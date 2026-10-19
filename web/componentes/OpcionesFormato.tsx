"use client";

import { ChevronDown } from "lucide-react";
import { Campo, Casilla, claseEntrada } from "@/componentes/ui";
import type { Opciones } from "@/lib/api";
import { CASILLAS, FUENTES } from "@/lib/opciones";

/** Paso 9.5: la fuente y las casillas. Las principales se ven siempre; el resto, en «Más opciones». */
export default function OpcionesFormato({ opciones, onChange, desactivada }: {
  opciones: Opciones; onChange: (opciones: Opciones) => void; desactivada?: boolean;
}) {
  const casilla = (c: (typeof CASILLAS)[number]) => (
    <Casilla
      key={c.clave}
      id={`opcion-${c.clave}`}
      etiqueta={c.etiqueta}
      ayuda={c.ayuda}
      checked={opciones[c.clave] ?? false}
      onChange={(marcada) => onChange({ ...opciones, [c.clave]: marcada })}
    />
  );
  return (
    <fieldset className="flex flex-col gap-5" disabled={desactivada}>
      <Campo
        id="opcion-fuente"
        etiqueta="Fuente"
        ayuda="Si no eliges ninguna, se mantiene la de tu trabajo si APA la recomienda o, si no, la recomendada más parecida."
      >
        <select
          id="opcion-fuente"
          value={opciones.fuente ?? ""}
          onChange={(e) => onChange({ ...opciones, fuente: e.target.value || undefined })}
          aria-describedby="opcion-fuente-ayuda"
          className={claseEntrada}
        >
          <option value="">La de mi trabajo (o la más parecida)</option>
          {FUENTES.map(([nombre, tamano]) => (
            <option key={nombre} value={nombre}>{nombre} {tamano}</option>
          ))}
        </select>
      </Campo>
      {CASILLAS.filter((c) => c.principal).map(casilla)}
      <details className="group rounded-xl border border-borde">
        <summary className="flex cursor-pointer list-none items-center justify-between rounded-xl px-4 py-3 text-sm font-semibold hover:bg-hundido">
          Más opciones
          <ChevronDown className="size-4 transition-transform group-open:rotate-180" aria-hidden />
        </summary>
        <div className="flex flex-col gap-4 border-t border-borde px-4 py-4">
          {CASILLAS.filter((c) => !c.principal).map(casilla)}
        </div>
      </details>
    </fieldset>
  );
}
