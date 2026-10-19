"use client";

import { Plus, X } from "lucide-react";
import { Boton, Campo, claseEntrada } from "@/componentes/ui";
import type { DatosPortada } from "@/lib/api";

/** Paso 9.5: los datos de la portada de estudiante. Se envían tal cual se escriben. */
export default function Portada({ datos, onChange, desactivada }: {
  datos: DatosPortada; onChange: (datos: DatosPortada) => void; desactivada?: boolean;
}) {
  const poner = (clave: keyof DatosPortada) => (valor: string | string[]) => onChange({ ...datos, [clave]: valor });
  const texto = (clave: "titulo" | "departamento" | "universidad" | "codigo_asignatura" | "asignatura") => ({
    id: `portada-${clave}`,
    value: datos[clave] ?? "",
    onChange: (e: React.ChangeEvent<HTMLInputElement>) => poner(clave)(e.target.value),
    className: claseEntrada,
    disabled: desactivada,
  });

  return (
    <fieldset className="flex flex-col gap-4" disabled={desactivada}>
      <Campo id="portada-titulo" etiqueta="Título del trabajo">
        <input {...texto("titulo")} placeholder="p. ej., La revolución industrial en España" />
      </Campo>
      <ListaDeNombres
        id="portada-autor"
        etiqueta="Autor o autores"
        singular="autor"
        nombres={datos.autor ?? [""]}
        onChange={poner("autor")}
        ejemplo="p. ej., Ana Pérez García"
      />
      <div className="grid gap-4 sm:grid-cols-2">
        <Campo id="portada-departamento" etiqueta="Departamento o facultad">
          <input {...texto("departamento")} placeholder="p. ej., Departamento de Historia" />
        </Campo>
        <Campo id="portada-universidad" etiqueta="Universidad">
          <input {...texto("universidad")} placeholder="p. ej., Universidad de Salamanca" />
        </Campo>
      </div>
      <div className="grid gap-4 sm:grid-cols-[9rem_minmax(0,1fr)]">
        <Campo id="portada-codigo_asignatura" etiqueta="Código">
          <input {...texto("codigo_asignatura")} placeholder="p. ej., HIS 201" />
        </Campo>
        <Campo id="portada-asignatura" etiqueta="Asignatura">
          <input {...texto("asignatura")} placeholder="p. ej., Historia Contemporánea" />
        </Campo>
      </div>
      <ListaDeNombres
        id="portada-profesor"
        etiqueta="Profesor o profesora"
        singular="profesor"
        nombres={datos.profesor ?? [""]}
        onChange={poner("profesor")}
        ejemplo="p. ej., Dra. Marta López"
      />
      <Campo id="portada-fecha" etiqueta="Fecha de entrega" className="sm:max-w-xs">
        <input
          id="portada-fecha"
          type="date"
          value={datos.fecha ?? ""}
          onChange={(e) => poner("fecha")(e.target.value)}
          className={claseEntrada}
        />
      </Campo>
    </fieldset>
  );
}

/** Una lista de nombres (autores o profesores): un campo por persona, con botones para añadir y quitar. */
function ListaDeNombres({ id, etiqueta, singular, nombres, onChange, ejemplo }: {
  id: string; etiqueta: string; singular: string; nombres: string[]; onChange: (n: string[]) => void; ejemplo: string;
}) {
  return (
    <div className="flex flex-col gap-1.5">
      <label htmlFor={`${id}-0`} className="text-sm font-semibold">{etiqueta}</label>
      {nombres.map((nombre, i) => (
        <div key={i} className="flex gap-2">
          <input
            id={`${id}-${i}`}
            value={nombre}
            onChange={(e) => onChange(nombres.map((n, j) => (j === i ? e.target.value : n)))}
            className={claseEntrada}
            placeholder={i === 0 ? ejemplo : undefined}
            aria-label={i > 0 ? `${etiqueta} (${i + 1})` : undefined}
          />
          {nombres.length > 1 && (
            <Boton
              variante="fantasma"
              tamano="pequeno"
              className="h-auto shrink-0"
              onClick={() => onChange(nombres.filter((_, j) => j !== i))}
              aria-label={`Quitar ${singular} ${i + 1}`}
            >
              <X className="size-4" aria-hidden />
            </Boton>
          )}
        </div>
      ))}
      <Boton variante="fantasma" tamano="pequeno" className="self-start" onClick={() => onChange([...nombres, ""])}>
        <Plus className="size-4" aria-hidden /> Añadir otro {singular}
      </Boton>
    </div>
  );
}
