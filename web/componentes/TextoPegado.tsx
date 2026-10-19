"use client";

import { useId } from "react";
import { claseEntrada } from "@/componentes/ui";

export const MAXIMO_CARACTERES = 500_000; // el mismo que el núcleo (apa_facil/desde_texto.py)

/** Paso 9.7: el cuadro para pegar el trabajo como texto. */
export default function TextoPegado({ texto, onChange, desactivada }: {
  texto: string; onChange: (texto: string) => void; desactivada?: boolean;
}) {
  const id = useId();
  const largo = texto.length;
  const demasiado = largo > MAXIMO_CARACTERES;
  return (
    <div className="flex flex-col gap-2">
      <label htmlFor={id} className="text-sm font-semibold">Tu trabajo, pegado</label>
      <textarea
        id={id}
        value={texto}
        onChange={(e) => onChange(e.target.value)}
        disabled={desactivada}
        rows={12}
        spellCheck={false}
        aria-describedby={`${id}-ayuda ${id}-cuenta`}
        aria-invalid={demasiado}
        placeholder={"El título de tu trabajo\n\nIntroducción\nEl primer párrafo…"}
        className={`${claseEntrada} min-h-60 resize-y font-mono text-sm leading-relaxed`}
      />
      <div className="flex flex-col gap-1 text-xs text-tinta-suave sm:flex-row sm:justify-between sm:gap-4">
        <ul id={`${id}-ayuda`} className="flex list-disc flex-col gap-0.5 pl-4">
          <li>Empieza por el <strong>título del trabajo</strong>, sin la portada: sus datos van en «Portada».</li>
          <li>Cada línea será un párrafo. Los títulos se reconocen por su forma: líneas cortas sin punto final.</li>
          <li>Al pegar se pierden las negritas y cursivas: revisa las de las referencias.</li>
        </ul>
        <p id={`${id}-cuenta`} className={`shrink-0 tabular-nums ${demasiado ? "font-semibold text-error" : ""}`}>
          {largo.toLocaleString("es-ES")} / {MAXIMO_CARACTERES.toLocaleString("es-ES")} caracteres
        </p>
      </div>
    </div>
  );
}
