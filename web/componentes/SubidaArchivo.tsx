"use client";

import { useId, useState } from "react";
import { FileText, FileUp, X } from "lucide-react";
import { Boton } from "@/componentes/ui";

export const TAMANO_MAXIMO_MB = 10; // el mismo que el servidor (servidor/seguridad.py)

/**
 * Paso 9.4: zona para arrastrar y soltar el Word, o elegirlo con el botón.
 * Antes de enviarlo comprueba lo que se puede comprobar en el navegador
 * (extensión, archivo vacío, tamaño) para avisar sin esperar al servidor, con
 * los mismos mensajes (documento.comprobar_nombre y servidor/seguridad.py).
 * El servidor lo vuelve a comprobar todo.
 */
export function comprobarArchivo(archivo: File): string | null {
  const nombre = archivo.name;
  if (nombre.startsWith("~$")) {
    return `«${nombre}» es un archivo temporal que Word crea mientras un documento está abierto. No es tu trabajo: elige el otro archivo.`;
  }
  const extension = nombre.includes(".") ? nombre.split(".").pop()!.toLowerCase() : "";
  if (extension === "doc") {
    return `«${nombre}» está en el formato antiguo de Word (.doc). Ábrelo en Word y usa «Guardar como» → «Documento de Word (.docx)».`;
  }
  if (extension !== "docx") {
    return `«${nombre}» no es un documento de Word. Solo se admiten archivos .docx.`;
  }
  if (archivo.size === 0) return `«${nombre}» está vacío (ocupa 0 bytes). Comprueba que has elegido el archivo bueno.`;
  if (archivo.size > TAMANO_MAXIMO_MB * 1024 * 1024) {
    return `«${nombre}» ocupa más de ${TAMANO_MAXIMO_MB} MB, el máximo que se admite. Si tiene muchas imágenes, ` +
      "comprímelas en Word (Formato de imagen › Comprimir imágenes) y vuelve a subirlo.";
  }
  return null;
}

function tamanoLegible(bytes: number) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1).replace(".", ",")} MB`;
}

export default function SubidaArchivo({ archivo, onChange, desactivada }: {
  archivo: File | null;
  onChange: (archivo: File | null) => void;
  desactivada?: boolean;
}) {
  const id = useId();
  const [encima, setEncima] = useState(false);
  const [error, setError] = useState<string | null>(null);

  function elegir(lista: FileList | null) {
    const elegido = lista?.[0];
    if (!elegido) return;
    if (lista.length > 1) {
      setError("Sube un solo documento cada vez.");
      return;
    }
    const problema = comprobarArchivo(elegido);
    setError(problema);
    onChange(problema ? null : elegido);
  }

  if (archivo) {
    return (
      <div className="flex items-center gap-3 rounded-xl border border-borde bg-hundido p-3.5">
        <span className="grid size-10 shrink-0 place-items-center rounded-lg bg-primario-suave text-primario" aria-hidden>
          <FileText className="size-5" />
        </span>
        <div className="min-w-0 flex-1">
          <p className="truncate text-sm font-semibold" title={archivo.name}>{archivo.name}</p>
          <p className="text-xs text-tinta-suave">{tamanoLegible(archivo.size)}</p>
        </div>
        <Boton
          variante="fantasma"
          tamano="pequeno"
          onClick={() => onChange(null)}
          disabled={desactivada}
          aria-label={`Quitar «${archivo.name}»`}
        >
          <X className="size-4" aria-hidden /> Quitar
        </Boton>
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-2">
      <label
        htmlFor={id}
        onDragOver={(e) => { e.preventDefault(); setEncima(true); }}
        onDragLeave={() => setEncima(false)}
        onDrop={(e) => { e.preventDefault(); setEncima(false); elegir(e.dataTransfer.files); }}
        className={[
          "flex cursor-pointer flex-col items-center gap-2 rounded-xl border-2 border-dashed px-4 py-8 text-center transition-colors",
          "has-[:focus-visible]:outline-2 has-[:focus-visible]:outline-offset-2 has-[:focus-visible]:outline-primario",
          encima ? "border-primario bg-primario-suave" : "border-borde-fuerte bg-hundido hover:border-primario/60",
        ].join(" ")}
      >
        <span className="grid size-11 place-items-center rounded-full bg-superficie text-primario shadow-sm" aria-hidden>
          <FileUp className="size-5" />
        </span>
        <span className="text-sm font-semibold">
          {/* En el móvil no se arrastra: se toca. */}
          <span className="sm:hidden">Toca aquí para <span className="text-primario underline underline-offset-4">elegir tu trabajo</span></span>
          <span className="hidden sm:inline">
            Arrastra aquí tu trabajo o <span className="text-primario underline underline-offset-4">elígelo</span>
          </span>
        </span>
        <span className="text-xs text-tinta-suave">Documento de Word (.docx), hasta {TAMANO_MAXIMO_MB} MB</span>
        <input
          id={id}
          type="file"
          accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          className="sr-only"
          disabled={desactivada}
          aria-describedby={error ? `${id}-error` : undefined}
          onChange={(e) => { elegir(e.target.files); e.target.value = ""; }}
        />
      </label>
      {error && <p id={`${id}-error`} role="alert" className="text-sm text-error">{error}</p>}
    </div>
  );
}
