/**
 * Piezas básicas de la interfaz (paso 9.2): botones, tarjetas, pastillas,
 * campos de formulario, casillas y pestañas. Todas las pantallas se montan con
 * ellas, para que la web se vea igual en todas partes.
 */
import Link from "next/link";
import type { ComponentProps, ReactNode } from "react";

function unir(...clases: (string | false | null | undefined)[]) {
  return clases.filter(Boolean).join(" ");
}

// --- Botones ------------------------------------------------------------------------

const ESTILOS_BOTON = {
  primario: "bg-primario text-white hover:bg-primario-hover shadow-sm disabled:bg-primario/50",
  secundario: "bg-superficie text-tinta border border-borde-fuerte hover:bg-hundido disabled:text-tinta-tenue",
  fantasma: "text-tinta-suave hover:bg-hundido hover:text-tinta",
};
const TAMANOS_BOTON = {
  normal: "h-11 px-5 text-sm",
  grande: "h-13 px-7 text-base",
  pequeno: "h-9 px-3 text-sm",
};
type PropsBoton = {
  variante?: keyof typeof ESTILOS_BOTON;
  tamano?: keyof typeof TAMANOS_BOTON;
};

function claseBoton({ variante = "primario", tamano = "normal" }: PropsBoton, extra?: string) {
  return unir(
    "inline-flex items-center justify-center gap-2 rounded-xl font-semibold transition-colors",
    "disabled:cursor-not-allowed",
    ESTILOS_BOTON[variante],
    TAMANOS_BOTON[tamano],
    extra,
  );
}

export function Boton({ variante, tamano, className, ...resto }: PropsBoton & ComponentProps<"button">) {
  return <button type="button" className={claseBoton({ variante, tamano }, className)} {...resto} />;
}

export function BotonEnlace({ variante, tamano, className, ...resto }: PropsBoton & ComponentProps<typeof Link>) {
  return <Link className={claseBoton({ variante, tamano }, className)} {...resto} />;
}

// --- Tarjetas y pastillas --------------------------------------------------------------

export function Tarjeta({ className, ...resto }: ComponentProps<"div">) {
  return (
    <div className={unir("rounded-2xl border border-borde bg-superficie shadow-tarjeta", className)} {...resto} />
  );
}

const TONOS = {
  exito: "bg-exito-suave text-exito border-exito-borde",
  aviso: "bg-aviso-suave text-aviso border-aviso-borde",
  error: "bg-error-suave text-error border-error-borde",
  info: "bg-primario-suave text-primario border-primario/20",
};

export function Pastilla({ tono = "info", children, className }: {
  tono?: keyof typeof TONOS; children: ReactNode; className?: string;
}) {
  return (
    <span className={unir(
      "inline-flex items-center gap-1.5 rounded-full border px-2.5 py-0.5 text-xs font-semibold",
      TONOS[tono], className,
    )}>
      {children}
    </span>
  );
}

// --- Formularios ---------------------------------------------------------------------

export const claseEntrada = unir(
  "block w-full rounded-xl border border-borde-fuerte bg-superficie px-3.5 py-2.5 text-[15px] text-tinta",
  "placeholder:text-tinta-tenue hover:border-tinta-tenue transition-colors",
  "focus-visible:outline-2 focus-visible:outline-primario focus-visible:outline-offset-0 focus-visible:border-primario",
);

/** Un campo con su etiqueta (siempre visible, enlazada al control por `id`) y una ayuda opcional. */
export function Campo({ id, etiqueta, ayuda, obligatorio, children, className }: {
  id: string; etiqueta: string; ayuda?: ReactNode; obligatorio?: boolean; children: ReactNode; className?: string;
}) {
  return (
    <div className={unir("flex flex-col gap-1.5", className)}>
      <label htmlFor={id} className="text-sm font-semibold text-tinta">
        {etiqueta}
        {obligatorio && <span className="text-error" aria-hidden> *</span>}
      </label>
      {children}
      {ayuda && <p id={`${id}-ayuda`} className="text-xs text-tinta-tenue">{ayuda}</p>}
    </div>
  );
}

/** Una casilla con su texto y una explicación debajo. */
export function Casilla({ id, etiqueta, ayuda, checked, onChange }: {
  id: string; etiqueta: string; ayuda?: string; checked: boolean; onChange: (marcada: boolean) => void;
}) {
  return (
    <div className="flex gap-3">
      <input
        id={id}
        type="checkbox"
        checked={checked}
        onChange={(e) => onChange(e.target.checked)}
        aria-describedby={ayuda ? `${id}-ayuda` : undefined}
        className="mt-0.5 size-4.5 shrink-0 cursor-pointer rounded accent-primario"
      />
      <div className="flex flex-col gap-0.5">
        <label htmlFor={id} className="cursor-pointer text-sm font-medium text-tinta">{etiqueta}</label>
        {ayuda && <p id={`${id}-ayuda`} className="text-xs text-tinta-suave">{ayuda}</p>}
      </div>
    </div>
  );
}

/**
 * Pestañas segmentadas, como en la imagen de referencia. Se usan con las flechas
 * del teclado, como piden las pautas de accesibilidad para las pestañas.
 */
export function Pestanas<T extends string>({ opciones, valor, onChange, etiqueta, idPanel }: {
  opciones: { valor: T; texto: string; icono: ReactNode }[];
  valor: T;
  onChange: (valor: T) => void;
  etiqueta: string;
  idPanel: string;
}) {
  function teclas(e: React.KeyboardEvent, i: number) {
    const paso = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0;
    if (!paso) return;
    e.preventDefault();
    const siguiente = opciones[(i + paso + opciones.length) % opciones.length];
    onChange(siguiente.valor);
    document.getElementById(`pestana-${siguiente.valor}`)?.focus();
  }
  return (
    <div role="tablist" aria-label={etiqueta} className="grid gap-2" style={{ gridTemplateColumns: `repeat(${opciones.length}, 1fr)` }}>
      {opciones.map((o, i) => {
        const activa = o.valor === valor;
        return (
          <button
            key={o.valor}
            id={`pestana-${o.valor}`}
            type="button"
            role="tab"
            aria-selected={activa}
            aria-controls={idPanel}
            tabIndex={activa ? 0 : -1}
            onClick={() => onChange(o.valor)}
            onKeyDown={(e) => teclas(e, i)}
            className={unir(
              "flex flex-col items-center justify-center gap-1.5 rounded-xl border px-3 py-3 text-sm font-medium transition-colors",
              activa
                ? "border-primario bg-primario text-white shadow-sm"
                : "border-borde bg-superficie text-tinta-suave hover:border-borde-fuerte hover:text-tinta",
            )}
          >
            {o.icono}
            {o.texto}
          </button>
        );
      })}
    </div>
  );
}
