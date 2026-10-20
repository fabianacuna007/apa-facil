import type { Metadata } from "next";
import Link from "next/link";
import { ArrowRight, BookOpen } from "lucide-react";
import { BotonEnlace, Tarjeta } from "@/componentes/ui";
import { GUIAS, guia } from "@/lib/guias";
import { NOMBRE_DEL_SITIO, URL_DEL_SITIO } from "@/lib/sitio";

// Paso 13.2: el formato común de las guías de APA 7.

export const ACTUALIZADA = "20 de octubre de 2026";

export function metadatosDe(slug: string): Metadata {
  const g = guia(slug);
  return {
    title: g.titulo,
    description: g.descripcion,
    alternates: { canonical: `/guias/${slug}` },
    openGraph: { type: "article", title: g.titulo, description: g.descripcion, url: `/guias/${slug}` },
  };
}

export default function Guia({ slug, children }: { slug: string; children: React.ReactNode }) {
  const g = guia(slug);
  const otras = GUIAS.filter((o) => o.slug !== slug);
  const datos = {
    "@context": "https://schema.org",
    "@type": "Article",
    headline: g.titulo,
    description: g.descripcion,
    inLanguage: "es",
    url: `${URL_DEL_SITIO}/guias/${slug}`,
    dateModified: "2026-10-20",
    publisher: { "@type": "Organization", name: NOMBRE_DEL_SITIO, url: URL_DEL_SITIO },
  };
  return (
    <main className="mx-auto flex w-full max-w-3xl flex-col gap-8 px-4 py-12 sm:px-6 sm:py-16">
      <script type="application/ld+json"
              dangerouslySetInnerHTML={{ __html: JSON.stringify(datos).replace(/</g, "\\u003c") }} />
      <nav aria-label="Ruta" className="text-sm text-tinta-suave">
        <Link href="/guias" className="underline-offset-4 hover:text-tinta hover:underline">Guías de APA 7</Link>
        <span aria-hidden> › </span>
        <span>{g.corto}</span>
      </nav>
      <header className="flex flex-col gap-3">
        <h1 className="font-titulo text-3xl font-extrabold tracking-tight text-balance sm:text-4xl">{g.titulo}</h1>
        <p className="text-sm text-tinta-suave">Actualizada: {ACTUALIZADA} · APA 7.ª edición, versión de estudiante</p>
      </header>
      <article className="flex flex-col gap-5 text-[16px] leading-relaxed text-tinta-suave [&_strong]:text-tinta">
        {children}
      </article>
      <Tarjeta className="flex flex-col gap-3 border-primario/30 bg-primario-suave p-6">
        <h2 className="font-titulo text-lg font-bold text-tinta">Hazlo en un minuto con APA Fácil</h2>
        <p className="text-tinta-suave">
          Sube tu trabajo en Word y descárgalo con el formato APA 7 aplicado. Tu texto no se toca, y te avisamos de
          lo que debes revisar a mano. Gratis y sin registro.
        </p>
        <div>
          <BotonEnlace href="/formatear">Formatear mi trabajo <ArrowRight className="size-4" aria-hidden /></BotonEnlace>
        </div>
      </Tarjeta>
      <section className="flex flex-col gap-2 text-sm text-tinta-suave">
        <h2 className="font-titulo text-base font-bold text-tinta">Fuente</h2>
        <p>
          Esta guía resume, con palabras y ejemplos propios, lo que dice la web oficial de APA (en inglés):{" "}
          {g.fuentes.map((f, i) => (
            <span key={f.url}>
              <a href={f.url} className="font-medium text-primario underline underline-offset-4" rel="noopener">{f.texto}</a>
              {i < g.fuentes.length - 2 ? ", " : i === g.fuentes.length - 2 ? " y " : "."}
            </span>
          ))}{" "}
          Los ejemplos son inventados. Si tu profesor o tu universidad piden otra cosa, manda lo que pidan ellos.
        </p>
      </section>
      <section className="flex flex-col gap-3">
        <h2 className="flex items-center gap-2 font-titulo text-lg font-bold text-tinta">
          <BookOpen className="size-5 text-primario" aria-hidden /> Otras guías
        </h2>
        <ul className="grid gap-2 sm:grid-cols-2">
          {otras.map((o) => (
            <li key={o.slug}>
              <Link href={`/guias/${o.slug}`}
                    className="block rounded-xl border border-borde bg-superficie px-4 py-3 font-medium text-tinta hover:border-primario">
                {o.corto}
              </Link>
            </li>
          ))}
        </ul>
      </section>
    </main>
  );
}

/** Un apartado de la guía, con su ancla para enlazarlo. */
export function Apartado({ id, titulo, children }: { id: string; titulo: string; children: React.ReactNode }) {
  return (
    <section id={id} className="flex scroll-mt-24 flex-col gap-3">
      <h2 className="mt-3 font-titulo text-xl font-bold text-tinta">{titulo}</h2>
      {children}
    </section>
  );
}

/** Un ejemplo de cómo queda (inventado). */
export function Ejemplo({ titulo = "Ejemplo", children }: { titulo?: string; children: React.ReactNode }) {
  return (
    <figure className="rounded-xl border border-borde bg-hundido p-4">
      <figcaption className="mb-2 text-xs font-semibold tracking-wide text-tinta-tenue uppercase">{titulo}</figcaption>
      <div className="flex flex-col gap-2 font-[Georgia,serif] text-[15px] text-tinta">{children}</div>
    </figure>
  );
}

/** Una lista de puntos. */
export function Lista({ children, numerada = false }: { children: React.ReactNode; numerada?: boolean }) {
  const clase = `flex flex-col gap-2 pl-5 ${numerada ? "list-decimal" : "list-disc"}`;
  return numerada ? <ol className={clase}>{children}</ol> : <ul className={clase}>{children}</ul>;
}

/** Una referencia de ejemplo, con sangría francesa como en la lista de referencias. */
export function Referencia({ children }: { children: React.ReactNode }) {
  return <p className="pl-8 -indent-8">{children}</p>;
}
