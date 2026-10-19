import { ArrowRight, ClipboardPaste, Download, FileUp, ListChecks, PenLine, ShieldCheck, TriangleAlert } from "lucide-react";
import { BotonEnlace, Pastilla, Tarjeta } from "@/componentes/ui";

const PASOS = [
  { icono: FileUp, titulo: "Sube tu trabajo", texto: "Tu archivo de Word (.docx) o, si lo prefieres, pega el texto." },
  { icono: PenLine, titulo: "Rellena la portada", texto: "Título, autores, asignatura, profesor y fecha. Lo que no pongas no se inventa." },
  { icono: Download, titulo: "Descárgalo en APA", texto: "Una copia nueva con el formato aplicado y un informe de lo que debes revisar." },
];

const CORRIGE = [
  "Márgenes, fuente, interlineado doble y sangrías",
  "Títulos de los cinco niveles y número de página",
  "Portada de estudiante, resumen y palabras clave",
  "Tablas, figuras, citas en bloque y apéndices",
  "Lista de referencias: rótulo, sangría francesa y orden",
];
const AVISA = [
  "Citas sin su referencia, y referencias que nadie cita",
  "Años, autores o puntuación que no siguen APA",
  "Datos que faltan en la portada o en las referencias",
];

// Paso 9.3: qué hace la herramienta, en una frase, y un botón para empezar.
export default function Inicio() {
  return (
    <main className="mx-auto flex w-full max-w-6xl flex-col gap-16 px-4 py-14 sm:px-6 sm:py-20">
      <section className="mx-auto flex max-w-3xl flex-col items-center gap-6 text-center">
        <Pastilla>APA 7 · versión de estudiante</Pastilla>
        <h1 className="font-titulo text-4xl font-extrabold tracking-tight text-balance sm:text-6xl">
          Tu trabajo en formato APA, sin pelearte con Word
        </h1>
        <p className="max-w-2xl text-lg text-pretty text-tinta-suave">
          Sube tu trabajo en Word y descárgalo con el formato APA 7 de estudiante aplicado. Solo cambiamos el
          formato: tu texto se queda tal cual.
        </p>
        <div className="flex flex-col gap-3 sm:flex-row">
          <BotonEnlace href="/formatear" tamano="grande">
            Empezar <ArrowRight className="size-5" aria-hidden />
          </BotonEnlace>
          <BotonEnlace href="/formatear?modo=texto" tamano="grande" variante="secundario">
            <ClipboardPaste className="size-5" aria-hidden /> Pegar texto
          </BotonEnlace>
        </div>
        <p className="flex items-center gap-1.5 text-sm text-tinta-suave">
          <ShieldCheck className="size-4 text-exito" aria-hidden />
          Gratis y sin registro. Tu archivo se borra en cuanto termina.
        </p>
      </section>

      <section aria-labelledby="como-funciona" className="flex flex-col gap-6">
        <h2 id="como-funciona" className="text-center font-titulo text-2xl font-bold">Cómo funciona</h2>
        <ol className="grid gap-4 md:grid-cols-3">
          {PASOS.map(({ icono: Icono, titulo, texto }, i) => (
            <li key={titulo}>
              <Tarjeta className="flex h-full flex-col gap-3 p-6">
                <span className="grid size-10 place-items-center rounded-xl bg-primario-suave text-primario" aria-hidden>
                  <Icono className="size-5" />
                </span>
                <h3 className="font-titulo text-lg font-bold">
                  <span className="text-tinta-tenue">{i + 1}.</span> {titulo}
                </h3>
                <p className="text-tinta-suave">{texto}</p>
              </Tarjeta>
            </li>
          ))}
        </ol>
      </section>

      <section aria-label="Qué hace" className="grid gap-4 md:grid-cols-2">
        <Tarjeta className="flex flex-col gap-4 p-6">
          <h2 className="flex items-center gap-2 font-titulo text-lg font-bold">
            <ListChecks className="size-5 text-exito" aria-hidden /> Lo que corregimos
          </h2>
          <ul className="flex flex-col gap-2 text-tinta-suave">
            {CORRIGE.map((c) => <li key={c} className="flex gap-2"><span aria-hidden className="text-exito">✓</span>{c}</li>)}
          </ul>
        </Tarjeta>
        <Tarjeta className="flex flex-col gap-4 p-6">
          <h2 className="flex items-center gap-2 font-titulo text-lg font-bold">
            <TriangleAlert className="size-5 text-aviso" aria-hidden /> Lo que te avisamos para que lo revises
          </h2>
          <ul className="flex flex-col gap-2 text-tinta-suave">
            {AVISA.map((c) => <li key={c} className="flex gap-2"><span aria-hidden className="text-aviso">!</span>{c}</li>)}
          </ul>
          <p className="text-sm text-tinta-suave">
            Nunca inventamos datos de una fuente: si falta algo, te lo señalamos.
          </p>
        </Tarjeta>
      </section>
    </main>
  );
}
