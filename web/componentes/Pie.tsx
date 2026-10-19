import Link from "next/link";

export default function Pie() {
  return (
    <footer className="mt-auto border-t border-borde bg-superficie">
      <div className="mx-auto flex max-w-6xl flex-col gap-3 px-4 py-6 text-sm text-tinta-suave sm:flex-row sm:items-center sm:justify-between sm:px-6">
        <p>APA Fácil no está relacionado con la American Psychological Association.</p>
        <nav aria-label="Información legal" className="flex gap-5">
          <Link href="/privacidad" className="underline-offset-4 hover:text-tinta hover:underline">Privacidad</Link>
          <Link href="/condiciones" className="underline-offset-4 hover:text-tinta hover:underline">Condiciones de uso</Link>
        </nav>
      </div>
    </footer>
  );
}
