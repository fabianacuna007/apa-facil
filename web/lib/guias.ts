// Paso 13.2: las guías de APA 7 (contenido propio, basado en la web oficial de APA, con ejemplos inventados).
// Cada una tiene su página en app/guias/<slug>/page.tsx; esta lista alimenta el índice, el mapa del sitio
// y los enlaces entre guías.

export type Guia = {
  slug: string;
  titulo: string;          // el de la página (y el de Google)
  corto: string;           // el de las tarjetas y los enlaces
  descripcion: string;     // la de Google y la de las tarjetas
  fuentes: { texto: string; url: string }[];   // las páginas de la web oficial de APA en que se basa
};

const APA = "https://apastyle.apa.org/style-grammar-guidelines";

export const GUIAS: Guia[] = [
  {
    slug: "formato-apa-word",
    titulo: "Cómo poner el formato APA 7 en Word, paso a paso",
    corto: "Formato APA 7 en Word",
    descripcion: "Márgenes, fuente, interlineado, alineación, sangría y número de página: qué pide APA 7 en un trabajo " +
      "de estudiante y dónde se cambia cada cosa en Word.",
    fuentes: [
      { texto: "Margins", url: `${APA}/paper-format/margins` },
      { texto: "Font", url: `${APA}/paper-format/font` },
      { texto: "Line spacing", url: `${APA}/paper-format/line-spacing` },
      { texto: "Paragraph alignment and indentation", url: `${APA}/paper-format/paragraph-format` },
      { texto: "Page header", url: `${APA}/paper-format/page-header` },
      { texto: "Order of pages", url: `${APA}/paper-format/order-pages` },
    ],
  },
  {
    slug: "portada-apa",
    titulo: "Portada APA 7 de estudiante: qué lleva y un ejemplo",
    corto: "Portada APA 7",
    descripcion: "Los elementos de la portada de estudiante en APA 7 (título, autores, universidad, asignatura, " +
      "profesor y fecha), en qué orden van y cómo se colocan, con un ejemplo.",
    fuentes: [{ texto: "Title page setup", url: `${APA}/paper-format/title-page` }],
  },
  {
    slug: "titulos-apa",
    titulo: "Títulos en APA 7: los cinco niveles, con ejemplos",
    corto: "Títulos y apartados",
    descripcion: "Cómo se escriben los títulos de los apartados en APA 7: los cinco niveles, sin números, en negrita " +
      "y con mayúsculas en las palabras principales. Con ejemplos y cómo hacerlo en Word.",
    fuentes: [{ texto: "Headings", url: `${APA}/paper-format/headings` }],
  },
  {
    slug: "tablas-figuras-apa",
    titulo: "Tablas y figuras en APA 7: número, título y nota",
    corto: "Tablas y figuras",
    descripcion: "Cómo se numeran y se titulan las tablas y las figuras en APA 7, qué bordes lleva una tabla, cuándo " +
      "se pone una nota y dónde se colocan en el trabajo.",
    fuentes: [
      { texto: "Table setup", url: `${APA}/tables-figures/tables` },
      { texto: "Figure setup", url: `${APA}/tables-figures/figures` },
    ],
  },
  {
    slug: "citas-apa",
    titulo: "Cómo citar en APA 7 dentro del texto",
    corto: "Citas en el texto",
    descripcion: "Citas entre paréntesis y narrativas, uno, dos o más autores («et al.»), varias obras a la vez, citas " +
      "textuales cortas y en bloque, con ejemplos en español.",
    fuentes: [
      { texto: "Author–date citation system", url: `${APA}/citations/basic-principles/author-date` },
      { texto: "Parenthetical versus narrative in-text citations",
        url: `${APA}/citations/basic-principles/parenthetical-versus-narrative` },
      { texto: "Citing multiple works", url: `${APA}/citations/basic-principles/multiple-works` },
      { texto: "Quotations", url: `${APA}/citations/quotations` },
    ],
  },
  {
    slug: "referencias-apa",
    titulo: "Lista de referencias en APA 7: formato y ejemplos",
    corto: "Lista de referencias",
    descripcion: "Cómo se monta la lista de referencias en APA 7 (rótulo, orden, sangría francesa, DOI) y ejemplos de " +
      "artículo, libro, capítulo de libro y página web.",
    fuentes: [
      { texto: "Reference list setup", url: `${APA}/paper-format/reference-list` },
      { texto: "Elements of reference list entries", url: `${APA}/references/elements-list-entry` },
      { texto: "Journal article references", url: `${APA}/references/examples/journal-article-references` },
      { texto: "Book/ebook references", url: `${APA}/references/examples/book-references` },
      { texto: "Chapter in an edited book/ebook references",
        url: `${APA}/references/examples/edited-book-chapter-references` },
      { texto: "Webpage on a website references", url: `${APA}/references/examples/webpage-website-references` },
      { texto: "DOIs and URLs", url: `${APA}/references/dois-urls` },
    ],
  },
  {
    slug: "apa-tfg-tfm",
    titulo: "Normas APA en el TFG y el TFM: qué pide APA y qué pide tu universidad",
    corto: "APA en el TFG y el TFM",
    descripcion: "Qué partes del formato APA 7 se aplican a un trabajo de fin de grado o de máster, y en qué suelen " +
      "mandar las normas de cada universidad: portada, índice, resumen y anexos.",
    fuentes: [
      { texto: "Order of pages", url: `${APA}/paper-format/order-pages` },
      { texto: "Margins", url: `${APA}/paper-format/margins` },
      { texto: "Title page setup", url: `${APA}/paper-format/title-page` },
    ],
  },
];

export function guia(slug: string): Guia {
  const encontrada = GUIAS.find((g) => g.slug === slug);
  if (!encontrada) throw new Error(`No existe la guía «${slug}»`);
  return encontrada;
}
