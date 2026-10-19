/**
 * Las opciones del formulario (paso 9.5): las mismas que la línea de comandos y
 * el servidor. tests/test_web.py comprueba que esta lista coincide con la del
 * servidor y la de fuentes con la del núcleo.
 */
import type { Opciones } from "@/lib/api";

/** Las fuentes recomendadas por APA, con su tamaño (apa_facil/reglas/fuente.py). */
export const FUENTES: [nombre: string, tamano: number][] = [
  ["Aptos", 12],
  ["Calibri", 11],
  ["Arial", 11],
  ["Lucida Sans Unicode", 10],
  ["Times New Roman", 12],
  ["Georgia", 11],
  ["Computer Modern", 10],
];

export type ClaveCasilla = Exclude<keyof Opciones, "fuente">;

export type Casilla = {
  clave: ClaveCasilla;
  /** Cómo se llama en la línea de comandos: el informe la nombra así y la web la cambia por la etiqueta. */
  bandera: string;
  etiqueta: string;
  ayuda: string;
  /** Las principales se ven siempre; el resto, en «Más opciones». */
  principal?: boolean;
};

export const CASILLAS: Casilla[] = [
  {
    clave: "mantener_portada", bandera: "--mantener-portada", principal: true,
    etiqueta: "Mantener la portada de mi universidad",
    ayuda: "Si rellenas la portada, sustituimos la que tenga tu trabajo por la portada APA. Márcalo si tu universidad exige la suya.",
  },
  {
    clave: "quitar_indice", bandera: "--quitar-indice", principal: true,
    etiqueta: "Quitar el índice",
    ayuda: "APA no usa índice, pero muchas universidades lo piden, así que por defecto se mantiene.",
  },
  {
    clave: "orden_tal_cual", bandera: "--orden-tal-cual", principal: true,
    etiqueta: "No ordenar las referencias",
    ayuda: "Por defecto se ordenan alfabéticamente, como pide APA.",
  },
  {
    clave: "citas_tal_cual", bandera: "--citas-tal-cual", principal: true,
    etiqueta: "No corregir las citas y referencias (solo avisar)",
    ayuda: "Por defecto corregimos su forma (la coma antes del año, «et al.», «s. f.», «y» entre autores, el punto tras la fecha…) y te lo contamos en el informe. Lo que obligaría a inventar datos solo se avisa.",
  },
  {
    clave: "quitar_numeracion", bandera: "--quitar-numeracion",
    etiqueta: "Quitar la numeración de los títulos",
    ayuda: "«1.1. Objetivos» pasa a «Objetivos». APA no numera los títulos, pero muchas universidades sí.",
  },
  {
    clave: "quitar_introduccion", bandera: "--quitar-introduccion",
    etiqueta: "Quitar el título «Introducción»",
    ayuda: "APA no lo usa: el título del trabajo, al principio del texto, hace sus veces.",
  },
  {
    clave: "titulos_tal_cual", bandera: "--titulos-tal-cual",
    etiqueta: "Dejar las mayúsculas de los títulos como están",
    ayuda: "Por defecto llevan mayúscula inicial en las palabras principales, como pide APA.",
  },
  {
    clave: "titulo_tal_cual", bandera: "--titulo-tal-cual",
    etiqueta: "Dejar las mayúsculas del título del trabajo como están",
    ayuda: "Lo mismo, solo para el título del trabajo.",
  },
  {
    clave: "tablas_tal_cual", bandera: "--tablas-tal-cual",
    etiqueta: "Dejar los bordes y colores de las tablas",
    ayuda: "Por defecto solo quedan las líneas horizontales que pide APA.",
  },
  {
    clave: "rotulo_tal_cual", bandera: "--rotulo-tal-cual",
    etiqueta: "Dejar el rótulo de la bibliografía como está",
    ayuda: "Por defecto «Bibliografía» o «Webgrafía» pasan a «Referencias», como pide APA.",
  },
  {
    clave: "rotulo_apendice", bandera: "--rotulo-apendice",
    etiqueta: "Cambiar «Anexo» por «Apéndice»",
    ayuda: "APA los llama apéndices, pero muchas universidades piden «Anexo», así que por defecto se deja.",
  },
  {
    clave: "comentarios", bandera: "--comentarios",
    etiqueta: "Añadir los avisos de citas y referencias como comentarios en el Word",
    ayuda: "Los verás en su sitio al abrir el documento. Bórralos antes de entregarlo.",
  },
];

/**
 * Los mensajes del informe están escritos para la línea de comandos («usa la
 * opción --quitar-indice»). En la web se nombra la casilla en su lugar. Solo
 * cambia cómo se muestra el mensaje: el informe no se toca.
 */
export function mensajeParaLaWeb(mensaje: string): string {
  let texto = mensaje
    .replace(/dalos en un archivo con --datos o con las opciones de portada/g, "rellénalos en «Portada»")
    .replace(/Añádelo en tu archivo de datos o con la opción --titulo\./g, "Añádelo en «Portada».")
    .replace(/añádelos en tu archivo de datos o con (la opción|las opciones) [^.]*\./g, "añádelos en «Portada».");
  for (const c of CASILLAS) {
    texto = texto.replaceAll(c.bandera, `«${c.etiqueta}»`);
  }
  return texto;
}
