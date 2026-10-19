/**
 * Cliente del servidor de APA Fácil (fase 8). Toda la web habla con el servidor
 * a través de este archivo: si cambia la API, solo se toca aquí.
 *
 * La dirección del servidor sale de NEXT_PUBLIC_API_URL (ver .env.example).
 * El servidor tiene que permitir el origen de la web (APA_ORIGENES, paso 9.1).
 */

export const API_URL = (process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000").replace(/\/+$/, "");

/** Las mismas opciones que la línea de comandos (servidor/app.py), como casillas. */
export type Opciones = {
  fuente?: string;
  quitar_numeracion?: boolean;
  titulos_tal_cual?: boolean;
  quitar_introduccion?: boolean;
  titulo_tal_cual?: boolean;
  mantener_portada?: boolean;
  quitar_indice?: boolean;
  tablas_tal_cual?: boolean;
  rotulo_tal_cual?: boolean;
  orden_tal_cual?: boolean;
  rotulo_apendice?: boolean;
  comentarios?: boolean;
};

/** Datos de la portada de estudiante. Lo que no se dé, no se rellena. */
export type DatosPortada = {
  titulo?: string;
  autor?: string[];
  departamento?: string;
  universidad?: string;
  codigo_asignatura?: string;
  asignatura?: string;
  profesor?: string[];
  fecha?: string;
};

/** Un caso concreto de un aviso de citas y referencias (fase 7): dónde está y cómo se corrige. */
export type Caso = {
  donde: string;
  extracto: string | null;
  correccion: string | null;
  explicacion: string | null;
  parrafo: number | null;
};

/** Una línea del informe: la regla de la que viene y el mensaje para el estudiante. */
export type Entrada = { regla: string; nombre: string; mensaje: string; casos?: Caso[] };

/** El informe (apa_facil/informe.py): lo que se ha cambiado, lo que hay que revisar y lo que no se pudo hacer. */
export type Informe = {
  archivo: string;
  copia: string;
  resumen: { cambios: number; avisos: number; errores: number };
  cambios: Entrada[];
  avisos: Entrada[];
  errores: Entrada[];
};

export type Resultado = {
  /** El informe completo, el mismo que guarda la línea de comandos. */
  informe: Informe;
  /** El informe en texto legible. */
  resumen: string;
  /** La copia corregida, con su contenido en base64. */
  documento: { nombre: string; contenido: string };
};

/** Un error con un mensaje en español pensado para el estudiante. */
export class ErrorApi extends Error {
  constructor(mensaje: string, readonly estado?: number) {
    super(mensaje);
    this.name = "ErrorApi";
  }
}

const SIN_CONEXION =
  "No se ha podido conectar con el servidor de APA Fácil. Comprueba tu conexión e inténtalo de nuevo.";

async function pedir(ruta: string, init?: RequestInit): Promise<unknown> {
  let respuesta: Response;
  try {
    respuesta = await fetch(`${API_URL}${ruta}`, init);
  } catch {
    throw new ErrorApi(SIN_CONEXION);
  }
  // El servidor siempre responde en JSON; los errores llevan {"error": "mensaje en español"}.
  const datos = await respuesta.json().catch(() => null);
  if (!respuesta.ok) {
    const mensaje = datos && typeof datos.error === "string" ? datos.error : SIN_CONEXION;
    throw new ErrorApi(mensaje, respuesta.status);
  }
  if (datos === null) throw new ErrorApi(SIN_CONEXION, respuesta.status);
  return datos;
}

/** true si el servidor está en marcha. */
export async function salud(): Promise<boolean> {
  try {
    const datos = (await pedir("/salud")) as { estado?: string };
    return datos.estado === "bien";
  } catch {
    return false;
  }
}

/** Envía el Word con sus opciones y datos de portada, y devuelve la copia y el informe. */
export async function procesar(
  archivo: File,
  opciones: Opciones = {},
  portada: DatosPortada = {},
): Promise<Resultado> {
  const formulario = new FormData();
  formulario.append("archivo", archivo);
  for (const [clave, valor] of Object.entries(opciones)) {
    if (valor === true) formulario.append(clave, "true");
    else if (typeof valor === "string" && valor.trim()) formulario.append(clave, valor);
  }
  for (const [clave, valor] of Object.entries(portada)) {
    // autor y profesor se repiten, uno por persona, como espera el servidor.
    for (const v of Array.isArray(valor) ? valor : [valor]) {
      if (typeof v === "string" && v.trim()) formulario.append(clave, v);
    }
  }
  return (await pedir("/procesar", { method: "POST", body: formulario })) as Resultado;
}

/** Convierte la copia en un archivo que el navegador puede descargar. */
export function archivoDescargable(documento: Resultado["documento"]): Blob {
  const bytes = Uint8Array.from(atob(documento.contenido), (c) => c.charCodeAt(0));
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
