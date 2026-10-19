"use client";

import { useEffect, useState } from "react";
import { API_URL, salud } from "@/lib/api";

/** Paso 9.1: comprueba desde el navegador que la web llega al servidor. */
export default function EstadoServidor() {
  const [estado, setEstado] = useState<"comprobando" | "conectado" | "sin-conexion">("comprobando");

  useEffect(() => {
    salud().then((bien) => setEstado(bien ? "conectado" : "sin-conexion"));
  }, []);

  const textos = {
    comprobando: "Comprobando la conexión con el servidor…",
    conectado: "Conectado con el servidor.",
    "sin-conexion": `No se puede conectar con el servidor (${API_URL}). ¿Está en marcha?`,
  };
  return (
    <p role="status" data-estado={estado} className="text-sm text-zinc-600 dark:text-zinc-400">
      {textos[estado]}
    </p>
  );
}
