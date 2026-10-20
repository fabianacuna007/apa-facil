import { ImageResponse } from "next/og";

// Paso 13.1: la imagen que se ve al compartir el enlace (WhatsApp, redes, Google Discover).
// Para las páginas que no tienen la suya; los colores son los de la web (globals.css).
export const alt = "APA Fácil: formato APA 7 automático para tu trabajo en Word";
export const size = { width: 1200, height: 630 };
export const contentType = "image/png";

export default function Imagen() {
  return new ImageResponse(
    (
      <div style={{ width: "100%", height: "100%", display: "flex", flexDirection: "column", justifyContent: "center",
                    padding: "80px", background: "#f6f8fb", color: "#0f172a" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "24px", fontSize: 44, fontWeight: 700 }}>
          <div style={{ width: 72, height: 72, borderRadius: 16, background: "#1d5cff", color: "white", display: "flex",
                        alignItems: "center", justifyContent: "center", fontSize: 40 }}>A</div>
          APA Fácil
        </div>
        <div style={{ marginTop: 48, fontSize: 72, fontWeight: 800, lineHeight: 1.1, maxWidth: 1000 }}>
          Tu trabajo en formato APA 7, sin pelearte con Word
        </div>
        <div style={{ marginTop: 32, fontSize: 34, color: "#475569" }}>
          Gratis y sin registro · No guardamos tu trabajo
        </div>
      </div>
    ),
    size,
  );
}
