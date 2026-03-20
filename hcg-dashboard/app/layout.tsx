import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "HCG Dashboard - Registro Cosmico de Equipos",
  description: "Caballeros de Informatica - Hospital Civil de Guadalajara",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="es">
      <body className="min-h-screen">
        <nav className="border-b" style={{ borderColor: "var(--gold-dark)", background: "var(--night)" }}>
          <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <span className="text-2xl">&#9733;</span>
              <h1 className="text-lg font-bold" style={{ color: "var(--text-gold)" }}>
                REGISTRO COSMICO DE EQUIPOS
              </h1>
              <span className="text-2xl">&#9733;</span>
            </div>
            <div className="flex gap-6 text-sm">
              <a href="/" className="hover:opacity-80" style={{ color: "var(--text-gold)" }}>
                Dashboard
              </a>
              <a href="/equipos" className="hover:opacity-80" style={{ color: "var(--pegasus-cyan)" }}>
                Equipos
              </a>
              <a href="/diagnostico" className="hover:opacity-80" style={{ color: "var(--activo-text)" }}>
                Diagnostico
              </a>
            </div>
          </div>
        </nav>
        <main className="max-w-7xl mx-auto px-4 py-6">{children}</main>
        <footer className="text-center py-4 text-xs" style={{ color: "var(--athena-purple)" }}>
          &#9733; Los Caballeros de Informatica protegen esta red &#9733; | &#10022; Enciende tu cosmo! &#10022;
        </footer>
      </body>
    </html>
  );
}
