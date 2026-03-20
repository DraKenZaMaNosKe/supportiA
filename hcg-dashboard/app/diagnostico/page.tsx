"use client";

import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabase";
import type { DiagnosticoSalud } from "@/lib/types";

export default function DiagnosticoPage() {
  const [diagnosticos, setDiagnosticos] = useState<DiagnosticoSalud[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function load() {
      const { data } = await supabase
        .from("diagnostico_salud")
        .select("*")
        .order("ram_pct", { ascending: false });
      if (data) setDiagnosticos(data);
      setLoading(false);
    }
    load();
  }, []);

  if (loading) {
    return <div className="text-center py-20" style={{ color: "var(--text-gold)" }}>&#10022; Cargando diagnosticos... &#10022;</div>;
  }

  const criticos = diagnosticos.filter((d) => d.estado === "Critico");
  const atencion = diagnosticos.filter((d) => d.estado === "Atencion");
  const ok = diagnosticos.filter((d) => d.estado === "OK");

  return (
    <div>
      <h2 className="text-xl font-bold mb-6" style={{ color: "var(--text-gold)" }}>
        ♥ Diagnostico de Salud ({diagnosticos.length} equipos)
      </h2>

      <div className="flex gap-4 mb-6 text-sm">
        <span style={{ color: "var(--baja-text)" }}>&#9888; Criticos: {criticos.length}</span>
        <span style={{ color: "var(--proceso-text)" }}>&#9888; Atencion: {atencion.length}</span>
        <span style={{ color: "var(--activo-text)" }}>&#10003; OK: {ok.length}</span>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {diagnosticos.map((d) => {
          let borderColor = "var(--activo-text)";
          if (d.estado === "Critico") borderColor = "var(--baja-text)";
          else if (d.estado === "Atencion") borderColor = "var(--proceso-text)";

          const ramColor = d.ram_pct > 85 ? "var(--baja-text)" : d.ram_pct > 70 ? "var(--proceso-text)" : "var(--activo-text)";
          const cpuColor = d.cpu_pct > 85 ? "var(--baja-text)" : d.cpu_pct > 70 ? "var(--proceso-text)" : "var(--activo-text)";

          return (
            <a
              key={d.id}
              href={`/equipos/${d.inv_st}`}
              className="rounded-lg p-4 border-l-4 hover:opacity-80 transition-opacity"
              style={{ background: "var(--night)", borderColor }}
            >
              <div className="flex justify-between items-center mb-3">
                <span className="font-bold" style={{ color: "var(--text-gold)" }}>{d.inv_st}</span>
                <span className="text-xs px-2 py-1 rounded" style={{ background: borderColor, color: "var(--night-dark)" }}>
                  {d.estado}
                </span>
              </div>
              <div className="text-xs opacity-70 mb-3">{d.nombre_equipo}</div>
              <div className="grid grid-cols-3 gap-2 text-center">
                <div>
                  <div className="text-xs opacity-50">RAM</div>
                  <div className="text-lg font-bold" style={{ color: ramColor }}>{d.ram_pct}%</div>
                </div>
                <div>
                  <div className="text-xs opacity-50">CPU</div>
                  <div className="text-lg font-bold" style={{ color: cpuColor }}>{d.cpu_pct}%</div>
                </div>
                <div>
                  <div className="text-xs opacity-50">Disco</div>
                  <div className="text-sm font-bold" style={{ color: "var(--pegasus-cyan)" }}>{d.disco_libre_gb}</div>
                </div>
              </div>
              {d.recomendacion && (
                <div className="text-xs mt-3 opacity-70 truncate">{d.recomendacion}</div>
              )}
            </a>
          );
        })}
      </div>

      {diagnosticos.length === 0 && (
        <div className="text-center py-12 opacity-50">No hay diagnosticos registrados</div>
      )}
    </div>
  );
}
