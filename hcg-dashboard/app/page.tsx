"use client";

import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabase";
import type { DashboardStats, Equipo } from "@/lib/types";
import StatsPanel from "@/components/StatsPanel";
import EquipoTable from "@/components/EquipoTable";

export default function DashboardPage() {
  const [stats, setStats] = useState<DashboardStats>({
    total: 0, activos: 0, en_proceso: 0, baja: 0, online_24h: 0, criticos: 0,
  });
  const [recientes, setRecientes] = useState<Equipo[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function load() {
      const { data: equipos } = await supabase
        .from("equipos")
        .select("*")
        .order("updated_at", { ascending: false });

      if (equipos) {
        const now = Date.now();
        const h24 = 24 * 60 * 60 * 1000;
        setStats({
          total: equipos.length,
          activos: equipos.filter((e) => e.estado === "Activo").length,
          en_proceso: equipos.filter((e) => e.estado === "En proceso").length,
          baja: equipos.filter((e) => e.estado === "Baja").length,
          online_24h: equipos.filter((e) => e.ultima_conexion && now - new Date(e.ultima_conexion).getTime() < h24).length,
          criticos: 0,
        });
        setRecientes(equipos.slice(0, 10));
      }

      // Criticos
      const { data: diag } = await supabase
        .from("diagnostico_salud")
        .select("id")
        .eq("estado", "Critico");
      if (diag) {
        setStats((s) => ({ ...s, criticos: diag.length }));
      }

      setLoading(false);
    }
    load();
  }, []);

  if (loading) {
    return (
      <div className="text-center py-20" style={{ color: "var(--text-gold)" }}>
        &#10022; Elevando el cosmo... &#10022;
      </div>
    );
  }

  return (
    <div>
      <h2 className="text-xl font-bold mb-6" style={{ color: "var(--text-gold)" }}>
        ◆ Dashboard - Caballeros de Informatica
      </h2>
      <StatsPanel stats={stats} />
      <h3 className="text-lg font-bold mb-4" style={{ color: "var(--pegasus-cyan)" }}>
        ★ Actividad Reciente
      </h3>
      <EquipoTable
        equipos={recientes}
        onSelect={(eq) => (window.location.href = `/equipos/${eq.inv_st}`)}
      />
    </div>
  );
}
