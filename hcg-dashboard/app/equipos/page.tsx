"use client";

import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabase";
import type { Equipo } from "@/lib/types";
import EquipoTable from "@/components/EquipoTable";
import FilterBar from "@/components/FilterBar";

export default function EquiposPage() {
  const [equipos, setEquipos] = useState<Equipo[]>([]);
  const [filtered, setFiltered] = useState<Equipo[]>([]);
  const [search, setSearch] = useState("");
  const [estadoFilter, setEstadoFilter] = useState("");
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function load() {
      const { data } = await supabase
        .from("equipos")
        .select("*")
        .order("inv_st", { ascending: false });
      if (data) {
        setEquipos(data);
        setFiltered(data);
      }
      setLoading(false);
    }
    load();
  }, []);

  useEffect(() => {
    let result = equipos;

    if (estadoFilter) {
      result = result.filter((e) => e.estado === estadoFilter);
    }

    if (search) {
      const q = search.toLowerCase();
      result = result.filter(
        (e) =>
          e.inv_st?.toLowerCase().includes(q) ||
          e.serie?.toLowerCase().includes(q) ||
          e.modelo?.toLowerCase().includes(q) ||
          e.departamento?.toLowerCase().includes(q) ||
          e.usuario?.toLowerCase().includes(q) ||
          e.ip_ethernet?.includes(q)
      );
    }

    setFiltered(result);
  }, [search, estadoFilter, equipos]);

  if (loading) {
    return (
      <div className="text-center py-20" style={{ color: "var(--text-gold)" }}>
        &#10022; Cargando equipos... &#10022;
      </div>
    );
  }

  return (
    <div>
      <h2 className="text-xl font-bold mb-6" style={{ color: "var(--text-gold)" }}>
        ◆ Registro de Equipos ({filtered.length} de {equipos.length})
      </h2>
      <FilterBar
        search={search}
        onSearchChange={setSearch}
        estadoFilter={estadoFilter}
        onEstadoChange={setEstadoFilter}
      />
      <EquipoTable
        equipos={filtered}
        onSelect={(eq) => (window.location.href = `/equipos/${eq.inv_st}`)}
      />
    </div>
  );
}
