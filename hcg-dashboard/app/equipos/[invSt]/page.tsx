"use client";

import { useEffect, useState } from "react";
import { useParams } from "next/navigation";
import { supabase } from "@/lib/supabase";
import type { Equipo, InventarioSoftware, ReporteSistema, DiagnosticoSalud } from "@/lib/types";
import StatusBadge from "@/components/StatusBadge";

type Tab = "hardware" | "software" | "sistema" | "salud";

export default function EquipoDetailPage() {
  const params = useParams();
  const invSt = params.invSt as string;
  const [equipo, setEquipo] = useState<Equipo | null>(null);
  const [software, setSoftware] = useState<InventarioSoftware | null>(null);
  const [reporte, setReporte] = useState<ReporteSistema | null>(null);
  const [diagnostico, setDiagnostico] = useState<DiagnosticoSalud | null>(null);
  const [tab, setTab] = useState<Tab>("hardware");
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function load() {
      const [eqRes, swRes, rpRes, dgRes] = await Promise.all([
        supabase.from("equipos").select("*").eq("inv_st", invSt).single(),
        supabase.from("inventario_software").select("*").eq("inv_st", invSt).single(),
        supabase.from("reporte_sistema").select("*").eq("inv_st", invSt).single(),
        supabase.from("diagnostico_salud").select("*").eq("inv_st", invSt).single(),
      ]);
      setEquipo(eqRes.data);
      setSoftware(swRes.data);
      setReporte(rpRes.data);
      setDiagnostico(dgRes.data);
      setLoading(false);
    }
    load();
  }, [invSt]);

  if (loading) {
    return <div className="text-center py-20" style={{ color: "var(--text-gold)" }}>&#10022; Cargando... &#10022;</div>;
  }
  if (!equipo) {
    return <div className="text-center py-20" style={{ color: "var(--baja-text)" }}>Equipo no encontrado: {invSt}</div>;
  }

  const tabs: { key: Tab; label: string }[] = [
    { key: "hardware", label: "◆ Hardware" },
    { key: "software", label: "★ Software" },
    { key: "sistema", label: "◉ Sistema" },
    { key: "salud", label: "♥ Salud" },
  ];

  return (
    <div>
      <a href="/equipos" className="text-sm mb-4 inline-block hover:opacity-80" style={{ color: "var(--pegasus-cyan)" }}>
        ← Volver a equipos
      </a>

      <div className="flex items-center gap-4 mb-6">
        <h2 className="text-2xl font-bold" style={{ color: "var(--text-gold)" }}>
          Equipo {equipo.inv_st}
        </h2>
        <StatusBadge estado={equipo.estado} />
        {equipo.faa && (
          <span
            className="text-sm font-bold"
            style={{ color: equipo.faa.startsWith("SI") ? "var(--activo-text)" : "var(--baja-text)" }}
          >
            FAA: {equipo.faa}
          </span>
        )}
      </div>

      {/* Tabs */}
      <div className="flex gap-1 mb-6">
        {tabs.map((t) => (
          <button
            key={t.key}
            onClick={() => setTab(t.key)}
            className="px-4 py-2 rounded-t text-sm font-bold transition-colors"
            style={{
              background: tab === t.key ? "var(--gold-dark)" : "var(--night)",
              color: tab === t.key ? "var(--text-gold)" : "var(--text-light)",
              borderBottom: tab === t.key ? "2px solid var(--gold)" : "2px solid transparent",
            }}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div className="rounded-lg p-6 border" style={{ background: "var(--night)", borderColor: "var(--border-dark)" }}>
        {tab === "hardware" && <HardwareTab equipo={equipo} />}
        {tab === "software" && <SoftwareTab software={software} />}
        {tab === "sistema" && <SistemaTab reporte={reporte} />}
        {tab === "salud" && <SaludTab diagnostico={diagnostico} />}
      </div>
    </div>
  );
}

function Field({ label, value, color }: { label: string; value: string | number | null; color?: string }) {
  return (
    <div className="py-2">
      <span className="text-xs opacity-50">{label}</span>
      <div className="font-bold" style={{ color: color || "var(--text-light)" }}>
        {value || "-"}
      </div>
    </div>
  );
}

function HardwareTab({ equipo }: { equipo: Equipo }) {
  return (
    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
      <Field label="Inventario ST" value={equipo.inv_st} color="var(--text-gold)" />
      <Field label="Serie" value={equipo.serie} color="var(--text-cyan)" />
      <Field label="Marca" value={equipo.marca} />
      <Field label="Modelo" value={equipo.modelo} />
      <Field label="Procesador" value={equipo.procesador} />
      <Field label="Nucleos" value={equipo.nucleos} />
      <Field label="RAM" value={equipo.ram} />
      <Field label="Disco" value={equipo.disco} />
      <Field label="Graficos" value={equipo.graficos} />
      <Field label="WiFi" value={equipo.wifi} />
      <Field label="Bluetooth" value={equipo.bluetooth} />
      <Field label="S.O." value={equipo.sistema_operativo} />
      <Field label="MAC Ethernet" value={equipo.mac_ethernet} color="var(--pegasus-cyan)" />
      <Field label="MAC WiFi" value={equipo.mac_wifi} color="var(--pegasus-cyan)" />
      <Field label="IP Ethernet" value={equipo.ip_ethernet} color="var(--pegasus-cyan)" />
      <Field label="IP WiFi" value={equipo.ip_wifi} color="var(--pegasus-cyan)" />
      <Field label="Product Key" value={equipo.product_key} />
      <Field label="Fecha Registro" value={equipo.fecha_registro} />
      <Field label="Fabricacion" value={equipo.fecha_fabricacion} />
      <Field label="Garantia" value={equipo.garantia} />
      <Field label="Ubicacion" value={equipo.ubicacion} />
      <Field label="Departamento" value={equipo.departamento} />
      <Field label="Usuario" value={equipo.usuario} />
      <Field label="Red WiFi" value={equipo.red_wifi} />
    </div>
  );
}

function SoftwareTab({ software }: { software: InventarioSoftware | null }) {
  if (!software) return <div className="opacity-50">Sin datos de software registrados</div>;
  return (
    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
      <Field label="Nombre Equipo" value={software.nombre_equipo} />
      <Field label="Windows" value={software.windows_version} />
      <Field label="Build" value={software.windows_build} />
      <Field label="Activado" value={software.windows_activado} />
      <Field label="Office" value={software.office} />
      <Field label="Chrome" value={software.chrome} />
      <Field label="Acrobat" value={software.acrobat} />
      <Field label=".NET 3.5" value={software.dotnet35} />
      <Field label="Dedalus" value={software.dedalus} />
      <Field label="ESET" value={software.eset} />
      <Field label="WinRAR" value={software.winrar} />
      <Field label="Otro Software" value={software.otro_software} />
      <Field label="Usuario Windows" value={software.usuario_windows} />
      <Field label="Fecha Config" value={software.fecha_config} />
    </div>
  );
}

function SistemaTab({ reporte }: { reporte: ReporteSistema | null }) {
  if (!reporte) return <div className="opacity-50">Sin reporte de sistema</div>;
  return (
    <div className="space-y-4">
      <Field label="Nombre Equipo" value={reporte.nombre_equipo} />
      <Field label="Espacio Libre" value={reporte.espacio_libre_gb} color="var(--activo-text)" />
      <Field label="MB Limpiados" value={reporte.mb_limpiados} />
      <Field label="Fecha Reporte" value={reporte.fecha_reporte} />
      <div>
        <span className="text-xs opacity-50">Impresoras</span>
        <div className="text-sm mt-1">{reporte.impresoras?.split("|").join(", ") || "-"}</div>
      </div>
      <div>
        <span className="text-xs opacity-50">Usuarios</span>
        <div className="text-sm mt-1">{reporte.usuarios?.split("|").join(", ") || "-"}</div>
      </div>
      <div>
        <span className="text-xs opacity-50">Accesos Escritorio</span>
        <div className="text-sm mt-1">{reporte.accesos_escritorio?.split("|").join(", ") || "-"}</div>
      </div>
    </div>
  );
}

function SaludTab({ diagnostico }: { diagnostico: DiagnosticoSalud | null }) {
  if (!diagnostico) return <div className="opacity-50">Sin diagnostico de salud</div>;

  const ramColor = diagnostico.ram_pct > 85 ? "var(--baja-text)" : diagnostico.ram_pct > 70 ? "var(--proceso-text)" : "var(--activo-text)";
  const cpuColor = diagnostico.cpu_pct > 85 ? "var(--baja-text)" : diagnostico.cpu_pct > 70 ? "var(--proceso-text)" : "var(--activo-text)";

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="text-center p-4 rounded" style={{ background: "var(--night-dark)" }}>
          <div className="text-xs opacity-50 mb-1">RAM</div>
          <div className="text-3xl font-bold" style={{ color: ramColor }}>{diagnostico.ram_pct}%</div>
          <div className="text-xs mt-1">{diagnostico.ram_usada_gb}/{diagnostico.ram_total_gb} GB</div>
        </div>
        <div className="text-center p-4 rounded" style={{ background: "var(--night-dark)" }}>
          <div className="text-xs opacity-50 mb-1">CPU</div>
          <div className="text-3xl font-bold" style={{ color: cpuColor }}>{diagnostico.cpu_pct}%</div>
        </div>
        <div className="text-center p-4 rounded" style={{ background: "var(--night-dark)" }}>
          <div className="text-xs opacity-50 mb-1">Disco Libre</div>
          <div className="text-lg font-bold" style={{ color: "var(--pegasus-cyan)" }}>{diagnostico.disco_libre_gb}</div>
        </div>
        <div className="text-center p-4 rounded" style={{ background: "var(--night-dark)" }}>
          <div className="text-xs opacity-50 mb-1">Uptime</div>
          <div className="text-lg font-bold">{diagnostico.uptime_dias.toFixed(1)} dias</div>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-4">
        <Field label="Chrome (MB / procs)" value={`${diagnostico.chrome_mb} MB / ${diagnostico.chrome_procs}`} />
        <Field label="Dedalus (MB / procs)" value={`${diagnostico.dedalus_mb} MB / ${diagnostico.dedalus_procs}`} />
        <Field label="Total Procesos" value={diagnostico.total_procs} />
        <Field label="PageFile" value={`${diagnostico.pagefile_usado} / ${diagnostico.pagefile_total}`} />
      </div>

      {diagnostico.recomendacion && (
        <div className="p-3 rounded border" style={{ borderColor: "var(--gold-dark)", background: "var(--night-dark)" }}>
          <span className="text-xs font-bold" style={{ color: "var(--text-gold)" }}>Recomendacion: </span>
          <span className="text-sm">{diagnostico.recomendacion}</span>
        </div>
      )}

      <div>
        <span className="text-xs opacity-50">Top 5 Procesos</span>
        <div className="text-sm mt-1">{diagnostico.top5_procesos?.split("|").join(", ") || "-"}</div>
      </div>
    </div>
  );
}
