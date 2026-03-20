"use client";

import type { Equipo } from "@/lib/types";
import StatusBadge from "./StatusBadge";

interface Props {
  equipos: Equipo[];
  onSelect?: (equipo: Equipo) => void;
}

export default function EquipoTable({ equipos, onSelect }: Props) {
  return (
    <div className="overflow-x-auto rounded-lg border" style={{ borderColor: "var(--border-dark)" }}>
      <table className="w-full text-sm">
        <thead>
          <tr style={{ background: "var(--gold-dark)" }}>
            {["#", "Inv.ST", "Fecha", "Marca", "Modelo", "Serie", "Procesador", "RAM", "Disco", "IP", "Estado", "FAA"].map(
              (h) => (
                <th
                  key={h}
                  className="px-3 py-2 text-left font-bold whitespace-nowrap"
                  style={{ color: "var(--text-gold)" }}
                >
                  {h}
                </th>
              )
            )}
          </tr>
        </thead>
        <tbody>
          {equipos.map((eq, i) => {
            let rowBg = i % 2 === 0 ? "var(--night)" : "var(--night-light)";
            if (eq.estado === "Activo") rowBg = "var(--activo-bg)";
            else if (eq.estado === "En proceso") rowBg = "var(--proceso-bg)";
            else if (eq.estado === "Baja") rowBg = "var(--baja-bg)";

            return (
              <tr
                key={eq.id}
                className="border-t cursor-pointer hover:opacity-80 transition-opacity"
                style={{ background: rowBg, borderColor: "var(--border-dark)" }}
                onClick={() => onSelect?.(eq)}
              >
                <td className="px-3 py-2">{i + 1}</td>
                <td className="px-3 py-2 font-bold" style={{ color: "var(--text-gold)" }}>
                  {eq.inv_st}
                </td>
                <td className="px-3 py-2 whitespace-nowrap">{eq.fecha_registro}</td>
                <td className="px-3 py-2">{eq.marca}</td>
                <td className="px-3 py-2 whitespace-nowrap">{eq.modelo}</td>
                <td className="px-3 py-2" style={{ color: "var(--text-cyan)" }}>
                  {eq.serie}
                </td>
                <td className="px-3 py-2 whitespace-nowrap">{eq.procesador}</td>
                <td className="px-3 py-2">{eq.ram}</td>
                <td className="px-3 py-2">{eq.disco}</td>
                <td className="px-3 py-2" style={{ color: "var(--pegasus-cyan)" }}>
                  {eq.ip_ethernet || "-"}
                </td>
                <td className="px-3 py-2">
                  <StatusBadge estado={eq.estado} />
                </td>
                <td className="px-3 py-2 text-xs">
                  {eq.faa?.startsWith("SI") ? (
                    <span style={{ color: "var(--activo-text)" }}>{eq.faa}</span>
                  ) : (
                    <span style={{ color: "var(--baja-text)" }}>{eq.faa || "-"}</span>
                  )}
                </td>
              </tr>
            );
          })}
          {equipos.length === 0 && (
            <tr>
              <td colSpan={12} className="text-center py-8 opacity-50">
                No se encontraron equipos
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}
