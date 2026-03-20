"use client";

import type { DashboardStats } from "@/lib/types";

const CARDS = [
  { key: "total", label: "Total Equipos", color: "var(--text-gold)", icon: "◆" },
  { key: "activos", label: "Activos", color: "var(--activo-text)", icon: "★" },
  { key: "en_proceso", label: "En Proceso", color: "var(--proceso-text)", icon: "✦" },
  { key: "baja", label: "Baja", color: "var(--baja-text)", icon: "✖" },
  { key: "online_24h", label: "Online 24h", color: "var(--pegasus-cyan)", icon: "◉" },
  { key: "criticos", label: "Criticos", color: "#FF4444", icon: "⚠" },
] as const;

export default function StatsPanel({ stats }: { stats: DashboardStats }) {
  return (
    <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-8">
      {CARDS.map((card) => (
        <div
          key={card.key}
          className="rounded-lg p-4 text-center border"
          style={{ background: "var(--night)", borderColor: "var(--border-dark)" }}
        >
          <div className="text-2xl mb-1">{card.icon}</div>
          <div className="text-3xl font-bold" style={{ color: card.color }}>
            {stats[card.key as keyof DashboardStats]}
          </div>
          <div className="text-xs mt-1 opacity-70">{card.label}</div>
        </div>
      ))}
    </div>
  );
}
