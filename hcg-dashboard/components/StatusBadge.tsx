"use client";

const STATUS_STYLES: Record<string, { bg: string; text: string; label: string }> = {
  Activo: { bg: "var(--activo-bg)", text: "var(--activo-text)", label: "★ Activo" },
  "En proceso": { bg: "var(--proceso-bg)", text: "var(--proceso-text)", label: "✦ En proceso" },
  Baja: { bg: "var(--baja-bg)", text: "var(--baja-text)", label: "✖ Baja" },
};

export default function StatusBadge({ estado }: { estado: string }) {
  const style = STATUS_STYLES[estado] || STATUS_STYLES["En proceso"];
  return (
    <span
      className="px-2 py-1 rounded text-xs font-bold whitespace-nowrap"
      style={{ background: style.bg, color: style.text }}
    >
      {style.label}
    </span>
  );
}
