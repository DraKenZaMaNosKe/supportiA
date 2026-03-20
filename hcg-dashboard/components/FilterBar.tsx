"use client";

interface Props {
  search: string;
  onSearchChange: (val: string) => void;
  estadoFilter: string;
  onEstadoChange: (val: string) => void;
}

export default function FilterBar({ search, onSearchChange, estadoFilter, onEstadoChange }: Props) {
  return (
    <div className="flex flex-wrap gap-4 mb-6">
      <input
        type="text"
        placeholder="Buscar por inventario, serie, modelo..."
        value={search}
        onChange={(e) => onSearchChange(e.target.value)}
        className="flex-1 min-w-[200px] px-4 py-2 rounded border text-sm"
        style={{
          background: "var(--night)",
          borderColor: "var(--border-dark)",
          color: "var(--text-light)",
        }}
      />
      <select
        value={estadoFilter}
        onChange={(e) => onEstadoChange(e.target.value)}
        className="px-4 py-2 rounded border text-sm"
        style={{
          background: "var(--night)",
          borderColor: "var(--border-dark)",
          color: "var(--text-gold)",
        }}
      >
        <option value="">Todos los estados</option>
        <option value="Activo">★ Activo</option>
        <option value="En proceso">✦ En proceso</option>
        <option value="Baja">✖ Baja</option>
      </select>
    </div>
  );
}
