export function formatMAC(mac: string | null | undefined): string {
  if (!mac) return "";
  const clean = mac.replace(/[^a-fA-F0-9]/g, "").toUpperCase();
  if (clean.length !== 12) return mac;
  return clean.match(/.{2}/g)!.join(":");
}

export function formatRAM(ram: string | number | null | undefined): string {
  if (!ram) return "";
  const val = String(ram).replace(/[^0-9]/g, "");
  const gb = parseInt(val);
  if (isNaN(gb)) return String(ram);
  const tipo = gb >= 8 ? "DDR5" : "DDR4";
  return `${gb} GB ${tipo}`;
}

export function formatDisco(
  disco: string | number | null | undefined,
  tipo?: string
): string {
  if (!disco) return "";
  const val = String(disco).replace(/[^0-9]/g, "");
  const gb = parseInt(val);
  if (isNaN(gb)) return String(disco);
  const discoTipo = tipo || (gb >= 256 ? "SSD" : "HDD");
  return `${gb} GB ${discoTipo}`;
}

export function normalizeMAC(mac: string): string {
  return mac.toUpperCase().replace(/[^A-F0-9]/g, "");
}
