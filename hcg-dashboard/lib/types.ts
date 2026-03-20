export interface Equipo {
  id: string;
  inv_st: string;
  fecha_registro: string;
  marca: string;
  modelo: string;
  serie: string;
  procesador: string;
  nucleos: number;
  ram: string;
  disco: string;
  graficos: string;
  wifi: string;
  bluetooth: string;
  sistema_operativo: string;
  mac_ethernet: string;
  mac_wifi: string;
  product_key: string;
  fecha_fabricacion: string;
  garantia: string;
  ubicacion: string;
  departamento: string;
  usuario: string;
  estado: "Activo" | "En proceso" | "Baja";
  faa: string;
  ip_ethernet: string;
  ip_wifi: string;
  red_wifi: string;
  ultima_conexion: string;
  created_at: string;
  updated_at: string;
}

export interface InventarioSoftware {
  id: string;
  inv_st: string;
  nombre_equipo: string;
  windows_version: string;
  windows_build: string;
  windows_activado: string;
  product_key: string;
  office: string;
  chrome: string;
  acrobat: string;
  dotnet35: string;
  dedalus: string;
  eset: string;
  winrar: string;
  otro_software: string;
  usuario_windows: string;
  fecha_config: string;
  notas: string;
}

export interface ReporteSistema {
  id: string;
  inv_st: string;
  nombre_equipo: string;
  mac_ethernet: string;
  impresoras: string;
  usuarios: string;
  apps_instaladas: string;
  accesos_escritorio: string;
  espacio_libre_gb: string;
  mb_limpiados: string;
  fecha_reporte: string;
}

export interface DiagnosticoSalud {
  id: string;
  inv_st: string;
  nombre_equipo: string;
  mac_ethernet: string;
  ram_total_gb: number;
  ram_usada_gb: number;
  ram_libre_gb: number;
  ram_pct: number;
  top5_procesos: string;
  chrome_mb: number;
  chrome_procs: number;
  dedalus_mb: number;
  dedalus_procs: number;
  total_procs: number;
  cpu_pct: number;
  pagefile_usado: number;
  pagefile_total: number;
  uptime_dias: number;
  disco_libre_gb: string;
  estado: "OK" | "Atencion" | "Critico";
  recomendacion: string;
  fecha_reporte: string;
}

export interface DashboardStats {
  total: number;
  activos: number;
  en_proceso: number;
  baja: number;
  online_24h: number;
  criticos: number;
}
