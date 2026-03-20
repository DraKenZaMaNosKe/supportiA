-- =============================================================================
-- HCG DASHBOARD - Schema para Supabase (PostgreSQL)
-- =============================================================================
-- Ejecutar en: Supabase > SQL Editor > New Query > Pegar y Run
-- =============================================================================

-- 1. EQUIPOS (reemplaza hoja "Registro")
CREATE TABLE equipos (
  id                UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  inv_st            TEXT NOT NULL UNIQUE,
  fecha_registro    TEXT NOT NULL DEFAULT '',
  marca             TEXT DEFAULT 'Lenovo',
  modelo            TEXT DEFAULT 'ThinkCentre M70s Gen 5',
  serie             TEXT NOT NULL,
  procesador        TEXT DEFAULT '',
  nucleos           INTEGER DEFAULT 14,
  ram               TEXT DEFAULT '',
  disco             TEXT DEFAULT '',
  graficos          TEXT DEFAULT 'Intel UHD 770',
  wifi              TEXT DEFAULT 'Wi-Fi 6',
  bluetooth         TEXT DEFAULT '5.1',
  sistema_operativo TEXT DEFAULT 'Win 11 Pro',
  mac_ethernet      TEXT DEFAULT '',
  mac_wifi          TEXT DEFAULT '',
  product_key       TEXT DEFAULT '',
  fecha_fabricacion TEXT DEFAULT '',
  garantia          TEXT DEFAULT '',
  ubicacion         TEXT DEFAULT '',
  departamento      TEXT DEFAULT '',
  usuario           TEXT DEFAULT '',
  estado            TEXT NOT NULL DEFAULT 'En proceso'
                    CHECK (estado IN ('Activo','En proceso','Baja')),
  faa               TEXT DEFAULT '',
  ip_ethernet       TEXT DEFAULT '',
  ip_wifi           TEXT DEFAULT '',
  red_wifi          TEXT DEFAULT '',
  ultima_conexion   TIMESTAMPTZ,
  created_at        TIMESTAMPTZ DEFAULT NOW(),
  updated_at        TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_equipos_inv_st ON equipos(inv_st);
CREATE INDEX idx_equipos_mac_ethernet ON equipos(mac_ethernet);
CREATE INDEX idx_equipos_estado ON equipos(estado);

-- 2. INVENTARIO SOFTWARE (reemplaza hoja "Inventario_Software")
CREATE TABLE inventario_software (
  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  inv_st           TEXT NOT NULL UNIQUE,
  nombre_equipo    TEXT DEFAULT '',
  windows_version  TEXT DEFAULT '',
  windows_build    TEXT DEFAULT '',
  windows_activado TEXT DEFAULT '',
  product_key      TEXT DEFAULT '',
  office           TEXT DEFAULT '',
  chrome           TEXT DEFAULT '',
  acrobat          TEXT DEFAULT '',
  dotnet35         TEXT DEFAULT '',
  dedalus          TEXT DEFAULT '',
  eset             TEXT DEFAULT '',
  winrar           TEXT DEFAULT '',
  otro_software    TEXT DEFAULT '',
  usuario_windows  TEXT DEFAULT '',
  fecha_config     TIMESTAMPTZ,
  notas            TEXT DEFAULT '',
  created_at       TIMESTAMPTZ DEFAULT NOW(),
  updated_at       TIMESTAMPTZ DEFAULT NOW()
);

-- 3. REPORTE SISTEMA (reemplaza hoja "Reporte_Sistema")
CREATE TABLE reporte_sistema (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  inv_st              TEXT NOT NULL DEFAULT '',
  nombre_equipo       TEXT DEFAULT '',
  mac_ethernet        TEXT NOT NULL UNIQUE,
  impresoras          TEXT DEFAULT '',
  usuarios            TEXT DEFAULT '',
  apps_instaladas     TEXT DEFAULT '',
  accesos_escritorio  TEXT DEFAULT '',
  espacio_libre_gb    TEXT DEFAULT '',
  mb_limpiados        TEXT DEFAULT '',
  fecha_reporte       TIMESTAMPTZ,
  created_at          TIMESTAMPTZ DEFAULT NOW(),
  updated_at          TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_reporte_mac ON reporte_sistema(mac_ethernet);

-- 4. DIAGNOSTICO SALUD (reemplaza hoja "Diagnostico_Salud")
CREATE TABLE diagnostico_salud (
  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  inv_st           TEXT NOT NULL DEFAULT '',
  nombre_equipo    TEXT DEFAULT '',
  mac_ethernet     TEXT NOT NULL UNIQUE,
  ram_total_gb     REAL DEFAULT 0,
  ram_usada_gb     REAL DEFAULT 0,
  ram_libre_gb     REAL DEFAULT 0,
  ram_pct          INTEGER DEFAULT 0,
  top5_procesos    TEXT DEFAULT '',
  chrome_mb        INTEGER DEFAULT 0,
  chrome_procs     INTEGER DEFAULT 0,
  dedalus_mb       INTEGER DEFAULT 0,
  dedalus_procs    INTEGER DEFAULT 0,
  total_procs      INTEGER DEFAULT 0,
  cpu_pct          INTEGER DEFAULT 0,
  pagefile_usado   INTEGER DEFAULT 0,
  pagefile_total   INTEGER DEFAULT 0,
  uptime_dias      REAL DEFAULT 0,
  disco_libre_gb   TEXT DEFAULT '',
  estado           TEXT DEFAULT 'OK'
                   CHECK (estado IN ('OK','Atencion','Critico')),
  recomendacion    TEXT DEFAULT '',
  fecha_reporte    TIMESTAMPTZ,
  created_at       TIMESTAMPTZ DEFAULT NOW(),
  updated_at       TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX idx_diagnostico_mac ON diagnostico_salud(mac_ethernet);

-- 5. SERIES FAA (reemplaza la hoja FAA externa)
CREATE TABLE series_faa (
  id        SERIAL PRIMARY KEY,
  numero_si TEXT NOT NULL,
  serie     TEXT NOT NULL UNIQUE
);

CREATE INDEX idx_faa_serie ON series_faa(serie);

-- 6. TRIGGER: auto-actualizar updated_at
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_equipos_updated BEFORE UPDATE ON equipos
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_software_updated BEFORE UPDATE ON inventario_software
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_reporte_updated BEFORE UPDATE ON reporte_sistema
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_diagnostico_updated BEFORE UPDATE ON diagnostico_salud
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- 7. ROW LEVEL SECURITY
ALTER TABLE equipos ENABLE ROW LEVEL SECURITY;
ALTER TABLE inventario_software ENABLE ROW LEVEL SECURITY;
ALTER TABLE reporte_sistema ENABLE ROW LEVEL SECURITY;
ALTER TABLE diagnostico_salud ENABLE ROW LEVEL SECURITY;
ALTER TABLE series_faa ENABLE ROW LEVEL SECURITY;

-- Permitir lectura publica (para el dashboard)
CREATE POLICY "Public read equipos" ON equipos FOR SELECT USING (true);
CREATE POLICY "Public read software" ON inventario_software FOR SELECT USING (true);
CREATE POLICY "Public read reportes" ON reporte_sistema FOR SELECT USING (true);
CREATE POLICY "Public read diagnostico" ON diagnostico_salud FOR SELECT USING (true);
CREATE POLICY "Public read faa" ON series_faa FOR SELECT USING (true);
