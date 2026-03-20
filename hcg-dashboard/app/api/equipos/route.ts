import { NextResponse } from "next/server";
import { supabaseAdmin } from "@/lib/supabase";
import { formatMAC, formatRAM, formatDisco, normalizeMAC } from "@/lib/format";

function jsonResponse(data: Record<string, unknown>) {
  return NextResponse.json(data);
}

// GET /api/equipos - Info de la API
export async function GET() {
  return jsonResponse({
    status: "OK",
    mensaje: "API Registro Cosmico de Equipos HCG v5.0",
    fecha: new Date().toLocaleString("es-MX", { timeZone: "America/Mexico_City" }),
  });
}

// POST /api/equipos - Endpoint principal (reemplaza doPost de Apps Script)
export async function POST(request: Request) {
  try {
    const data = await request.json();
    const accion = (data.Accion || "crear").toLowerCase();

    switch (accion) {
      case "crear":
        return handleCrear(data);
      case "actualizar":
        return handleActualizar(data);
      case "ip":
        return handleIP(data);
      case "software":
        return handleSoftware(data);
      case "sistema":
        return handleSistema(data);
      case "diagnostico":
        return handleDiagnostico(data);
      case "verificar":
        return handleVerificar(data);
      default:
        return jsonResponse({ status: "ERROR", mensaje: "Accion no reconocida: " + accion });
    }
  } catch (error) {
    const msg = error instanceof Error ? error.message : "Error desconocido";
    return jsonResponse({ status: "ERROR", mensaje: msg });
  }
}

// ─── CREAR REGISTRO ─────────────────────────────────────────────────────────
async function handleCrear(data: Record<string, unknown>) {
  if (!data.InvST || !data.Serie) {
    return jsonResponse({ status: "ERROR", mensaje: "InvST y Serie son obligatorios" });
  }

  const invSt = String(data.InvST);
  const serie = String(data.Serie);

  // Verificar si ya existe
  const { data: existente } = await supabaseAdmin
    .from("equipos")
    .select("id, inv_st")
    .eq("inv_st", invSt)
    .single();

  if (existente) {
    await supabaseAdmin
      .from("equipos")
      .update({ estado: "En proceso", updated_at: new Date().toISOString() })
      .eq("inv_st", invSt);

    return jsonResponse({
      status: "OK",
      accion: "actualizado",
      row: existente.id,
      mensaje: "Equipo ya existia, actualizado a En proceso",
    });
  }

  // Buscar en FAA
  let faaResult = String(data.FAA || "");
  if (!faaResult) {
    faaResult = await buscarFAA(serie);
  }

  // Si FAA = NO ENCONTRADO, auto-asignar OPD
  let departamento = String(data.Departamento || "");
  if (faaResult === "NO ENCONTRADO" && !departamento) {
    departamento = "OPD";
  }

  const hoy = new Date().toLocaleDateString("es-MX", {
    timeZone: "America/Mexico_City",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });

  const { data: nuevo, error } = await supabaseAdmin
    .from("equipos")
    .insert({
      inv_st: invSt,
      fecha_registro: data.Fecha || hoy,
      marca: data.Marca || "Lenovo",
      modelo: data.Modelo || "ThinkCentre M70s Gen 5",
      serie,
      procesador: data.Procesador || "Intel Core i5-14500 vPro",
      nucleos: data.Nucleos || 14,
      ram: formatRAM(data.RAM as string),
      disco: formatDisco(data.Disco as string, data.DiscoTipo as string),
      graficos: data.Graficos || "Intel UHD 770",
      wifi: data.WiFi || "Wi-Fi 6",
      bluetooth: data.BT || "5.1",
      sistema_operativo: data.SO || "Win 11 Pro",
      mac_ethernet: formatMAC(data.MACEthernet as string),
      mac_wifi: formatMAC(data.MACWiFi as string),
      product_key: data.ProductKey || "",
      fecha_fabricacion: data.FechaFab || data.Fecha || hoy,
      garantia: data.Garantia || "",
      ubicacion: data.Ubicacion || "",
      departamento,
      usuario: data.Usuario || "",
      estado: "En proceso",
      faa: faaResult,
    })
    .select()
    .single();

  if (error) {
    return jsonResponse({ status: "ERROR", mensaje: error.message });
  }

  return jsonResponse({
    status: "OK",
    accion: "crear",
    row: nuevo?.id,
    faa: faaResult,
    serie,
    inventario: invSt,
  });
}

// ─── ACTUALIZAR REGISTRO ────────────────────────────────────────────────────
async function handleActualizar(data: Record<string, unknown>) {
  const invSt = String(data.InvST || "");
  if (!invSt) {
    return jsonResponse({ status: "ERROR", mensaje: "InvST es obligatorio" });
  }

  const updates: Record<string, unknown> = {
    estado: "Activo",
    updated_at: new Date().toISOString(),
  };
  if (data.Ubicacion) updates.ubicacion = data.Ubicacion;
  if (data.Departamento) updates.departamento = data.Departamento;
  if (data.Usuario) updates.usuario = data.Usuario;

  const { data: row, error } = await supabaseAdmin
    .from("equipos")
    .update(updates)
    .eq("inv_st", invSt)
    .select("id")
    .single();

  if (error || !row) {
    return jsonResponse({ status: "ERROR", mensaje: "Equipo no encontrado: " + invSt });
  }

  return jsonResponse({
    status: "OK",
    accion: "actualizar",
    row: row.id,
    estado: "Activo",
  });
}

// ─── ACTUALIZAR IP ──────────────────────────────────────────────────────────
async function handleIP(data: Record<string, unknown>) {
  const macRaw = String(data.MACEthernet || "");
  const macNorm = normalizeMAC(macRaw);

  if (!macNorm || macNorm.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC Ethernet invalida" });
  }

  // Buscar equipo por MAC normalizada
  const { data: equipos } = await supabaseAdmin
    .from("equipos")
    .select("id, mac_ethernet")
    .not("mac_ethernet", "is", null);

  const equipo = equipos?.find(
    (e) => normalizeMAC(e.mac_ethernet) === macNorm
  );

  if (!equipo) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC no encontrada: " + macRaw });
  }

  const { error } = await supabaseAdmin
    .from("equipos")
    .update({
      ip_ethernet: data.IPEthernet || "",
      ip_wifi: data.IPWiFi || "",
      red_wifi: data.SSIDWiFi || "",
      ultima_conexion: data.FechaReporte || new Date().toISOString(),
      updated_at: new Date().toISOString(),
    })
    .eq("id", equipo.id);

  if (error) {
    return jsonResponse({ status: "ERROR", mensaje: error.message });
  }

  return jsonResponse({
    status: "OK",
    accion: "ip_actualizada",
    row: equipo.id,
    ip: data.IPEthernet || data.IPWiFi,
    mensaje: "IP actualizada",
  });
}

// ─── REGISTRAR SOFTWARE ─────────────────────────────────────────────────────
async function handleSoftware(data: Record<string, unknown>) {
  const invSt = String(data.InvST || "");
  if (!invSt) {
    return jsonResponse({ status: "ERROR", mensaje: "InvST es obligatorio" });
  }

  const row = {
    inv_st: invSt,
    nombre_equipo: data.NombreEquipo || "",
    windows_version: data.WindowsVersion || "",
    windows_build: data.WindowsBuild || "",
    windows_activado: data.WindowsActivado || "",
    product_key: data.ProductKey || "",
    office: data.Office || "",
    chrome: data.Chrome || "",
    acrobat: data.Acrobat || "",
    dotnet35: data.DotNet35 || "",
    dedalus: data.Dedalus || "",
    eset: data.ESET || "",
    winrar: data.WinRAR || "",
    otro_software: data.OtroSoftware || "",
    usuario_windows: data.UsuarioWindows || "",
    fecha_config: data.FechaConfig || new Date().toISOString(),
    notas: data.Notas || "",
  };

  const { data: result, error } = await supabaseAdmin
    .from("inventario_software")
    .upsert(row, { onConflict: "inv_st" })
    .select()
    .single();

  if (error) {
    return jsonResponse({ status: "ERROR", mensaje: error.message });
  }

  return jsonResponse({
    status: "OK",
    accion: "software",
    row: result?.id,
    inventario: invSt,
  });
}

// ─── REPORTE SISTEMA ────────────────────────────────────────────────────────
async function handleSistema(data: Record<string, unknown>) {
  const macRaw = String(data.MACEthernet || "");
  const macNorm = normalizeMAC(macRaw);

  if (!macNorm || macNorm.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC invalida" });
  }

  // Resolver inv_st desde equipos
  const { data: equipos } = await supabaseAdmin
    .from("equipos")
    .select("inv_st, mac_ethernet")
    .not("mac_ethernet", "is", null);

  const equipo = equipos?.find(
    (e) => normalizeMAC(e.mac_ethernet) === macNorm
  );
  const invSt = equipo?.inv_st || "";

  const row = {
    inv_st: invSt,
    nombre_equipo: data.NombreEquipo || "",
    mac_ethernet: formatMAC(macRaw),
    impresoras: data.Impresoras || "",
    usuarios: data.Usuarios || "",
    apps_instaladas: data.AppsInstaladas || "",
    accesos_escritorio: data.AccesosEscritorio || "",
    espacio_libre_gb: data.EspacioLibreGB || "",
    mb_limpiados: data.MBLimpiados || "",
    fecha_reporte: data.FechaReporte || new Date().toISOString(),
  };

  const { data: result, error } = await supabaseAdmin
    .from("reporte_sistema")
    .upsert(row, { onConflict: "mac_ethernet" })
    .select()
    .single();

  if (error) {
    return jsonResponse({ status: "ERROR", mensaje: error.message });
  }

  return jsonResponse({
    status: "OK",
    accion: "sistema",
    row: result?.id,
  });
}

// ─── DIAGNOSTICO SALUD ──────────────────────────────────────────────────────
async function handleDiagnostico(data: Record<string, unknown>) {
  const macRaw = String(data.MACEthernet || "");
  const macNorm = normalizeMAC(macRaw);

  if (!macNorm || macNorm.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC invalida" });
  }

  const { data: equipos } = await supabaseAdmin
    .from("equipos")
    .select("inv_st, mac_ethernet")
    .not("mac_ethernet", "is", null);

  const equipo = equipos?.find(
    (e) => normalizeMAC(e.mac_ethernet) === macNorm
  );
  const invSt = equipo?.inv_st || "";

  const row = {
    inv_st: invSt,
    nombre_equipo: data.NombreEquipo || "",
    mac_ethernet: formatMAC(macRaw),
    ram_total_gb: Number(data.RAMTotalGB) || 0,
    ram_usada_gb: Number(data.RAMUsadaGB) || 0,
    ram_libre_gb: Number(data.RAMLibreGB) || 0,
    ram_pct: Number(data.RAMPct) || 0,
    top5_procesos: data.Top5Procesos || "",
    chrome_mb: Number(data.ChromeMB) || 0,
    chrome_procs: Number(data.ChromeProcs) || 0,
    dedalus_mb: Number(data.DedalusMB) || 0,
    dedalus_procs: Number(data.DedalusProcs) || 0,
    total_procs: Number(data.TotalProcs) || 0,
    cpu_pct: Number(data.CPUPct) || 0,
    pagefile_usado: Number(data.PageFileUsado) || 0,
    pagefile_total: Number(data.PageFileTotal) || 0,
    uptime_dias: Number(data.UptimeDias) || 0,
    disco_libre_gb: data.DiscoLibreGB || "",
    estado: data.Estado || "OK",
    recomendacion: data.Recomendacion || "",
    fecha_reporte: data.FechaReporte || new Date().toISOString(),
  };

  const { data: result, error } = await supabaseAdmin
    .from("diagnostico_salud")
    .upsert(row, { onConflict: "mac_ethernet" })
    .select()
    .single();

  if (error) {
    return jsonResponse({ status: "ERROR", mensaje: error.message });
  }

  return jsonResponse({
    status: "OK",
    accion: "diagnostico",
    row: result?.id,
  });
}

// ─── VERIFICAR FAA ──────────────────────────────────────────────────────────
async function handleVerificar(data: Record<string, unknown>) {
  const serie = String(data.Serie || "");
  const faaResult = await buscarFAA(serie);
  return jsonResponse({ status: "OK", faa: faaResult, serie });
}

// ─── BUSCAR EN SERIES FAA ───────────────────────────────────────────────────
async function buscarFAA(serie: string): Promise<string> {
  if (!serie) return "NO ENCONTRADO";
  try {
    const { data } = await supabaseAdmin
      .from("series_faa")
      .select("numero_si")
      .ilike("serie", serie.trim())
      .single();

    return data ? "SI #" + data.numero_si : "NO ENCONTRADO";
  } catch {
    return "NO ENCONTRADO";
  }
}
