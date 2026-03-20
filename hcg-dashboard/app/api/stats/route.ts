import { NextResponse } from "next/server";
import { supabaseAdmin } from "@/lib/supabase";

export async function GET() {
  const [equiposRes, diagRes, onlineRes] = await Promise.all([
    supabaseAdmin.from("equipos").select("estado"),
    supabaseAdmin.from("diagnostico_salud").select("estado").eq("estado", "Critico"),
    supabaseAdmin
      .from("equipos")
      .select("id")
      .gte("ultima_conexion", new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString()),
  ]);

  const equipos = equiposRes.data || [];
  const stats = {
    total: equipos.length,
    activos: equipos.filter((e) => e.estado === "Activo").length,
    en_proceso: equipos.filter((e) => e.estado === "En proceso").length,
    baja: equipos.filter((e) => e.estado === "Baja").length,
    online_24h: onlineRes.data?.length || 0,
    criticos: diagRes.data?.length || 0,
  };

  return NextResponse.json(stats);
}
