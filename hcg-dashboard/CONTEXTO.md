# HCG Dashboard - Contexto del Proyecto

## Qué es
Plataforma web para reemplazar Google Sheets + Apps Script como sistema de registro y monitoreo de equipos IT del Hospital Civil de Guadalajara (HCG).

## Stack
- **Frontend**: Next.js 16, React 19, Tailwind CSS 4
- **Backend**: Next.js API Routes (app/api/equipos/route.ts)
- **Base de datos**: Supabase (PostgreSQL en la nube, gratis, sin instalar servidor)
- **Deploy**: Vercel (gratis)
- **Tema visual**: Cósmico / Saint Seiya - "Caballeros de Informática"

## Arquitectura
```
PowerShell/C# (equipos) ──POST JSON──> Next.js API (Vercel) ──SQL──> Supabase (PostgreSQL)
                                              │
                                     Frontend Web (tema cósmico)
```

## Compatibilidad con sistema actual
La API acepta **exactamente el mismo JSON** que los scripts PowerShell y C# ya envían. Solo hay que cambiar la URL:
```powershell
# De:
$GoogleSheetURL = "https://script.google.com/macros/s/AKfycby.../exec"
# A:
$GoogleSheetURL = "https://hcg-dashboard.vercel.app/api/equipos"
```

Acciones soportadas: `crear`, `actualizar`, `ip`, `software`, `sistema`, `diagnostico`, `verificar`

## Base de datos - 5 tablas
| Tabla | Reemplaza | Propósito |
|-------|-----------|-----------|
| `equipos` | Hoja "Registro" | Hardware, estado, IPs, MACs |
| `inventario_software` | Hoja "Inventario_Software" | Office, Chrome, ESET, etc. |
| `reporte_sistema` | Hoja "Reporte_Sistema" | Impresoras, usuarios, apps, disco |
| `diagnostico_salud` | Hoja "Diagnostico_Salud" | RAM, CPU, temperatura, procesos |
| `series_faa` | Hoja FAA externa | Verificación de números de serie |

Schema SQL: `supabase-schema.sql`

## Estructura del proyecto
```
hcg-dashboard/
├── app/
│   ├── api/equipos/route.ts     ← API principal (reemplaza doPost de Apps Script)
│   ├── api/stats/route.ts       ← Estadísticas para dashboard
│   ├── page.tsx                 ← Dashboard: contadores + actividad reciente
│   ├── equipos/page.tsx         ← Lista con filtros y búsqueda
│   ├── equipos/[invSt]/page.tsx ← Detalle: 4 tabs (hardware, software, sistema, salud)
│   ├── diagnostico/page.tsx     ← Panel de salud con gauges RAM/CPU/disco
│   ├── layout.tsx               ← Layout cósmico con navbar
│   └── globals.css              ← Paleta COSMOS (dorado, cyan, púrpura)
├── components/
│   ├── StatusBadge.tsx          ← Badge Activo/En proceso/Baja
│   ├── StatsPanel.tsx           ← 6 tarjetas de estadísticas
│   ├── EquipoTable.tsx          ← Tabla con colores alternados por estado
│   └── FilterBar.tsx            ← Búsqueda + filtro por estado
├── lib/
│   ├── supabase.ts              ← Clientes Supabase (admin + browser)
│   ├── types.ts                 ← Interfaces TypeScript
│   └── format.ts                ← formatMAC, formatRAM, formatDisco
└── supabase-schema.sql          ← SQL para crear tablas en Supabase
```

## Estado actual - Qué ya está hecho ✅
1. ✅ Proyecto Next.js creado y compila correctamente
2. ✅ API compatible con scripts PowerShell/C# (7 acciones)
3. ✅ Frontend: Dashboard, Lista equipos, Detalle equipo (4 tabs), Diagnóstico
4. ✅ Componentes: StatusBadge, StatsPanel, EquipoTable, FilterBar
5. ✅ Schema SQL para Supabase (5 tablas + triggers + RLS)
6. ✅ Tema cósmico Saint Seiya con paleta COSMOS
7. ✅ Rama `feature/hcg-dashboard` creada y subida al repo

## Pasos que faltan ⏳

### Paso 1: Crear proyecto en Supabase
1. Ir a https://supabase.com y crear cuenta (o login con GitHub)
2. Crear nuevo proyecto (nombre: `hcg-dashboard`, región: us-east-1)
3. Anotar las credenciales:
   - Project URL (`https://xxxxx.supabase.co`)
   - Anon Key (pública, para el frontend)
   - Service Role Key (privada, para la API)

### Paso 2: Crear las tablas
1. En Supabase, ir a **SQL Editor** > **New Query**
2. Pegar todo el contenido de `supabase-schema.sql`
3. Click en **Run** - crea las 5 tablas, índices, triggers y políticas RLS

### Paso 3: Configurar variables de entorno
1. Editar `hcg-dashboard/.env.local` con las credenciales de Supabase:
```
NEXT_PUBLIC_SUPABASE_URL=https://tu-proyecto.supabase.co
NEXT_PUBLIC_SUPABASE_ANON_KEY=eyJhbGci...
SUPABASE_SERVICE_ROLE_KEY=eyJhbGci...
```

### Paso 4: Probar en local
```bash
cd hcg-dashboard
npm run dev
```
- Abrir http://localhost:3000
- Verificar que el dashboard cargue (estará vacío, sin datos)
- Probar el API con curl:
```bash
curl -X POST http://localhost:3000/api/equipos \
  -H "Content-Type: application/json" \
  -d '{"Accion":"crear","InvST":"99999","Serie":"TEST123","Marca":"Lenovo"}'
```

### Paso 5: Migrar datos de Google Sheets
- Exportar cada hoja como CSV
- Importar a Supabase via SQL Editor o la UI de tabla
- Verificar conteos

### Paso 6: Desplegar en Vercel
1. Ir a https://vercel.com y conectar repo de GitHub
2. Seleccionar carpeta `hcg-dashboard` como root
3. Agregar variables de entorno (las 3 de Supabase)
4. Deploy automático
5. Obtener URL pública: `https://hcg-dashboard.vercel.app`

### Paso 7: Conectar equipos al nuevo backend
- Cambiar `$GoogleSheetURL` en todos los scripts PowerShell y en ConfigService.cs
- Apuntar a `https://hcg-dashboard.vercel.app/api/equipos`
- Probar con 1 equipo primero antes de cambiar todos

### Paso 8 (futuro): Mejoras opcionales
- Autenticación para el dashboard (Supabase Auth)
- Exportar a Excel/PDF desde el dashboard
- Notificaciones de equipos críticos
- Gráficas históricas de salud
- Real-time updates con Supabase Realtime

## Archivos clave del sistema actual (referencia)
- `ConfiguradorHCG/Google_Apps_Script_v4.js` - Backend actual (lógica replicada en route.ts)
- `ConfiguradorHCG/ConfigurarEquipoHCG.ps1` - Cliente principal (payloads JSON en líneas ~776, 859, 924)
- `iAUptowin/Services/ConfigService.cs` - Cliente C# (mismos payloads)
- `ConfiguradorHCG/Config.ps1` - Configuración centralizada

## Rama Git
- Branch: `feature/hcg-dashboard`
- Repo: https://github.com/DraKenZaMaNosKe/supportiA.git
