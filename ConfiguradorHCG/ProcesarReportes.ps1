# =============================================================================
# PROCESAR REPORTES Y ACTUALIZAR EXCEL - HOSPITAL CIVIL FAA
# =============================================================================
# Este script lee los reportes JSON generados por ConfigurarEquipo.ps1
# y actualiza el archivo Excel en Google Drive
# =============================================================================

# Cargar configuracion
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$ScriptPath\Config.ps1"

# Importar modulo de Excel (ImportExcel)
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Instalando modulo ImportExcel..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel

# =============================================================================
# CONFIGURACION
# =============================================================================

$RutaReportes = "C:\HCG_Reportes"
$RutaExcel = "C:\Users\lalog\Mi unidad (ejcontreras@hcg.gob.mx)\registroInstaller\RegistroEquipos_HCG.xlsx"
$RutaProcesados = "$RutaReportes\Procesados"

# =============================================================================
# FUNCIONES
# =============================================================================

function Get-ReportesPendientes {
    if (-not (Test-Path $RutaReportes)) {
        Write-Host "No existe la carpeta de reportes: $RutaReportes" -ForegroundColor Red
        return @()
    }

    $Reportes = Get-ChildItem -Path $RutaReportes -Filter "Equipo_*.json"
    return $Reportes
}

function Read-ReporteJSON {
    param([string]$Ruta)

    $Contenido = Get-Content -Path $Ruta -Raw | ConvertFrom-Json
    return $Contenido
}

function Add-EquipoAExcel {
    param(
        [object]$Reporte,
        [string]$RutaExcel
    )

    # Leer Excel existente
    if (Test-Path $RutaExcel) {
        $DatosExistentes = Import-Excel -Path $RutaExcel
        $UltimoNumero = ($DatosExistentes | Measure-Object -Property "No." -Maximum).Maximum
        if (-not $UltimoNumero) { $UltimoNumero = 0 }
    } else {
        $DatosExistentes = @()
        $UltimoNumero = 0
    }

    # Verificar si ya existe este equipo
    $YaExiste = $DatosExistentes | Where-Object { $_."No. Serie" -eq $Reporte.NumeroSerie }
    if ($YaExiste) {
        Write-Host "  [!] Equipo $($Reporte.NumeroSerie) ya existe en el Excel" -ForegroundColor Yellow
        return $false
    }

    # Calcular fecha de garantia
    $FechaFab = Get-Date
    $FechaGarantia = $FechaFab.AddYears(3).ToString("yyyy-MM-dd")

    # Crear nueva fila
    $NuevaFila = [PSCustomObject]@{
        "No." = $UltimoNumero + 1
        "Fecha" = (Get-Date -Format "yyyy-MM-dd")
        "Inv_ST" = $Reporte.NumeroInventarioST
        "Marca" = $Reporte.Especificaciones.Marca
        "Modelo" = $Reporte.Especificaciones.Modelo
        "No. Serie" = $Reporte.NumeroSerie
        "Procesador" = $Reporte.Especificaciones.Procesador
        "Nucleos" = $Reporte.Especificaciones.Nucleos
        "RAM" = $Reporte.Especificaciones.RAM
        "Disco Duro" = $Reporte.Especificaciones.Disco
        "Graficos" = $Reporte.Especificaciones.Graficos
        "WiFi" = $Reporte.Especificaciones.WiFi
        "Bluetooth" = $Reporte.Especificaciones.Bluetooth
        "Sistema Op." = $Reporte.Especificaciones.SistemaOp
        "MAC Ethernet" = $Reporte.MACEthernet
        "MAC WiFi" = $Reporte.MACWiFi
        "Product Key ID" = $Reporte.ProductKey
        "Fecha Fab." = (Get-Date -Format "yyyy-MM-dd")
        "Garantia Hasta" = $FechaGarantia
        "En Lista FAA" = "Pendiente verificar"
    }

    # Agregar al Excel
    $NuevaFila | Export-Excel -Path $RutaExcel -Append -AutoSize -TableStyle Medium2

    Write-Host "  [OK] Equipo $($Reporte.NumeroInventarioST) agregado al Excel" -ForegroundColor Green
    return $true
}

function Move-ReporteProcesado {
    param([string]$RutaReporte)

    if (-not (Test-Path $RutaProcesados)) {
        New-Item -ItemType Directory -Path $RutaProcesados -Force | Out-Null
    }

    $NombreArchivo = Split-Path -Leaf $RutaReporte
    Move-Item -Path $RutaReporte -Destination "$RutaProcesados\$NombreArchivo" -Force
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║       PROCESAR REPORTES - HOSPITAL CIVIL FAA                 ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Obtener reportes pendientes
$Reportes = Get-ReportesPendientes

if ($Reportes.Count -eq 0) {
    Write-Host "  No hay reportes pendientes por procesar" -ForegroundColor Yellow
    Write-Host "  Carpeta: $RutaReportes" -ForegroundColor Gray
    Write-Host ""
    exit
}

Write-Host "  Reportes encontrados: $($Reportes.Count)" -ForegroundColor Green
Write-Host ""

$Procesados = 0
$Errores = 0

foreach ($ArchivoReporte in $Reportes) {
    Write-Host "  Procesando: $($ArchivoReporte.Name)" -ForegroundColor Cyan

    try {
        $Reporte = Read-ReporteJSON -Ruta $ArchivoReporte.FullName

        $Agregado = Add-EquipoAExcel -Reporte $Reporte -RutaExcel $RutaExcel

        if ($Agregado) {
            Move-ReporteProcesado -RutaReporte $ArchivoReporte.FullName
            $Procesados++
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $Errores++
    }
}

Write-Host ""
Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Resumen:" -ForegroundColor Cyan
Write-Host "    Procesados correctamente: $Procesados" -ForegroundColor Green
Write-Host "    Errores: $Errores" -ForegroundColor $(if ($Errores -gt 0) { "Red" } else { "Gray" })
Write-Host ""
Write-Host "  El archivo Excel se sincronizara automaticamente con Google Drive" -ForegroundColor Yellow
Write-Host "  Ruta: $RutaExcel" -ForegroundColor Gray
Write-Host ""

Read-Host "  Presiona Enter para salir"
