# =============================================================================
# ENVIAR DATOS PENDIENTES A GOOGLE SHEETS
# =============================================================================
# Ejecuta este script desde tu PC para enviar los JSON que no se pudieron
# enviar desde los equipos nuevos
# =============================================================================

$GoogleSheetURL = "https://script.google.com/macros/s/AKfycbwDAZERvkP6V8sZczvpJgh6LvoXOqwuWymCauicX5dyYGRh5Iym4J5czjHg3tDHbPtP/exec"
$RutaPendientes = "C:\HCG_Logs"
$RutaProcesados = "C:\HCG_Logs\Procesados"

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "  ENVIAR DATOS PENDIENTES A GOOGLE SHEETS" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

# Forzar TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Buscar archivos JSON pendientes
$Archivos = Get-ChildItem -Path $RutaPendientes -Filter "Equipo_*.json" -ErrorAction SilentlyContinue

if ($Archivos.Count -eq 0) {
    Write-Host "  No hay archivos pendientes en: $RutaPendientes" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "  Presiona Enter para salir"
    exit
}

Write-Host "  Archivos pendientes encontrados: $($Archivos.Count)" -ForegroundColor Green
Write-Host ""

# Crear carpeta de procesados
if (-not (Test-Path $RutaProcesados)) {
    New-Item -ItemType Directory -Path $RutaProcesados -Force | Out-Null
}

$Enviados = 0
$Errores = 0

foreach ($Archivo in $Archivos) {
    Write-Host "  Procesando: $($Archivo.Name)" -ForegroundColor Cyan

    try {
        $Body = Get-Content -Path $Archivo.FullName -Raw

        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json" -TimeoutSec 60

        if ($Response.status -eq "OK") {
            Write-Host "    [OK] Enviado (fila $($Response.row))" -ForegroundColor Green
            # Mover a procesados
            Move-Item -Path $Archivo.FullName -Destination $RutaProcesados -Force
            $Enviados++
        } else {
            Write-Host "    [WARN] Respuesta inesperada" -ForegroundColor Yellow
            $Errores++
        }
    } catch {
        Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $Errores++
    }

    Start-Sleep -Milliseconds 500
}

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "  RESUMEN" -ForegroundColor Green
Write-Host "  ============================================" -ForegroundColor Green
Write-Host "  Enviados correctamente: $Enviados" -ForegroundColor White
Write-Host "  Errores: $Errores" -ForegroundColor $(if ($Errores -gt 0) { "Red" } else { "White" })
Write-Host ""

Read-Host "  Presiona Enter para salir"
