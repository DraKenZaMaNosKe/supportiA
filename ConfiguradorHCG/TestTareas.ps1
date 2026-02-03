# =============================================================================
# TEST - Crear tareas programadas HCG
# =============================================================================
# Ejecutar como administrador para verificar que se crean correctamente
# =============================================================================

Write-Host ""
Write-Host "  === TEST: Creando tareas programadas HCG ===" -ForegroundColor Cyan
Write-Host ""

# Crear carpeta
if (-not (Test-Path "C:\HCG_Logs")) {
    New-Item -ItemType Directory -Path "C:\HCG_Logs" -Force | Out-Null
    Write-Host "  Carpeta C:\HCG_Logs creada" -ForegroundColor Gray
}

# Crear un script de prueba minimo para que las tareas tengan algo que ejecutar
"Write-Output 'test'" | Out-File -FilePath "C:\HCG_Logs\report_ip.ps1" -Encoding UTF8 -Force
"Write-Output 'test'" | Out-File -FilePath "C:\HCG_Logs\report_system.ps1" -Encoding UTF8 -Force
Write-Host "  Scripts de prueba creados" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# TAREA 1: HCG_ReporteIP
# =============================================================================
Write-Host "  [1/2] Creando tarea HCG_ReporteIP..." -ForegroundColor Yellow

# Eliminar si existe
schtasks /delete /tn "HCG_ReporteIP" /f 2>$null | Out-Null

try {
    $Action1 = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"C:\HCG_Logs\report_ip.ps1`""
    Write-Host "    Action creada OK" -ForegroundColor Gray

    $TriggerLogon1 = New-ScheduledTaskTrigger -AtLogOn
    $TriggerLogon1.Delay = "PT60S"
    Write-Host "    Trigger AtLogOn creada OK" -ForegroundColor Gray

    $TriggerRepeat1 = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 3)
    Write-Host "    Trigger Repeticion creada OK" -ForegroundColor Gray

    $Settings1 = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
        -MultipleInstances IgnoreNew
    Write-Host "    Settings creados OK" -ForegroundColor Gray

    Register-ScheduledTask -TaskName "HCG_ReporteIP" -Action $Action1 `
        -Trigger @($TriggerLogon1, $TriggerRepeat1) `
        -Settings $Settings1 `
        -Description "HCG - Reporte automatico de IP cada 3 horas" `
        -RunLevel Limited -Force | Out-Null

    $Tarea1 = Get-ScheduledTask -TaskName "HCG_ReporteIP" -ErrorAction SilentlyContinue
    if ($Tarea1) {
        Write-Host "    [OK] HCG_ReporteIP creada exitosamente - Estado: $($Tarea1.State)" -ForegroundColor Green
    } else {
        Write-Host "    [FALLO] La tarea no se creo" -ForegroundColor Red
    }
} catch {
    Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# =============================================================================
# TAREA 2: HCG_ReporteSistema
# =============================================================================
Write-Host "  [2/2] Creando tarea HCG_ReporteSistema..." -ForegroundColor Yellow

schtasks /delete /tn "HCG_ReporteSistema" /f 2>$null | Out-Null

try {
    $Action2 = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"C:\HCG_Logs\report_system.ps1`""
    Write-Host "    Action creada OK" -ForegroundColor Gray

    $TriggerLogon2 = New-ScheduledTaskTrigger -AtLogOn
    $TriggerLogon2.Delay = "PT120S"
    Write-Host "    Trigger AtLogOn creada OK" -ForegroundColor Gray

    $Settings2 = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -MultipleInstances IgnoreNew
    Write-Host "    Settings creados OK" -ForegroundColor Gray

    Register-ScheduledTask -TaskName "HCG_ReporteSistema" -Action $Action2 `
        -Trigger $TriggerLogon2 `
        -Settings $Settings2 `
        -Description "HCG - Reporte de sistema y limpieza al iniciar sesion" `
        -RunLevel Limited -Force | Out-Null

    $Tarea2 = Get-ScheduledTask -TaskName "HCG_ReporteSistema" -ErrorAction SilentlyContinue
    if ($Tarea2) {
        Write-Host "    [OK] HCG_ReporteSistema creada exitosamente - Estado: $($Tarea2.State)" -ForegroundColor Green
    } else {
        Write-Host "    [FALLO] La tarea no se creo" -ForegroundColor Red
    }
} catch {
    Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
}

# =============================================================================
# VERIFICACION FINAL
# =============================================================================
Write-Host ""
Write-Host "  === VERIFICACION ===" -ForegroundColor Cyan
Write-Host ""

$Tareas = Get-ScheduledTask -TaskName "HCG_*" -ErrorAction SilentlyContinue
if ($Tareas) {
    foreach ($t in $Tareas) {
        Write-Host "  Tarea: $($t.TaskName)" -ForegroundColor White
        Write-Host "    Estado:      $($t.State)" -ForegroundColor Gray
        Write-Host "    Descripcion: $($t.Description)" -ForegroundColor Gray
        $Info = Get-ScheduledTaskInfo -TaskName $t.TaskName -ErrorAction SilentlyContinue
        if ($Info) {
            Write-Host "    Ultimo resultado: $($Info.LastTaskResult)" -ForegroundColor Gray
        }
        Write-Host ""
    }
} else {
    Write-Host "  [ERROR] No se encontraron tareas HCG_*" -ForegroundColor Red
}

Write-Host "  === FIN DEL TEST ===" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Presiona Enter para salir"
