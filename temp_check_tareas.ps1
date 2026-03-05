# Verificar estado de tareas programadas HCG
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  ESTADO DE TAREAS PROGRAMADAS HCG" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

$tareasHCG = @("HCG_ReporteIP", "HCG_ReporteSistema", "HCG_ReporteDiagnostico", "HCG_ActualizacionWindows")

foreach ($nombre in $tareasHCG) {
    $tarea = Get-ScheduledTask -TaskName $nombre -ErrorAction SilentlyContinue
    if ($tarea) {
        $info = Get-ScheduledTaskInfo -TaskName $nombre -ErrorAction SilentlyContinue
        $ultimaEjec = if ($info.LastRunTime -and $info.LastRunTime.Year -gt 2000) { $info.LastRunTime.ToString("dd/MM/yyyy HH:mm") } else { "Nunca" }
        $proximaEjec = if ($info.NextRunTime -and $info.NextRunTime.Year -gt 2000) { $info.NextRunTime.ToString("dd/MM/yyyy HH:mm") } else { "N/A" }
        $resultado = $info.LastTaskResult

        $color = switch ($tarea.State) {
            "Ready" { "Green" }
            "Running" { "Yellow" }
            "Disabled" { "Red" }
            default { "Gray" }
        }

        Write-Host "  $nombre" -ForegroundColor White
        Write-Host "    Estado:           $($tarea.State)" -ForegroundColor $color
        Write-Host "    Descripcion:      $($tarea.Description)" -ForegroundColor Gray
        Write-Host "    Ultima ejecucion: $ultimaEjec" -ForegroundColor Gray
        Write-Host "    Ultimo resultado: $resultado $(if ($resultado -eq 0) {'(OK)'} elseif ($resultado -eq 267011) {'(No ha corrido aun)'} else {'(Verificar)'})" -ForegroundColor Gray
        Write-Host "    Proxima ejecucion: $proximaEjec" -ForegroundColor Gray
        Write-Host ""
    } else {
        Write-Host "  $nombre" -ForegroundColor White
        Write-Host "    [NO EXISTE] Esta tarea no esta creada en este equipo" -ForegroundColor Red
        Write-Host ""
    }
}

# Verificar scripts en C:\HCG_Logs
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  SCRIPTS EN C:\HCG_Logs" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path "C:\HCG_Logs") {
    $archivos = Get-ChildItem "C:\HCG_Logs" -File
    if ($archivos) {
        foreach ($a in $archivos) {
            $tamano = if ($a.Length -gt 1KB) { "$([math]::Round($a.Length/1KB, 1)) KB" } else { "$($a.Length) bytes" }
            Write-Host "    $($a.Name)  ($tamano)  - Modificado: $($a.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))" -ForegroundColor Gray
        }
    } else {
        Write-Host "    [Carpeta vacia]" -ForegroundColor Yellow
    }
} else {
    Write-Host "    [NO EXISTE] La carpeta C:\HCG_Logs no existe" -ForegroundColor Red
}
Write-Host ""
