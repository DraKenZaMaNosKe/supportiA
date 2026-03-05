# =============================================================================
# ACTUALIZADOR DE WINDOWS - HOSPITAL CIVIL FAA
# Script standalone para dejar equipos 100% actualizados
# =============================================================================
# Se eleva automaticamente como Administrador si no lo es
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- Deshabilitar QuickEdit Mode (evita pausas accidentales al hacer clic) ---
try {
    Add-Type -Name ConsoleUtils -Namespace Win32 -MemberDefinition @"
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetStdHandle(int nStdHandle);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
"@ -ErrorAction SilentlyContinue
    $handle = [Win32.ConsoleUtils]::GetStdHandle(-10)
    $mode = 0
    [Win32.ConsoleUtils]::GetConsoleMode($handle, [ref]$mode) | Out-Null
    $mode = $mode -band (-bnot 0x0040)
    $mode = $mode -bor 0x0080
    [Win32.ConsoleUtils]::SetConsoleMode($handle, $mode) | Out-Null
} catch {}

# --- Variables globales ---
$RutaLogs = "C:\HCG_Logs"
$LogFile = "$RutaLogs\winupdate.log"
$MaxRondas = 5
$Script:FailedKBs = @{}  # Rastreo de KBs que fallan repetidamente

# =============================================================================
# FUNCIONES DE UI (self-contained)
# =============================================================================

function Write-Log {
    param([string]$Mensaje, [string]$Tipo = "INFO")
    $Icon = switch ($Tipo) {
        "OK"    { [char]0x2605 }  # Estrella dorada
        "ERROR" { [char]0x2716 }  # Cruz roja
        "WARN"  { [char]0x26A0 }  # Triangulo alerta
        default { [char]0x2192 }  # Flecha cosmica
    }
    $Color = switch ($Tipo) {
        "OK"    { "Green" }
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        default { "Cyan" }
    }
    Write-Host "  $Icon [$Tipo] $Mensaje" -ForegroundColor $Color
    if (-not (Test-Path $RutaLogs)) { New-Item -ItemType Directory -Path $RutaLogs -Force | Out-Null }
    Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Tipo] $Mensaje"
}

function Write-Separator {
    $Star = [char]0x2606  # Estrella vacia
    $Dot  = [char]0x00B7  # Punto medio
    $Sep = "  "
    for ($i = 0; $i -lt 15; $i++) {
        if ($i % 3 -eq 0) { $Sep += "$Star " } else { $Sep += "$Dot$Dot " }
    }
    Write-Host $Sep -ForegroundColor DarkGray
}

function Write-Banner {
    Clear-Host
    $OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    # Animacion de encendido del cosmo
    $CosmosChars = @(".", "*", "+", "x", "*", "+", ".", "*")
    for ($i = 0; $i -lt 3; $i++) {
        $Line = "  "
        for ($j = 0; $j -lt 60; $j++) {
            $Line += $CosmosChars[(Get-Random -Maximum $CosmosChars.Count)]
        }
        Write-Host $Line -ForegroundColor DarkYellow -NoNewline
        Start-Sleep -Milliseconds 80
        Write-Host "`r" -NoNewline
    }

    Clear-Host
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "         ___  ___  ___ ___    _   _ ___  ___   _ _____ ___ " -ForegroundColor Cyan
    Write-Host "        | _ \| __|/ __| __|  | | | | _ \|   \ /_\_   _| __|" -ForegroundColor Cyan
    Write-Host "        |  _/| _|| (_ | _|   | |_| |  _/| |) / _ \| | | _| " -ForegroundColor Cyan
    Write-Host "        |_|  |___|\___|\___\   \___/|_|  |___/_/ \_\_| |___|" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""

    # Titulo con animacion progresiva
    $Title = "     ACTUALIZADOR DE WINDOWS - HCG"
    foreach ($char in $Title.ToCharArray()) {
        Write-Host $char -NoNewline -ForegroundColor Cyan
        Start-Sleep -Milliseconds 15
    }
    Write-Host ""

    $Subtitle = "     Hospital Civil FAA - Caballeros de Informatica"
    foreach ($char in $Subtitle.ToCharArray()) {
        Write-Host $char -NoNewline -ForegroundColor Magenta
        Start-Sleep -Milliseconds 10
    }
    Write-Host ""

    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host "           Ext. 54425 - Enciende tu cosmo!  " -ForegroundColor DarkYellow
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""

    # Sonido de inicio cosmico
    try {
        [Console]::Beep(523, 100)
        [Console]::Beep(659, 100)
        [Console]::Beep(784, 100)
        [Console]::Beep(1047, 200)
    } catch {}
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N1} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# =============================================================================
# LIMPIEZA DE TAREAS ANTERIORES (por si quedaron de versiones previas)
# =============================================================================

function Remove-OldUpdateTasks {
    $TareasViejas = @("HCG_WinUpdate", "HCG_UpdateLoop", "HCG_AutoUpdate")
    foreach ($tarea in $TareasViejas) {
        try {
            if (Get-ScheduledTask -TaskName $tarea -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $tarea -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "Tarea '$tarea' eliminada (version anterior)" "OK"
            }
        } catch {}
    }
    # Limpiar archivos residuales de loops anteriores
    $ArchivosViejos = @(
        "C:\HCG_Logs\winupdate_loop.ps1",
        "C:\HCG_Logs\winupdate_loop_launcher.vbs",
        "C:\HCG_Logs\winupdate_loop_count.txt",
        "C:\HCG_Logs\winupdate_loop_failed.txt",
        "C:\HCG_Logs\update_loop.ps1",
        "C:\HCG_Logs\update_loop_launcher.vbs",
        "C:\HCG_Logs\update_loop_count.txt",
        "C:\HCG_Logs\update_loop_failed.txt",
        "C:\HCG_Logs\auto_update.ps1",
        "C:\HCG_Logs\auto_update_launcher.vbs"
    )
    foreach ($archivo in $ArchivosViejos) {
        if (Test-Path $archivo) {
            Remove-Item $archivo -Force -ErrorAction SilentlyContinue
        }
    }
}

# =============================================================================
# FLUJO PRINCIPAL
# =============================================================================

# --- Banner cosmico ---
Write-Banner

# --- Crear carpeta de logs ---
if (-not (Test-Path $RutaLogs)) {
    New-Item -ItemType Directory -Path $RutaLogs -Force | Out-Null
}
Write-Log "Inicio de ActualizarWindows.ps1 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "INFO"
Write-Log "Equipo: $env:COMPUTERNAME | Usuario: $env:USERNAME" "INFO"
Write-Separator

# --- Limpiar tareas programadas de versiones anteriores ---
Remove-OldUpdateTasks

# --- Habilitar servicios de Windows Update ---
Write-Host ""
Write-Host "  $([char]0x2734) PREPARANDO SERVICIOS DE WINDOWS UPDATE" -ForegroundColor Yellow
Write-Host ""

$Servicios = @("wuauserv", "BITS")
foreach ($svc in $Servicios) {
    try {
        $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.StartType -eq 'Disabled') {
                Set-Service -Name $svc -StartupType Manual
                Write-Log "Servicio '$svc' habilitado (estaba deshabilitado)" "OK"
            }
            if ($service.Status -ne 'Running') {
                Start-Service -Name $svc -ErrorAction SilentlyContinue
                Write-Log "Servicio '$svc' iniciado" "OK"
            } else {
                Write-Log "Servicio '$svc' ya esta ejecutandose" "OK"
            }
        }
    } catch {
        Write-Log "No se pudo configurar servicio '$svc': $($_.Exception.Message)" "WARN"
    }
}
Write-Separator

# --- Loop principal de actualizaciones ---
$TotalInstaladas = 0
$RequiereReboot = $false

for ($Ronda = 1; $Ronda -le $MaxRondas; $Ronda++) {
    Write-Host ""
    Write-Host "  $([char]0x2734) RONDA $Ronda DE $MaxRondas - BUSCANDO ACTUALIZACIONES" -ForegroundColor Yellow
    Write-Host ""

    try {
        Write-Log "Buscando actualizaciones pendientes (ronda $Ronda)..." "INFO"

        $UpdateSession = New-Object -ComObject Microsoft.Update.Session

        # Paso 1: Buscar TODAS las updates (incluyendo las que requieren EULA)
        $Searcher = $UpdateSession.CreateUpdateSearcher()
        $SearchResult = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

        # Paso 2: Pre-aceptar EULAs de TODAS las updates encontradas
        $EulaCount = 0
        foreach ($Update in $SearchResult.Updates) {
            if (-not $Update.EulaAccepted) {
                try {
                    $Update.AcceptEula()
                    $EulaCount++
                } catch {}
            }
        }
        if ($EulaCount -gt 0) {
            Write-Log "EULAs pre-aceptadas: $EulaCount actualizaciones" "OK"
        }

        $PendingCount = $SearchResult.Updates.Count
        Write-Log "Actualizaciones encontradas: $PendingCount" "INFO"

        # Si no hay updates, equipo actualizado
        if ($PendingCount -eq 0) {
            Write-Host ""
            Write-Host "  $([char]0x2605) EQUIPO 100% ACTUALIZADO $([char]0x2605)" -ForegroundColor Green
            Write-Host ""
            Write-Log "No hay actualizaciones pendientes. Equipo al dia!" "OK"
            break
        }

        # Mostrar lista de updates con tamanio
        Write-Host ""
        Write-Host "  Actualizaciones encontradas:" -ForegroundColor White
        Write-Host ""
        $idx = 1
        foreach ($Update in $SearchResult.Updates) {
            $Size = Format-FileSize $Update.MaxDownloadSize
            Write-Host "    $idx. $($Update.Title)" -ForegroundColor White -NoNewline
            Write-Host " [$Size]" -ForegroundColor DarkGray
            $idx++
        }
        Write-Host ""

        # Crear coleccion (aceptar EULA y filtrar atascadas)
        $Collection = New-Object -ComObject Microsoft.Update.UpdateColl
        $Omitidas = 0
        foreach ($Update in $SearchResult.Updates) {
            $KB = ""
            if ($Update.KBArticleIDs.Count -gt 0) { $KB = $Update.KBArticleIDs.Item(0) }
            if (-not $KB) { $KB = $Update.Title }

            # Si esta KB ya fallo 2+ veces, ocultarla para que no bloquee
            if ($Script:FailedKBs.ContainsKey($KB) -and $Script:FailedKBs[$KB] -ge 2) {
                Write-Log "Omitiendo update atascada (fallo $($Script:FailedKBs[$KB])x): $($Update.Title)" "WARN"
                try { $Update.IsHidden = $true } catch {}
                $Omitidas++
                continue
            }

            # Aceptar EULA automaticamente (critico para Office y otros)
            if (-not $Update.EulaAccepted) {
                try {
                    $Update.AcceptEula()
                } catch {
                    Write-Log "Excepcion al aceptar EULA: $($Update.Title) - $($_.Exception.Message)" "WARN"
                }
                # Verificar que realmente se acepto
                if ($Update.EulaAccepted) {
                    Write-Log "EULA aceptada: $($Update.Title)" "OK"
                } else {
                    Write-Log "EULA NO se pudo aceptar via COM: $($Update.Title) - Se intentara instalar de todos modos" "WARN"
                }
            }
            $Collection.Add($Update) | Out-Null
        }
        if ($Omitidas -gt 0) {
            Write-Log "$Omitidas actualizaciones omitidas por fallos repetidos (ocultadas)" "WARN"
        }

        # Si despues de filtrar no queda nada
        if ($Collection.Count -eq 0) {
            Write-Host ""
            Write-Host "  $([char]0x2605) No quedan actualizaciones instalables $([char]0x2605)" -ForegroundColor Green
            Write-Host ""
            Write-Log "Todas las updates restantes fueron omitidas por fallos repetidos" "WARN"
            break
        }

        # Descargar
        Write-Log "Descargando $PendingCount actualizaciones..." "INFO"
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $Collection
        $DownloadResult = $Downloader.Download()

        $DownloadResultCode = switch ($DownloadResult.ResultCode) {
            0 { "No iniciada" }
            1 { "En progreso" }
            2 { "Completada" }
            3 { "Completada con errores" }
            4 { "Fallida" }
            5 { "Cancelada" }
            default { "Desconocido ($($DownloadResult.ResultCode))" }
        }
        Write-Log "Descarga: $DownloadResultCode" $(if ($DownloadResult.ResultCode -eq 2) { "OK" } else { "WARN" })

        # Instalar
        Write-Log "Instalando $PendingCount actualizaciones..." "INFO"
        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $Collection
        $InstallResult = $Installer.Install()

        $RondaInstaladas = 0
        $RondaFallidas = 0
        for ($i = 0; $i -lt $Collection.Count; $i++) {
            $Result = $InstallResult.GetUpdateResult($i)
            $UpdateObj = $Collection.Item($i)
            $UpdateTitle = $UpdateObj.Title
            $KB = ""
            if ($UpdateObj.KBArticleIDs.Count -gt 0) { $KB = $UpdateObj.KBArticleIDs.Item(0) }
            if (-not $KB) { $KB = $UpdateTitle }

            if ($Result.ResultCode -eq 2) {
                $RondaInstaladas++
                Write-Log "Instalada: $UpdateTitle" "OK"
                # Limpiar de la lista de fallos si antes habia fallado
                if ($Script:FailedKBs.ContainsKey($KB)) { $Script:FailedKBs.Remove($KB) }
            } else {
                $RondaFallidas++
                $HResult = "0x{0:X8}" -f $Result.HResult
                Write-Log "Fallo: $UpdateTitle (codigo: $($Result.ResultCode), HR: $HResult)" "ERROR"
                # Incrementar contador de fallos para esta KB
                if ($Script:FailedKBs.ContainsKey($KB)) {
                    $Script:FailedKBs[$KB]++
                } else {
                    $Script:FailedKBs[$KB] = 1
                }
            }
            if ($Result.RebootRequired) { $RequiereReboot = $true }
        }

        $TotalInstaladas += $RondaInstaladas

        # Resumen de ronda
        Write-Separator
        Write-Host ""
        Write-Host "  Resumen ronda $Ronda`: $RondaInstaladas instaladas, $RondaFallidas fallidas" -ForegroundColor White
        Write-Host ""

        # Si requiere reboot, no podemos continuar con mas rondas interactivas
        if ($RequiereReboot) {
            Write-Log "Se requiere reinicio despues de ronda $Ronda" "WARN"
            break
        }

    } catch {
        Write-Log "Error en ronda $Ronda`: $($_.Exception.Message)" "ERROR"
        break
    }
}

# =============================================================================
# RESUMEN FINAL
# =============================================================================

Write-Host ""
Write-Host "  $([char]0x2605)$([char]0x2605)$([char]0x2605) RESUMEN FINAL $([char]0x2605)$([char]0x2605)$([char]0x2605)" -ForegroundColor Cyan
Write-Host ""
Write-Log "Total de actualizaciones instaladas en esta sesion: $TotalInstaladas" "OK"
Write-Separator

if ($RequiereReboot) {
    Write-Host ""
    Write-Host "  $([char]0x26A0) EL EQUIPO NECESITA REINICIARSE" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Reinicia el equipo manualmente y vuelve a ejecutar este script" -ForegroundColor Cyan
    Write-Host "  para continuar instalando las actualizaciones restantes." -ForegroundColor Cyan
    Write-Host ""
    Write-Log "Reinicio requerido. El tecnico debe reiniciar y ejecutar de nuevo." "WARN"
} else {
    if ($TotalInstaladas -gt 0) {
        Write-Host ""
        Write-Host "  $([char]0x2605) Se instalaron $TotalInstaladas actualizaciones sin requerir reinicio." -ForegroundColor Green
        Write-Host ""
    }
    Write-Host "  $([char]0x2605) Equipo actualizado. No requiere reinicio." -ForegroundColor Green
    Write-Host ""
}

Write-Log "Fin de ActualizarWindows.ps1" "INFO"
Write-Host ""
Write-Host "  Presione una tecla para salir..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
