# =============================================================================
# INSTALADOR DE ACTUALIZACIONES DE WINDOWS - COSMOS UPGRADE
# Hospital Civil de Guadalajara - Soporte Tecnico
# Instala TODAS las actualizaciones: Windows 11 25H2, drivers, .NET, seguridad
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- Deshabilitar QuickEdit Mode (evita que la consola se pause al hacer clic) ---
try {
    $QuickEditCode = @"
using System;
using System.Runtime.InteropServices;
public class ConsoleQuickEdit {
    const uint ENABLE_QUICK_EDIT = 0x0040;
    const int STD_INPUT_HANDLE = -10;
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll")]
    static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll")]
    static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    public static void Disable() {
        IntPtr consoleHandle = GetStdHandle(STD_INPUT_HANDLE);
        uint consoleMode;
        GetConsoleMode(consoleHandle, out consoleMode);
        consoleMode &= ~ENABLE_QUICK_EDIT;
        SetConsoleMode(consoleHandle, consoleMode);
    }
}
"@
    Add-Type -TypeDefinition $QuickEditCode -Language CSharp -ErrorAction SilentlyContinue
    [ConsoleQuickEdit]::Disable()
} catch {
    # Silencioso si falla
}

# =============================================================================
# FUNCIONES DE SONIDO - SAINT SEIYA
# =============================================================================

function Play-StepSound {
    try {
        [Console]::Beep(659, 80)   # Mi
        [Console]::Beep(784, 80)   # Sol
        [Console]::Beep(1047, 120) # Do alto
    } catch {}
}

function Play-ErrorSound {
    try {
        [Console]::Beep(392, 200)  # Sol bajo
        [Console]::Beep(330, 300)  # Mi bajo
    } catch {}
}

function Play-PhaseSound {
    try {
        [Console]::Beep(523, 100)  # Do
        [Console]::Beep(659, 100)  # Mi
        [Console]::Beep(784, 100)  # Sol
        [Console]::Beep(1047, 200) # Do alto
    } catch {}
}

function Play-VictorySound {
    try {
        # Fanfarria de victoria tipo Saint Seiya - Pegasus Fantasy
        [Console]::Beep(523, 120)  # Do
        [Console]::Beep(659, 120)  # Mi
        [Console]::Beep(784, 120)  # Sol
        [Console]::Beep(1047, 200) # Do alto
        Start-Sleep -Milliseconds 50
        [Console]::Beep(988, 120)  # Si
        [Console]::Beep(1047, 120) # Do alto
        [Console]::Beep(1175, 120) # Re alto
        [Console]::Beep(1319, 300) # Mi alto
        Start-Sleep -Milliseconds 100
        [Console]::Beep(1568, 150) # Sol alto
        [Console]::Beep(1319, 150) # Mi alto
        [Console]::Beep(1568, 400) # Sol alto (final)
    } catch {}
}

function Play-DownloadSound {
    try {
        # Sonido cosmico de descarga
        [Console]::Beep(440, 100)  # La
        [Console]::Beep(523, 100)  # Do
        [Console]::Beep(659, 150)  # Mi
    } catch {}
}

function Show-PegasusAnimation {
    Write-Host ""
    Write-Host "       *  .  +  *  .  +  *  .  +  *  .  +  *" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "              ,~~~." -ForegroundColor Cyan
    Write-Host "             (  o  )    PEGASUS SEIYA" -ForegroundColor Cyan
    Write-Host "              )   (     Buscando en el cosmos..." -ForegroundColor Cyan
    Write-Host "             /|   |\" -ForegroundColor Cyan
    Write-Host "            / |   | \" -ForegroundColor Cyan
    Write-Host "           *  *   *  *" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "       *  .  +  *  .  +  *  .  +  *  .  +  *" -ForegroundColor DarkCyan
    Write-Host ""
    Play-PhaseSound
}

function Show-SearchProgress {
    param([string]$Message, [int]$Step)

    $cosmos = @(".", "*", "+", "x", "*", "+", ".", " ")
    $icons = @(
        "[  *  ]",
        "[ * * ]",
        "[* * *]",
        "[ * * ]"
    )

    $icon = $icons[$Step % $icons.Count]
    $stars = ""
    for ($i = 0; $i -lt 5; $i++) {
        $stars += $cosmos[(($Step + $i) % $cosmos.Count)]
    }

    Write-Host "`r  $stars $icon $Message $stars                    " -NoNewline -ForegroundColor Cyan

    # Beep cosmico
    $freq = 400 + (($Step % 4) * 100)
    try { [Console]::Beep($freq, 50) } catch {}
}

# =============================================================================
# FUNCIONES DE VISUALIZACION
# =============================================================================

function Show-CosmosBanner {
    Clear-Host
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "       ================================================" -ForegroundColor Cyan
    Write-Host "       |                                              |" -ForegroundColor Cyan
    Write-Host "       |   COSMOS UPDATE - WINDOWS UPDATE             |" -ForegroundColor Cyan
    Write-Host "       |                                              |" -ForegroundColor Cyan
    Write-Host "       |      Hospital Civil de Guadalajara           |" -ForegroundColor Cyan
    Write-Host "       |         Soporte Tecnico - iA                 |" -ForegroundColor Cyan
    Write-Host "       |                                              |" -ForegroundColor Cyan
    Write-Host "       ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *" -ForegroundColor DarkYellow
    Write-Host ""
}

function Write-CosmosLog {
    param([string]$Mensaje, [string]$Tipo = "INFO")
    $Color = switch ($Tipo) {
        "OK"    { "Green" }
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "PHASE" { "Cyan" }
        "PROG"  { "Magenta" }
        default { "White" }
    }
    $Icon = switch ($Tipo) {
        "OK"    { "[OK]" }
        "ERROR" { "[X]" }
        "WARN"  { "[!]" }
        "PHASE" { "[*]" }
        "PROG"  { "[>]" }
        default { "[-]" }
    }
    Write-Host "  $Icon $Mensaje" -ForegroundColor $Color
}

function Show-ProgressBar {
    param([int]$Percent, [string]$Status)
    $Width = 40
    $Complete = [math]::Round($Width * $Percent / 100)
    $Remaining = $Width - $Complete
    $Bar = "[" + ("=" * $Complete) + (" " * $Remaining) + "]"
    Write-Host "`r  $Bar $Percent% - $Status                    " -NoNewline -ForegroundColor Cyan
}

# =============================================================================
# FUNCION PRINCIPAL DE ACTUALIZACION
# =============================================================================

function Install-AllUpdates {
    Write-Host ""
    Play-PhaseSound
    Write-CosmosLog "FASE 1: Buscando actualizaciones disponibles..." "PHASE"

    # Mostrar Pegasus
    Show-PegasusAnimation

    try {
        # Crear sesion de Windows Update
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

        # Animacion mientras busca
        Write-Host ""
        $searchMessages = @(
            "Conectando al Santuario de Microsoft...",
            "Elevando el cosmo...",
            "Consultando a Athena...",
            "Escaneando las 12 casas...",
            "Buscando caballeros dorados...",
            "Analizando constelaciones...",
            "Pegasus Ryu Sei Ken...",
            "Despertando el septimo sentido..."
        )

        # Iniciar busqueda en runspace separado
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.Open()
        $runspace.SessionStateProxy.SetVariable('UpdateSearcher', $UpdateSearcher)

        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        $ps.AddScript({
            $UpdateSearcher.Search("IsInstalled=0 and IsHidden=0")
        }) | Out-Null

        $handle = $ps.BeginInvoke()

        # Mostrar animacion mientras busca
        $step = 0
        while (-not $handle.IsCompleted) {
            $msg = $searchMessages[$step % $searchMessages.Count]
            Show-SearchProgress -Message $msg -Step $step
            Start-Sleep -Milliseconds 300
            $step++
        }

        # Obtener resultado
        $SearchResult = $ps.EndInvoke($handle)
        $ps.Dispose()
        $runspace.Close()

        Write-Host "`r                                                                        "
        Write-Host ""
        Play-StepSound
        Write-CosmosLog "Busqueda completada!" "OK"

        $Updates = $SearchResult.Updates

        if ($Updates.Count -eq 0) {
            Write-Host ""
            Write-Host "  =========================================================" -ForegroundColor Green
            Write-Host "  |                                                       |" -ForegroundColor Green
            Write-Host "  |   NO HAY ACTUALIZACIONES PENDIENTES                   |" -ForegroundColor Green
            Write-Host "  |   El sistema esta completamente actualizado           |" -ForegroundColor Green
            Write-Host "  |                                                       |" -ForegroundColor Green
            Write-Host "  =========================================================" -ForegroundColor Green
            return $false
        }

        # Mostrar actualizaciones encontradas
        Write-Host ""
        Write-CosmosLog "Se encontraron $($Updates.Count) actualizaciones:" "OK"
        Write-Host ""
        Write-Host "  ---------------------------------------------------------" -ForegroundColor DarkGray

        $TotalSize = 0
        $Counter = 1
        foreach ($Update in $Updates) {
            $Size = [math]::Round($Update.MaxDownloadSize / 1MB, 1)
            $TotalSize += $Update.MaxDownloadSize
            $Title = $Update.Title
            if ($Title.Length -gt 55) { $Title = $Title.Substring(0, 52) + "..." }
            Write-Host "  $Counter. $Title" -ForegroundColor White
            Write-Host "     Tamano: $Size MB" -ForegroundColor DarkGray
            $Counter++
        }

        Write-Host "  ---------------------------------------------------------" -ForegroundColor DarkGray
        $TotalSizeMB = [math]::Round($TotalSize / 1MB, 1)
        Write-Host "  Total a descargar: $TotalSizeMB MB" -ForegroundColor Yellow
        Write-Host ""

        Write-Host "  =========================================================" -ForegroundColor Green
        Write-Host "  |  INSTALACION AUTOMATICA - NO REQUIERE INTERVENCION    |" -ForegroundColor Green
        Write-Host "  =========================================================" -ForegroundColor Green
        Write-Host ""
        Play-StepSound

        # Crear coleccion de actualizaciones
        Write-Host ""
        Play-PhaseSound
        Write-CosmosLog "FASE 2: Descargando actualizaciones..." "PHASE"
        Write-Host ""

        $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($Update in $Updates) {
            $UpdatesToDownload.Add($Update) | Out-Null
        }

        # Descargar
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesToDownload

        Write-CosmosLog "Iniciando descarga de $($Updates.Count) actualizaciones..." "PROG"
        Write-Host ""

        # Mostrar progreso de descarga
        $DownloadJob = $Downloader.Download()

        Play-DownloadSound
        Write-CosmosLog "Descarga completada" "OK"
        Write-Host ""

        # Instalar
        Play-PhaseSound
        Write-CosmosLog "FASE 3: Instalando actualizaciones..." "PHASE"
        Write-Host ""

        $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($Update in $Updates) {
            if ($Update.IsDownloaded) {
                $UpdatesToInstall.Add($Update) | Out-Null
            }
        }

        if ($UpdatesToInstall.Count -eq 0) {
            Write-CosmosLog "No hay actualizaciones descargadas para instalar" "ERROR"
            return $false
        }

        Write-CosmosLog "Instalando $($UpdatesToInstall.Count) actualizaciones..." "PROG"
        Write-Host ""
        Write-Host "  =========================================================" -ForegroundColor Yellow
        Write-Host "  |  INSTALACION EN PROGRESO - NO APAGAR EL EQUIPO        |" -ForegroundColor Yellow
        Write-Host "  =========================================================" -ForegroundColor Yellow
        Write-Host ""

        # Instalar una por una para mostrar progreso
        $Installed = 0
        $Failed = 0
        $NeedsReboot = $false
        $Total = $UpdatesToInstall.Count
        $Current = 0

        foreach ($Update in $UpdatesToInstall) {
            $Current++
            $SingleUpdate = New-Object -ComObject Microsoft.Update.UpdateColl
            $SingleUpdate.Add($Update) | Out-Null

            # Crear instalador usando la sesion (forma correcta)
            $SingleInstaller = $UpdateSession.CreateUpdateInstaller()
            $SingleInstaller.Updates = $SingleUpdate

            $Title = $Update.Title
            if ($Title.Length -gt 45) { $Title = $Title.Substring(0, 42) + "..." }

            Write-Host "  [$Current/$Total] Instalando: $Title" -ForegroundColor Cyan

            try {
                $Result = $SingleInstaller.Install()

                if ($Result.ResultCode -eq 2) {
                    Play-StepSound
                    Write-Host "      [OK] Instalado correctamente" -ForegroundColor Green
                    $Installed++
                } elseif ($Result.ResultCode -eq 3) {
                    Play-StepSound
                    Write-Host "      [OK] Instalado - Requiere reinicio" -ForegroundColor Yellow
                    $Installed++
                    $NeedsReboot = $true
                } else {
                    Play-ErrorSound
                    Write-Host "      [X] Error (codigo: $($Result.ResultCode))" -ForegroundColor Red
                    $Failed++
                }
            } catch {
                Play-ErrorSound
                Write-Host "      [X] Error: $($_.Exception.Message)" -ForegroundColor Red
                $Failed++
            }
            Write-Host ""
        }

        # Resumen
        Write-Host ""
        Play-VictorySound
        Write-Host "  =========================================================" -ForegroundColor Green
        Write-Host "  |              RESUMEN DE INSTALACION                   |" -ForegroundColor Green
        Write-Host "  =========================================================" -ForegroundColor Green
        Write-Host ""
        Write-CosmosLog "Actualizaciones instaladas: $Installed" "OK"
        if ($Failed -gt 0) {
            Write-CosmosLog "Actualizaciones fallidas: $Failed" "WARN"
        }

        if ($NeedsReboot) {
            Write-Host ""
            Write-Host "  =========================================================" -ForegroundColor Yellow
            Write-Host "  |  REINICIO AUTOMATICO EN 15 SEGUNDOS                   |" -ForegroundColor Yellow
            Write-Host "  |                                                       |" -ForegroundColor Yellow
            Write-Host "  |  Las actualizaciones se aplicaran al reiniciar        |" -ForegroundColor Yellow
            Write-Host "  |  NO APAGUES EL EQUIPO DURANTE EL PROCESO              |" -ForegroundColor Yellow
            Write-Host "  =========================================================" -ForegroundColor Yellow
            Write-Host ""

            # Cuenta regresiva visual
            for ($i = 15; $i -ge 1; $i--) {
                Write-Host "`r  Reiniciando en $i segundos...   " -NoNewline -ForegroundColor Cyan
                [Console]::Beep(800, 100)
                Start-Sleep -Seconds 1
            }
            Write-Host ""
            Write-CosmosLog "Reiniciando ahora..." "WARN"

            # Forzar reinicio (cierra todas las ventanas incluyendo el Assistant)
            shutdown /r /t 0 /f
        }

        return $true

    } catch {
        Write-CosmosLog "Error: $($_.Exception.Message)" "ERROR"

        # Fallback: usar usoclient
        Write-Host ""
        Write-CosmosLog "Intentando metodo alternativo (usoclient)..." "WARN"
        Write-Host ""

        Write-CosmosLog "Buscando actualizaciones..." "PROG"
        Start-Process -FilePath "usoclient.exe" -ArgumentList "StartInteractiveScan" -Wait -WindowStyle Hidden

        Write-CosmosLog "Iniciando descarga..." "PROG"
        Start-Process -FilePath "usoclient.exe" -ArgumentList "StartDownload" -Wait -WindowStyle Hidden

        Write-CosmosLog "Iniciando instalacion..." "PROG"
        Start-Process -FilePath "usoclient.exe" -ArgumentList "StartInstall" -Wait -WindowStyle Hidden

        Write-Host ""
        Write-CosmosLog "Proceso iniciado via Windows Update" "OK"
        Write-CosmosLog "Revisa Configuracion > Windows Update para ver el progreso" "INFO"

        return $true
    }
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Show-CosmosBanner

# Verificar conexion a Internet
Write-CosmosLog "Verificando conexion a Internet..." "INFO"
$Internet = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet -ErrorAction SilentlyContinue
if (-not $Internet) {
    $Internet = Test-Connection -ComputerName "dns.google" -Count 2 -Quiet -ErrorAction SilentlyContinue
}
if (-not $Internet) {
    Write-CosmosLog "No hay conexion a Internet" "ERROR"
    Write-Host ""
    Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}
Write-CosmosLog "Conexion a Internet: OK" "OK"

# Ejecutar instalacion
$Result = Install-AllUpdates

Write-Host ""
Write-Host "  Presiona cualquier tecla para salir..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
