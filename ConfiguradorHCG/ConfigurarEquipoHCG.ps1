# =============================================================================
# CONFIGURADOR DE EQUIPOS - HOSPITAL CIVIL FAA v4.0
# =============================================================================
# Se eleva automaticamente como Administrador si no lo es
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# --- DESHABILITAR QUICKEDIT MODE (evita que el script se pause al hacer clic) ---
# Esto previene que la consola entre en modo seleccion cuando el usuario hace clic
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class ConsoleQuickEdit {
    const uint ENABLE_QUICK_EDIT = 0x0040;
    const uint ENABLE_EXTENDED_FLAGS = 0x0080;
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
        if (GetConsoleMode(consoleHandle, out consoleMode)) {
            consoleMode &= ~ENABLE_QUICK_EDIT;
            consoleMode |= ENABLE_EXTENDED_FLAGS;
            SetConsoleMode(consoleHandle, consoleMode);
        }
    }
}
"@ -ErrorAction SilentlyContinue

try { [ConsoleQuickEdit]::Disable() } catch { }

# --- CONFIGURACION ---
$Servidor = "10.2.1.13"
$Usuario = "2010201"
$Password = "7v3l73v37nG06"

# Ruta base del pack de instalacion
$RutaBase = "\\$Servidor\soportefaa\pack_installer_iA"
$RutaAccesos = "$RutaBase\accesos_directos"
$RutaAcrobat = "$RutaBase\acrobat_reader"
$RutaAntivirus = "$RutaBase\antivirus"
$RutaChrome = "$RutaBase\chrome"
$RutaDedalus = "$RutaBase\dedalus_expedienteclinico"
$RutaOffice = "$RutaBase\office"
$RutaWallpaper = "$RutaBase\wallpaper"
$RutaDotNet = "$RutaBase\netframework3_5\sxs"
$RutaWinRAR = "$RutaBase\winrar_licence"

$ZonaHoraria = "Central Standard Time (Mexico)"
$UsuarioSoporte = "Soporte"
$PasswordSoporte = "*TIsoporte"
$RutaLogs = "C:\HCG_Logs"

$GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"

# Variable global para tracking de software instalado
$Script:SoftwareInstalado = @()
$Script:UsuarioOriginal = $env:USERNAME
$Script:Departamento = "Soporte Tecnico - Ext. 54425"
$Script:EsOPD = $false

# =============================================================================
# PEGASUS FANTASY - Notas de la melodia (Saint Seiya)
# Cada paso completado toca la siguiente nota de la melodia
# AUDIO WAV - Se reproduce por el dispositivo de audio predeterminado
# =============================================================================

# Generador de tonos WAV (audio real)
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Media;

public class PegasusWavPlayer {
    public static void PlayTone(double frequency, int durationMs, double volume) {
        if (frequency <= 0) return;

        int sampleRate = 22050;
        int samples = (int)(sampleRate * durationMs / 1000.0);

        using (MemoryStream ms = new MemoryStream()) {
            using (BinaryWriter bw = new BinaryWriter(ms)) {
                // Escribir cabecera WAV
                bw.Write(new char[] {'R','I','F','F'});
                bw.Write(36 + samples * 2);
                bw.Write(new char[] {'W','A','V','E'});
                bw.Write(new char[] {'f','m','t',' '});
                bw.Write(16);
                bw.Write((short)1);
                bw.Write((short)1);
                bw.Write(sampleRate);
                bw.Write(sampleRate * 2);
                bw.Write((short)2);
                bw.Write((short)16);
                bw.Write(new char[] {'d','a','t','a'});
                bw.Write(samples * 2);

                // Generar onda senoidal con fade out
                for (int i = 0; i < samples; i++) {
                    double t = (double)i / sampleRate;
                    double fadeOut = 1.0 - ((double)i / samples) * 0.3;
                    short sample = (short)(Math.Sin(2 * Math.PI * frequency * t) * 32767 * volume * fadeOut);
                    bw.Write(sample);
                }

                ms.Position = 0;
                using (SoundPlayer sp = new SoundPlayer(ms)) {
                    sp.PlaySync();
                }
            }
        }
    }
}
"@ -ReferencedAssemblies "System.dll" -ErrorAction SilentlyContinue

$Script:PegasusNotas = @(
    @{f=659; d=150},  # E5
    @{f=659; d=150},  # E5
    @{f=587; d=150},  # D5
    @{f=659; d=250},  # E5
    @{f=784; d=150},  # G5
    @{f=784; d=150},  # G5
    @{f=698; d=150},  # F5
    @{f=659; d=250},  # E5
    @{f=659; d=150},  # E5
    @{f=659; d=150},  # E5
    @{f=587; d=150},  # D5
    @{f=659; d=200},  # E5
    @{f=587; d=100},  # D5
    @{f=523; d=100},  # C5
    @{f=494; d=250},  # B4
    @{f=440; d=150},  # A4
    @{f=494; d=150},  # B4
    @{f=523; d=150},  # C5
    @{f=587; d=150},  # D5
    @{f=659; d=400},  # E5 (nota final epica)
    @{f=784; d=200},  # G5 (bonus)
    @{f=880; d=300}   # A5 (final!)
)
$Script:PegasusIndex = 0
$Script:AudioVolume = 0.25  # Volumen (0.0 a 1.0)

# Funcion para tocar la siguiente nota de Pegasus Fantasy
function Play-PegasusNote {
    try {
        if ($Script:PegasusIndex -lt $Script:PegasusNotas.Count) {
            $nota = $Script:PegasusNotas[$Script:PegasusIndex]
            [PegasusWavPlayer]::PlayTone($nota.f, $nota.d, $Script:AudioVolume)
            $Script:PegasusIndex++
        } else {
            # Si ya termino la melodia, reiniciar para seguir tocando
            $Script:PegasusIndex = 0
            $nota = $Script:PegasusNotas[$Script:PegasusIndex]
            [PegasusWavPlayer]::PlayTone($nota.f, $nota.d, $Script:AudioVolume)
            $Script:PegasusIndex++
        }
    } catch {}
}

# Funcion para sonidos de error/advertencia (tambien audio real)
function Play-AlertSound {
    param([string]$Type = "ERROR")
    try {
        switch ($Type) {
            "ERROR" {
                [PegasusWavPlayer]::PlayTone(330, 150, 0.2)
                [PegasusWavPlayer]::PlayTone(220, 200, 0.2)
            }
            "WARN"  { [PegasusWavPlayer]::PlayTone(440, 100, 0.15) }
        }
    } catch {}
}

# ============================================================================
# SISTEMA DE MELODIA AMBIENTAL DE FONDO (para procesos largos)
# Usa un RUNSPACE separado para que el audio no se detenga durante Start-Process -Wait
# Inspirado en los momentos contemplativos del Santuario - Saint Seiya
# ============================================================================
$Script:AmbientRunspace = $null
$Script:AmbientPowerShell = $null
$Script:AmbientWavPath = "$env:TEMP\cosmos_ambient.wav"

# Generar WAV ambiental largo (melodia suave tipo Sanctuary)
function New-AmbientWav {
    try {
        $sampleRate = 22050
        $duracionTotal = 20  # 20 segundos de loop
        $samples = $sampleRate * $duracionTotal

        # Melodia ambiental continua - estilo Sanctuary/Saint Seiya
        $notasAmbient = @(
            @{F=262; Start=0;    Dur=1.8},   # Do4
            @{F=330; Start=1.8;  Dur=1.8},   # Mi4
            @{F=392; Start=3.6;  Dur=1.8},   # Sol4
            @{F=440; Start=5.4;  Dur=2.2},   # La4 - punto alto
            @{F=392; Start=7.6;  Dur=1.8},   # Sol4
            @{F=349; Start=9.4;  Dur=1.8},   # Fa4
            @{F=330; Start=11.2; Dur=1.8},   # Mi4
            @{F=294; Start=13.0; Dur=1.8},   # Re4
            @{F=262; Start=14.8; Dur=2.0},   # Do4
            @{F=294; Start=16.8; Dur=1.6},   # Re4
            @{F=330; Start=18.4; Dur=1.6}    # Mi4 - prepara loop
        )

        $volumen = 0.35  # Volumen audible

        $ms = New-Object System.IO.MemoryStream
        $bw = New-Object System.IO.BinaryWriter($ms)

        # Cabecera WAV
        $bw.Write([char[]]@('R','I','F','F'))
        $bw.Write([int](36 + $samples * 2))
        $bw.Write([char[]]@('W','A','V','E'))
        $bw.Write([char[]]@('f','m','t',' '))
        $bw.Write([int]16)
        $bw.Write([int16]1)
        $bw.Write([int16]1)
        $bw.Write([int]$sampleRate)
        $bw.Write([int]($sampleRate * 2))
        $bw.Write([int16]2)
        $bw.Write([int16]16)
        $bw.Write([char[]]@('d','a','t','a'))
        $bw.Write([int]($samples * 2))

        for ($i = 0; $i -lt $samples; $i++) {
            $t = [double]$i / $sampleRate
            $sampleValue = 0

            foreach ($nota in $notasAmbient) {
                if ($t -ge $nota.Start -and $t -lt ($nota.Start + $nota.Dur) -and $nota.F -gt 0) {
                    $notaT = $t - $nota.Start
                    $fadeIn = [Math]::Min(1.0, $notaT / 0.3)
                    $fadeOut = [Math]::Min(1.0, ($nota.Dur - $notaT) / 0.5)
                    $sampleValue = [Math]::Sin(2 * [Math]::PI * $nota.F * $t) * 32767 * $volumen * $fadeIn * $fadeOut
                    break
                }
            }
            $bw.Write([int16]$sampleValue)
        }

        $bytes = $ms.ToArray()
        [System.IO.File]::WriteAllBytes($Script:AmbientWavPath, $bytes)
        $bw.Dispose()
        $ms.Dispose()
        return $true
    } catch {
        return $false
    }
}

# Iniciar melodia ambiental en RUNSPACE SEPARADO (no se bloquea con Start-Process -Wait)
function Start-BackgroundMelody {
    param([string]$Mensaje = "Proceso en curso...")

    try {
        # Detener si ya hay una corriendo
        Stop-BackgroundMelody

        # Regenerar WAV
        if (Test-Path $Script:AmbientWavPath) {
            Remove-Item $Script:AmbientWavPath -Force -ErrorAction SilentlyContinue
        }
        $null = New-AmbientWav

        if (Test-Path $Script:AmbientWavPath) {
            # Crear Runspace separado para reproducir audio
            $Script:AmbientRunspace = [runspacefactory]::CreateRunspace()
            $Script:AmbientRunspace.ApartmentState = "STA"
            $Script:AmbientRunspace.Open()

            $Script:AmbientPowerShell = [powershell]::Create()
            $Script:AmbientPowerShell.Runspace = $Script:AmbientRunspace

            # Script que se ejecuta en el runspace separado
            $audioScript = {
                param($wavPath)
                Add-Type -AssemblyName System.Windows.Forms
                $player = New-Object System.Media.SoundPlayer($wavPath)
                $player.PlayLooping()
                # Mantener el runspace vivo mientras reproduce
                while ($true) {
                    Start-Sleep -Milliseconds 500
                }
            }

            $null = $Script:AmbientPowerShell.AddScript($audioScript)
            $null = $Script:AmbientPowerShell.AddArgument($Script:AmbientWavPath)
            $Script:AmbientHandle = $Script:AmbientPowerShell.BeginInvoke()

            Write-Host "  [$([char]0x266B)] $Mensaje" -ForegroundColor DarkCyan
        }
    } catch {
        # Silencioso si falla
    }
}

# Detener melodia ambiental
function Stop-BackgroundMelody {
    try {
        if ($Script:AmbientPowerShell) {
            $Script:AmbientPowerShell.Stop()
            $Script:AmbientPowerShell.Dispose()
            $Script:AmbientPowerShell = $null
        }
        if ($Script:AmbientRunspace) {
            $Script:AmbientRunspace.Close()
            $Script:AmbientRunspace.Dispose()
            $Script:AmbientRunspace = $null
        }
    } catch { }
}

# Nombres de grupos locales via SID (independiente del idioma de Windows)
$GrupoAdmin = (New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-544")).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
$GrupoUsuarios = (New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-545")).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]

# --- FUNCIONES AUXILIARES ---

function Write-Log {
    param([string]$Mensaje, [string]$Tipo = "INFO")
    $Icon = switch ($Tipo) {
        "OK"    { [char]0x2605 }  # Estrella dorada
        "ERROR" { [char]0x2716 }  # Cruz roja
        "WARN"  { [char]0x26A0 }  # Triangulo alerta
        default { [char]0x2192 }  # Flecha cosmica
    }
    $Color = switch ($Tipo) { "OK" { "Green" } "ERROR" { "Red" } "WARN" { "Yellow" } default { "Cyan" } }
    Write-Host "  $Icon [$Tipo] $Mensaje" -ForegroundColor $Color
    if (-not (Test-Path $RutaLogs)) { New-Item -ItemType Directory -Path $RutaLogs -Force | Out-Null }
    Add-Content -Path "$RutaLogs\config.log" -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Tipo] $Mensaje"

    # Sonidos segun tipo - Pegasus Fantasy para exitos! (Audio real WAV)
    try {
        switch ($Tipo) {
            "OK"    { Play-PegasusNote }  # Toca siguiente nota de Pegasus Fantasy
            "ERROR" { Play-AlertSound -Type "ERROR" }
            "WARN"  { Play-AlertSound -Type "WARN" }
        }
    } catch {}
}

function Show-Banner {
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
    Write-Host "                    /\        /\" -ForegroundColor DarkYellow
    Write-Host "                   /  \  /\  /  \" -ForegroundColor DarkYellow
    Write-Host "                  / /\ \/ /\/ /\ \" -ForegroundColor DarkYellow
    Write-Host "                 / /  \  / /\/  \ \" -ForegroundColor DarkYellow
    Write-Host "                /_/    \/\/ /    \_\" -ForegroundColor DarkYellow
    Write-Host "                \  PEGASUS  \/    /" -ForegroundColor Cyan
    Write-Host "                 \   / /\  /\   /" -ForegroundColor DarkYellow
    Write-Host "                  \_/ /  \/  \_/" -ForegroundColor DarkYellow
    Write-Host "                     /   /\" -ForegroundColor DarkYellow
    Write-Host "                    /   /  \" -ForegroundColor DarkYellow
    Write-Host "                   /   /    \" -ForegroundColor DarkYellow
    Write-Host "                  /___/______\" -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  .  +  *  ." -ForegroundColor DarkYellow
    Write-Host ""

    # Titulo epico con animacion progresiva
    $Title = "     CONFIGURADOR COSMICO DE EQUIPOS v4.0"
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
        [Console]::Beep(523, 100)  # Do
        [Console]::Beep(659, 100)  # Mi
        [Console]::Beep(784, 100)  # Sol
        [Console]::Beep(1047, 200) # Do alto
    } catch {}
}

function Write-StepHeader {
    param([int]$Step, [int]$Total = 25, [string]$Title)
    $Stars = [char]0x2605  # Estrella
    $Cosmos = [char]0x2734 # Estrella de 8 puntas
    $Filled = [char]0x2593 # Bloque lleno
    $Empty  = [char]0x2591 # Bloque vacio

    Write-Host ""

    # Barra de progreso cosmica
    $Pct = [math]::Floor(($Step / $Total) * 30)
    $Bar = ""
    for ($i = 0; $i -lt 30; $i++) {
        if ($i -lt $Pct) { $Bar += $Filled } else { $Bar += $Empty }
    }
    Write-Host "  $Cosmos " -NoNewline -ForegroundColor Magenta
    Write-Host "[$Bar]" -NoNewline -ForegroundColor DarkYellow
    Write-Host " $([math]::Floor(($Step / $Total) * 100))%" -ForegroundColor DarkYellow

    # Titulo del paso
    Write-Host "  $Stars " -NoNewline -ForegroundColor DarkYellow
    Write-Host "[$Step/$Total] " -NoNewline -ForegroundColor DarkYellow
    Write-Host "$Title" -ForegroundColor Cyan

    # Separador cosmico
    Write-Separator
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

function Play-StepSound {
    try {
        [Console]::Beep(659, 80)   # Mi
        [Console]::Beep(784, 80)   # Sol
        [Console]::Beep(1047, 120) # Do alto
    } catch {}
}

function Play-VictorySound {
    try {
        # Fanfarria de victoria tipo Saint Seiya
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

function Play-PhoenixMelody {
    # Melodia del Ave Fenix - Ikki (epica pero suave)
    # Inspirada en "Phoenix Ikki Theme" - notas ascendentes como el renacimiento
    try {
        # Intro suave - el fenix despierta
        [Console]::Beep(330, 300)   # Mi - despertar
        Start-Sleep -Milliseconds 100
        [Console]::Beep(392, 300)   # Sol
        [Console]::Beep(440, 400)   # La - ascenso
        Start-Sleep -Milliseconds 150

        # Crecimiento - las alas se abren
        [Console]::Beep(523, 350)   # Do5 - fuerza
        [Console]::Beep(587, 300)   # Re5
        [Console]::Beep(659, 500)   # Mi5 - vuelo
        Start-Sleep -Milliseconds 200

        # Climax suave - el fenix renace
        [Console]::Beep(784, 400)   # Sol5 - majestuoso
        [Console]::Beep(880, 350)   # La5
        [Console]::Beep(784, 300)   # Sol5 - descenso suave
        [Console]::Beep(659, 600)   # Mi5 - paz
        Start-Sleep -Milliseconds 300

        # Final - alas extendidas
        [Console]::Beep(523, 400)   # Do5
        [Console]::Beep(659, 500)   # Mi5
        [Console]::Beep(784, 800)   # Sol5 - sostenido final
    } catch {}
}

function Show-PhoenixRebootCountdown {
    param([int]$Seconds = 60)

    Clear-Host

    $phoenix = @'

                           . - ~ ~ ~ - .
                       . '       _      ' .
                      /        .'         \
                     /        /             \
                    |        |   .-'''-.    |
                    |        |  /   _   \   |
                     \       | |   (_)   | /
                      \      |  \       / /
                       \      \  '-----' /
                        \      '.     .'/
                         \   .  '---'   /
                          \ / \       / \
                      .-'  |   '-----'   '.
                    .'      \             /
                   /     .-' '._     _.' '-.
                  /    .'       '---'       \
                 /   .'                       \
                |  .'      A V E   F E N I X   '.
                |.'                              \
               .'     EL SISTEMA RENACERA...      '.
              /                                     \
             |      *  *  *  *  *  *  *  *  *  *     |
              \                                     /
               '.                                 .'
                 ' .    Hospital Civil GDL    . '
                    ' - . _   _ . - '

'@

    Write-Host $phoenix -ForegroundColor Red
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Yellow
    Write-Host "  |     REINICIO AUTOMATICO - ACTUALIZACIONES SE APLICARAN      |" -ForegroundColor Yellow
    Write-Host "  |                 NO APAGUES EL EQUIPO                        |" -ForegroundColor Yellow
    Write-Host "  ================================================================" -ForegroundColor Yellow
    Write-Host ""

    # Tocar melodia del Fenix
    Play-PhoenixMelody

    # Cuenta regresiva con beeps suaves
    for ($i = $Seconds; $i -ge 1; $i--) {
        $bar = "=" * [math]::Floor((($Seconds - $i) / $Seconds) * 40)
        $space = " " * (40 - $bar.Length)

        Write-Host "`r  [$bar$space] $i segundos para reiniciar...   " -NoNewline -ForegroundColor Cyan

        # Beep cada 10 segundos, y cada segundo en los ultimos 5
        if ($i -le 5 -or $i % 10 -eq 0) {
            try { [Console]::Beep(600 + (($Seconds - $i) * 10), 100) } catch {}
        }

        Start-Sleep -Seconds 1
    }

    Write-Host ""
    Write-Host ""
    Write-Host "  *** EL FENIX RENACE - REINICIANDO AHORA ***" -ForegroundColor Red
    Write-Host ""
}

function Show-ProgressCosmos {
    param([int]$Step, [int]$Total = 25)
    $Pct = [math]::Floor(($Step / $Total) * 100)
    $title = "Cosmo: $Pct% - Paso $Step/$Total"
    $Host.UI.RawUI.WindowTitle = "$([char]0x2605) $title $([char]0x2605)"
}

function Show-CosmosAnimation {
    param([string]$Message = "Encendiendo el cosmo...")
    $Frames = @(
        "    .  *  .     ",
        "   * . + . *    ",
        "  . + * + .  *  ",
        " * . + * + . *  ",
        "  . + * + .  *  ",
        "   * . + . *    ",
        "    .  *  .     "
    )
    foreach ($frame in $Frames) {
        Write-Host "`r  $frame $Message" -NoNewline -ForegroundColor DarkYellow
        Start-Sleep -Milliseconds 80
    }
    Write-Host ""
}

function Find-Installer {
    param([string]$Carpeta, [string]$Filtro = "*.exe")
    if (Test-Path $Carpeta) {
        $Archivo = Get-ChildItem -Path $Carpeta -Filter $Filtro -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($Archivo) { return $Archivo.FullName }
    }
    return $null
}

function Pin-ToTaskbar {
    param([string]$ShortcutPath, [string]$AppName = "App")

    try {
        # Metodo para Windows 10/11: Copiar a la carpeta de TaskBar
        $TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

        if (-not (Test-Path $TaskBarFolder)) {
            New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null
        }

        if (Test-Path $ShortcutPath) {
            $FileName = Split-Path $ShortcutPath -Leaf
            $Destino = "$TaskBarFolder\$FileName"
            Copy-Item -Path $ShortcutPath -Destination $Destino -Force
            Write-Log "$AppName anclado a barra de tareas" "OK"
        }

        # Metodo alternativo usando Shell
        $Shell = New-Object -ComObject Shell.Application
        $Folder = $Shell.Namespace((Split-Path $ShortcutPath -Parent))
        $Item = $Folder.ParseName((Split-Path $ShortcutPath -Leaf))

        if ($Item) {
            $Verbs = $Item.Verbs()
            foreach ($Verb in $Verbs) {
                if ($Verb.Name -match "Anclar a la barra de tareas|Pin to taskbar|Anclar al inicio") {
                    $Verb.DoIt()
                    break
                }
            }
        }
    } catch {
        Write-Log "No se pudo anclar $AppName a la barra de tareas" "WARN"
    }
}

function Create-TaskbarShortcut {
    param(
        [string]$TargetPath,
        [string]$ShortcutName,
        [string]$IconPath = "",
        [string]$Arguments = ""
    )

    $TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $ShortcutPath = "$TaskBarFolder\$ShortcutName.lnk"

    if (-not (Test-Path $TaskBarFolder)) {
        New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null
    }

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    if ($Arguments) { $Shortcut.Arguments = $Arguments }
    if ($IconPath -and (Test-Path $IconPath)) { $Shortcut.IconLocation = $IconPath }
    $Shortcut.Save()

    return $ShortcutPath
}

# Obtiene el nombre del grupo "Users/Usuarios" independiente del idioma de Windows
function Get-UsersGroupName {
    try {
        # SID S-1-5-32-545 = BUILTIN\Users (funciona en cualquier idioma)
        $UsersSID = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-545")
        return $UsersSID.Translate([System.Security.Principal.NTAccount]).Value
    } catch {
        return "Users"  # Fallback
    }
}

# Crea la carpeta C:\HCG_Logs con permisos para todos los usuarios
function Initialize-HCGLogsFolder {
    $LogsFolder = "C:\HCG_Logs"
    if (-not (Test-Path $LogsFolder)) {
        New-Item -ItemType Directory -Path $LogsFolder -Force | Out-Null
    }
    # Establecer permisos para que todos los usuarios puedan leer y ejecutar
    try {
        $UsersGroup = Get-UsersGroupName
        $Acl = Get-Acl $LogsFolder
        $Rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $UsersGroup, "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
        $Acl.SetAccessRule($Rule)
        Set-Acl -Path $LogsFolder -AclObject $Acl -ErrorAction SilentlyContinue
    } catch {}
    return $LogsFolder
}

# Crea un script VBS que lanza PowerShell de forma completamente oculta
# Esto evita el destello de ventana que ocurre con Task Scheduler + PowerShell -WindowStyle Hidden
function New-HiddenLauncher {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PowerShellScriptPath,
        [string]$VbsPath = ""
    )

    if (-not $VbsPath) {
        $VbsPath = $PowerShellScriptPath -replace '\.ps1$', '_launcher.vbs'
    }

    $VbsContent = @"
' Launcher oculto para $PowerShellScriptPath
' Ejecuta PowerShell completamente invisible (sin destello de ventana)
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -NonInteractive -File ""$PowerShellScriptPath""", 0, False
Set objShell = Nothing
"@

    $VbsContent | Out-File -FilePath $VbsPath -Encoding ASCII -Force

    # Establecer permisos para que todos los usuarios puedan ejecutar el VBS
    try {
        $UsersGroup = Get-UsersGroupName
        $Acl = Get-Acl $VbsPath
        $Rule = New-Object System.Security.AccessControl.FileSystemAccessRule($UsersGroup, "ReadAndExecute", "Allow")
        $Acl.SetAccessRule($Rule)
        Set-Acl -Path $VbsPath -AclObject $Acl -ErrorAction SilentlyContinue
    } catch {}

    return $VbsPath
}

function Get-DatosEquipo {
    $Bios = Get-WmiObject Win32_BIOS
    $MACEth = (Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Ethernet" -and $_.MACAddress } | Select-Object -First 1).MACAddress -replace ":", ""
    $MACWifi = (Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Wi-Fi" -and $_.MACAddress } | Select-Object -First 1).MACAddress -replace ":", ""
    $ProductKey = ""
    try { $ProductKey = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" | Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey } catch {}
    return @{ Serie = $Bios.SerialNumber; MACEthernet = $MACEth; MACWiFi = $MACWifi; ProductKey = $ProductKey }
}

# --- FUNCIONES DE GOOGLE SHEETS ---

function Send-DatosInicio {
    param([string]$InvST, [hashtable]$Datos)

    Write-StepHeader -Step 1 -Title "REGISTRANDO INICIO EN GOOGLE SHEETS"
    Show-ProgressCosmos -Step 1
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $FechaHoy = Get-Date -Format "dd/MM/yyyy"
    $FechaGarantia = (Get-Date).AddYears(3).ToString("dd/MM/yyyy")

    $CPU = (Get-WmiObject Win32_Processor | Select-Object -First 1).Name -replace '\(R\)|\(TM\)|CPU|@.*', '' -replace '\s+', ' '
    $Nucleos = (Get-WmiObject Win32_Processor | Select-Object -First 1).NumberOfCores
    $RAMTotal = [math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 0)
    $Disco = Get-WmiObject Win32_DiskDrive | Select-Object -First 1
    $DiscoGB = [math]::Round($Disco.Size / 1GB, 0)
    $DiscoTipo = if ($Disco.Model -like "*SSD*" -or $Disco.Model -like "*NVMe*") { "SSD" } else { "HDD" }

    $Body = @{
        Accion = "crear"
        Fecha = $FechaHoy
        InvST = $InvST
        Serie = $Datos.Serie
        Marca = "Lenovo"
        Modelo = "ThinkCentre M70s Gen 5"
        Procesador = $CPU.Trim()
        Nucleos = $Nucleos
        RAM = $RAMTotal
        Disco = $DiscoGB
        DiscoTipo = $DiscoTipo
        MACEthernet = $Datos.MACEthernet
        MACWiFi = $Datos.MACWiFi
        ProductKey = $Datos.ProductKey
        FechaFab = $FechaHoy
        Garantia = $FechaGarantia
    } | ConvertTo-Json

    try {
        Write-Host "  Enviando datos del equipo..." -ForegroundColor Cyan
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 60

        if ($Response.status -eq "OK") {
            Write-Host ""
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |       EQUIPO REGISTRADO EN SHEETS        |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  Inventario:  $InvST" -ForegroundColor White
            Write-Host "  |  No. Serie:   $($Datos.Serie)" -ForegroundColor White
            Write-Host "  |  FILA:        $($Response.row)" -ForegroundColor Yellow
            if ($Response.faa -and $Response.faa -ne "") {
                if ($Response.faa -eq "NO ENCONTRADO") {
                    Write-Host "  |  FAA:         $($Response.faa)" -ForegroundColor Red
                } else {
                    Write-Host "  |  FAA:         $($Response.faa)" -ForegroundColor Green
                }
            }
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host ""

            # Verificar resultado FAA - detectar equipo OPD
            if ($Response.faa -eq "NO ENCONTRADO") {
                # Sonido de alerta: 3 beeps
                [Console]::Beep(800, 300); Start-Sleep -Milliseconds 100
                [Console]::Beep(800, 300); Start-Sleep -Milliseconds 100
                [Console]::Beep(800, 300)

                Write-Host ""
                Write-Host "  $([char]0x2716)$([char]0x2716)$([char]0x2716)  ALERTA: EQUIPO NO ENCONTRADO EN LISTA FAA  $([char]0x2716)$([char]0x2716)$([char]0x2716)" -ForegroundColor Red
                Write-Host ""
                Write-Host "  El numero de serie '$($Datos.Serie)' NO aparece en la lista oficial." -ForegroundColor Yellow
                Write-Host "  Este equipo sera configurado como OPD." -ForegroundColor Yellow
                Write-Host ""
                $Respuesta = Read-Host "  Continuar como equipo OPD? (S/N)"
                if ($Respuesta -ne "S" -and $Respuesta -ne "s") {
                    Write-Host "  Configuracion cancelada." -ForegroundColor Red
                    exit
                }
                $Script:Departamento = "OPD"
                $Script:EsOPD = $true
                Write-Host "  Configurando como equipo OPD..." -ForegroundColor Yellow
                Write-Host ""
            }

            $Script:FilaRegistro = $Response.row
            return $true
        }
    } catch {
        Write-Host "  [ERROR] No se pudo registrar: $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

function Send-DatosFin {
    param([string]$InvST)

    Write-StepHeader -Step 26 -Title "ACTUALIZANDO GOOGLE SHEETS - COMPLETADO"
    Show-ProgressCosmos -Step 26
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $SoftwareList = if ($Script:SoftwareInstalado.Count -gt 0) { $Script:SoftwareInstalado -join ", " } else { "Configuracion completa" }

    $Body = @{
        Accion = "actualizar"
        InvST = $InvST
        SoftwareInstalado = $SoftwareList
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 60
        if ($Response.status -eq "OK") {
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |     EQUIPO MARCADO COMO ACTIVO           |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            return $true
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

function Get-SoftwareInfo {
    $OS = Get-WmiObject Win32_OperatingSystem
    $WindowsVersion = $OS.Caption -replace "Microsoft ", ""
    $WindowsBuild = $OS.BuildNumber
    $LicenseStatus = (Get-WmiObject -Query "SELECT LicenseStatus FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" | Select-Object -First 1).LicenseStatus
    $WindowsActivado = if ($LicenseStatus -eq 1) { "Si" } else { "No" }
    $ProductKey = ""
    try { $ProductKey = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" | Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey } catch {}

    $OfficeVersion = "No"
    if (Test-Path "C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE") { $OfficeVersion = "2007" }
    elseif (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") { $OfficeVersion = "365/2016+" }

    $ChromeVersion = "No"
    if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
        try { $ChromeVersion = (Get-Item "C:\Program Files\Google\Chrome\Application\chrome.exe").VersionInfo.FileVersion } catch { $ChromeVersion = "Si" }
    }

    $AcrobatVersion = "No"
    $AcrobatPaths = @("C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe", "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe")
    foreach ($path in $AcrobatPaths) { if (Test-Path $path) { $AcrobatVersion = "Si"; break } }

    $DotNet35 = if ((Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State -eq "Enabled") { "Si" } else { "No" }
    $DedalusInstalado = if (Test-Path "C:\Dedalus") { "Si" } else { "No" }
    $ESETInstalado = if (Get-Service -Name "ekrn" -ErrorAction SilentlyContinue) { "Si" } else { "No" }
    $WinRARInstalado = if (Test-Path "C:\Program Files\WinRAR\WinRAR.exe") { "Si" } else { "No" }

    return @{
        WindowsVersion = $WindowsVersion; WindowsBuild = $WindowsBuild; WindowsActivado = $WindowsActivado
        ProductKey = $ProductKey; NombreEquipo = $env:COMPUTERNAME; UsuarioWindows = $env:USERNAME
        Office = $OfficeVersion; Chrome = $ChromeVersion; Acrobat = $AcrobatVersion
        DotNet35 = $DotNet35; Dedalus = $DedalusInstalado; ESET = $ESETInstalado; WinRAR = $WinRARInstalado
    }
}

function Send-SoftwareInfo {
    param([string]$InvST)

    Write-StepHeader -Step 27 -Title "REGISTRANDO INVENTARIO DE SOFTWARE"
    Show-ProgressCosmos -Step 27
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $Info = Get-SoftwareInfo
    # Usar el nombre nuevo (PC-InvST) ya que Rename-Computer no actualiza $env:COMPUTERNAME hasta reiniciar
    $NombreReal = "PC-$InvST"
    $Body = @{
        Accion = "software"; InvST = $InvST; NombreEquipo = $NombreReal
        WindowsVersion = $Info.WindowsVersion; WindowsBuild = $Info.WindowsBuild
        WindowsActivado = $Info.WindowsActivado; ProductKey = $Info.ProductKey
        Office = $Info.Office; Chrome = $Info.Chrome; Acrobat = $Info.Acrobat
        DotNet35 = $Info.DotNet35; Dedalus = $Info.Dedalus; ESET = $Info.ESET
        WinRAR = $Info.WinRAR; UsuarioWindows = $Info.UsuarioWindows
        FechaConfig = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json

    try {
        $Response = Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 60
        if ($Response.status -eq "OK") {
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |   INVENTARIO SOFTWARE REGISTRADO         |" -ForegroundColor Green
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  Windows:  $($Info.WindowsVersion) (Build $($Info.WindowsBuild))" -ForegroundColor White
            Write-Host "  |  Activado: $($Info.WindowsActivado)" -ForegroundColor $(if ($Info.WindowsActivado -eq "Si") { "Green" } else { "Red" })
            Write-Host "  |  Office:   $($Info.Office)" -ForegroundColor White
            Write-Host "  |  Chrome:   $($Info.Chrome)" -ForegroundColor White
            Write-Host "  |  ESET:     $($Info.ESET)" -ForegroundColor $(if ($Info.ESET -eq "Si") { "Green" } else { "Red" })
            Write-Host "  |  WinRAR:   $($Info.WinRAR)" -ForegroundColor White
            Write-Host "  |  Dedalus:  $($Info.Dedalus)" -ForegroundColor White
            Write-Host "  +------------------------------------------+" -ForegroundColor Green
            return $true
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    return $false
}

# --- FUNCIONES DE CONFIGURACION ---

function Connect-Servidor {
    Write-StepHeader -Step 2 -Title "CONECTANDO AL SERVIDOR"
    Show-ProgressCosmos -Step 2
    net use \\$Servidor /delete /y 2>$null | Out-Null
    $Result = net use \\$Servidor\soportefaa /user:$Usuario $Password /persistent:no 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Log "Conectado a \\$Servidor\soportefaa" "OK"
        return $true
    } else {
        Write-Log "Error al conectar: $Result" "ERROR"
        return $false
    }
}

function Remove-OfficePrevio {
    Write-StepHeader -Step 3 -Title "ELIMINANDO VERSIONES PREVIAS DE OFFICE"
    Show-ProgressCosmos -Step 3

    $Eliminado = $false

    # --- Check idempotencia: si no hay CTR ni Office 365/OneNote, no hay nada que hacer ---
    $CTR = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
    if (-not (Test-Path $CTR)) {
        $RegPaths_Check = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
        $Office365Check = @()
        foreach ($rp in $RegPaths_Check) {
            Get-ItemProperty $rp -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -and ($_.DisplayName -like "*Microsoft 365*" -or $_.DisplayName -like "*Office 365*" -or $_.DisplayName -like "*OneNote*es-es*")
            } | ForEach-Object { $Office365Check += $_ }
        }
        if (-not $Office365Check) {
            Write-Log "No hay versiones previas de Office que remover" "INFO"
            return
        }
    }

    # Iniciar melodia ambiental durante la desinstalacion
    Start-BackgroundMelody -Mensaje "Melodia del Santuario - Desinstalando Office..."

    # --- 1. Desinstalar Office Click-to-Run (Microsoft 365, OneNote, etc.) ---
    if (Test-Path $CTR) {
        $Config = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
        if ($Config -and $Config.ProductReleaseIds) {
            $ProductIds = $Config.ProductReleaseIds
            Write-Log "Office Click-to-Run detectado: $ProductIds"

            # Desinstalar cada producto por separado
            $Productos = $ProductIds -split ","
            foreach ($prod in $Productos) {
                $prod = $prod.Trim()
                if ($prod) {
                    Write-Log "Removiendo: $prod"
                    Start-Process -FilePath $CTR -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=$prod DisplayLevel=False" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                }
            }

            # Esperar a que termine el proceso de desinstalacion
            $Intentos = 0
            while ((Get-Process -Name "OfficeClickToRun" -ErrorAction SilentlyContinue) -and $Intentos -lt 30) {
                Start-Sleep -Seconds 2
                $Intentos++
            }

            Write-Log "Office Click-to-Run desinstalado" "OK"
            $Eliminado = $true
        }
    }

    # --- 2. Desinstalar por registro (captura cualquier version restante) ---
    $RegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -and (
                $_.DisplayName -like "*Microsoft 365*" -or
                $_.DisplayName -like "*Office 365*" -or
                $_.DisplayName -like "*OneNote*es-es*"
            )
        } | ForEach-Object {
            $Nombre = $_.DisplayName
            $Uninstall = $_.UninstallString
            if ($Uninstall) {
                Write-Log "Desinstalando: $Nombre"
                try {
                    # Usar cmd /c para ejecutar la cadena de desinstalacion tal cual
                    Start-Process "cmd.exe" -ArgumentList "/c $Uninstall DisplayLevel=False" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    Write-Log "Desinstalado: $Nombre" "OK"
                    $Eliminado = $true
                } catch {
                    Write-Log "Error al desinstalar $Nombre : $($_.Exception.Message)" "WARN"
                }
            }
        }
    }

    # --- 3. Verificar que se elimino (ambos paths de registro) ---
    Start-Sleep -Seconds 5
    $Restante = @()
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -and (
                $_.DisplayName -like "*Microsoft 365*" -or
                $_.DisplayName -like "*Office 365*" -or
                $_.DisplayName -like "*OneNote*es-es*"
            )
        } | ForEach-Object { $Restante += $_ }
    }

    if ($Restante) {
        Write-Log "Aun quedan restos de Office, intentando limpieza final..." "WARN"
        foreach ($r in $Restante) {
            if ($r.UninstallString) {
                Start-Process "cmd.exe" -ArgumentList "/c $($r.UninstallString) DisplayLevel=False" -Wait -NoNewWindow -ErrorAction SilentlyContinue
            }
        }
    }

    # Detener melodia ambiental
    Stop-BackgroundMelody

    if ($Eliminado) {
        $Script:SoftwareInstalado += "Office previo removido"
    } else {
        Write-Log "No se encontraron versiones previas de Office" "INFO"
    }
}

function New-UsuarioSoporte {
    Write-StepHeader -Step 4 -Title "CREANDO USUARIO SOPORTE (ADMINISTRADOR)"
    Show-ProgressCosmos -Step 4

    $Existe = Get-LocalUser -Name $UsuarioSoporte -ErrorAction SilentlyContinue
    if (-not $Existe) {
        $SecurePass = ConvertTo-SecureString $PasswordSoporte -AsPlainText -Force
        New-LocalUser -Name $UsuarioSoporte -Password $SecurePass -Description "Soporte Tecnico HCG" -PasswordNeverExpires -UserMayNotChangePassword | Out-Null
        Write-Log "Usuario '$UsuarioSoporte' creado" "OK"
    } else {
        Write-Log "Usuario '$UsuarioSoporte' ya existe" "INFO"
    }

    # Asegurar que es administrador
    Add-LocalGroupMember -Group $GrupoAdmin -Member $UsuarioSoporte -ErrorAction SilentlyContinue
    Write-Log "Usuario '$UsuarioSoporte' es Administrador" "OK"
    $Script:SoftwareInstalado += "Usuario Soporte"
}

function New-UsuarioEquipo {
    param([string]$NumInventario)

    $NombreUsuario = if ($Script:EsOPD) { "OPD" } else { $NumInventario }
    $DescUsuario = if ($Script:EsOPD) { "Usuario OPD - HCG" } else { "Usuario Equipo $NumInventario - HCG FAA" }

    Write-StepHeader -Step 5 -Title "CREANDO USUARIO NORMAL ($NombreUsuario)"
    Show-ProgressCosmos -Step 5

    $Existe = Get-LocalUser -Name $NombreUsuario -ErrorAction SilentlyContinue

    if (-not $Existe) {
        New-LocalUser -Name $NombreUsuario -NoPassword -Description $DescUsuario -PasswordNeverExpires -UserMayNotChangePassword | Out-Null
        Write-Log "Usuario '$NombreUsuario' creado" "OK"
    } else {
        Write-Log "Usuario '$NombreUsuario' ya existe" "INFO"
    }

    # Asegurar que es usuario estandar (grupo Usuarios)
    Add-LocalGroupMember -Group $GrupoUsuarios -Member $NombreUsuario -ErrorAction SilentlyContinue

    # Asegurar que NO es administrador
    Remove-LocalGroupMember -Group $GrupoAdmin -Member $NombreUsuario -ErrorAction SilentlyContinue

    # Habilitar el usuario
    Enable-LocalUser -Name $NombreUsuario -ErrorAction SilentlyContinue

    # Configurar auto-login para que este usuario inicie sesion automaticamente
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Set-ItemProperty -Path $RegPath -Name "AutoAdminLogon" -Value "1"
    Set-ItemProperty -Path $RegPath -Name "DefaultUserName" -Value $NombreUsuario
    Set-ItemProperty -Path $RegPath -Name "DefaultPassword" -Value ""
    Set-ItemProperty -Path $RegPath -Name "DefaultDomainName" -Value $env:COMPUTERNAME
    Write-Log "Auto-login configurado para '$NombreUsuario'" "OK"

    # Crear script de primer inicio para configurar barra de tareas del nuevo usuario
    $FirstLoginScript = @'
# HCG - Configurar barra de tareas en primer inicio de sesion
Start-Sleep -Seconds 15
$TaskBarFolder = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
if (-not (Test-Path $TaskBarFolder)) { New-Item -ItemType Directory -Path $TaskBarFolder -Force | Out-Null }
$ChromeShortcut = "C:\Users\Public\Desktop\Google Chrome.lnk"
if (Test-Path $ChromeShortcut) { Copy-Item -Path $ChromeShortcut -Destination $TaskBarFolder -Force }
$xHISShortcut = "C:\Dedalus\xHIS\xHIS v6.lnk"
if (Test-Path $xHISShortcut) { Copy-Item -Path $xHISShortcut -Destination $TaskBarFolder -Force }
foreach ($folder in @("C:\Dedalus\EscritorioClinico", "C:\Dedalus\xFARMA", "C:\Dedalus\hPRESC")) {
    if (Test-Path $folder) {
        $lnk = Get-ChildItem -Path $folder -Filter "*.lnk" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($lnk) { Copy-Item -Path $lnk.FullName -Destination $TaskBarFolder -Force }
    }
}
'@
    Initialize-HCGLogsFolder | Out-Null
    $FirstLoginScript | Out-File -FilePath "C:\HCG_Logs\setup_firstlogin.ps1" -Encoding UTF8 -Force

    # Configurar perfil por defecto para heredar configuracion al nuevo usuario
    try {
        reg load "HKU\DefaultUser" "C:\Users\Default\NTUSER.DAT" 2>$null
        $DefRunOnce = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
        if (-not (Test-Path $DefRunOnce)) { New-Item -Path $DefRunOnce -Force | Out-Null }
        Set-ItemProperty -Path $DefRunOnce -Name "HCG_Setup" -Value 'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\HCG_Logs\setup_firstlogin.ps1"'
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "Script de primer inicio configurado" "OK"
    } catch {
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "No se pudo configurar script de primer inicio" "WARN"
    }

    Write-Log "Usuario '$NombreUsuario' listo (estandar, auto-login)" "OK"
    $Script:SoftwareInstalado += "Usuario $NombreUsuario"
}

function Set-ImagenesUsuarios {
    param([string]$NumInventario)

    Write-StepHeader -Step 6 -Title "CONFIGURANDO IMAGENES DE PERFIL"
    Show-ProgressCosmos -Step 6

    try {
        Add-Type -AssemblyName System.Drawing

        $Size = 448

        # =================================================================
        # PALETA DE COLORES - Tonos suaves, relajantes y profesionales
        # Cada equipo recibe una combinacion unica basada en su inventario
        # =================================================================
        $PaletaColores = @(
            @{ Fondo1 = @(100, 149, 237); Fondo2 = @(65, 105, 180);  Nombre = "Cielo" },
            @{ Fondo1 = @(72, 191, 170);  Fondo2 = @(45, 150, 140);  Nombre = "Jade" },
            @{ Fondo1 = @(180, 136, 200); Fondo2 = @(140, 100, 170); Nombre = "Lavanda" },
            @{ Fondo1 = @(210, 150, 100); Fondo2 = @(175, 120, 75);  Nombre = "Ambar" },
            @{ Fondo1 = @(95, 170, 200);  Fondo2 = @(60, 130, 165);  Nombre = "Oceano" },
            @{ Fondo1 = @(120, 190, 130); Fondo2 = @(80, 155, 95);   Nombre = "Esmeralda" },
            @{ Fondo1 = @(200, 130, 150); Fondo2 = @(165, 95, 120);  Nombre = "Coral" },
            @{ Fondo1 = @(150, 160, 200); Fondo2 = @(110, 120, 170); Nombre = "Lila" },
            @{ Fondo1 = @(200, 175, 110); Fondo2 = @(170, 145, 80);  Nombre = "Miel" },
            @{ Fondo1 = @(110, 180, 190); Fondo2 = @(75, 145, 160);  Nombre = "Turquesa" },
            @{ Fondo1 = @(175, 140, 170); Fondo2 = @(140, 105, 140); Nombre = "Amatista" },
            @{ Fondo1 = @(140, 185, 140); Fondo2 = @(100, 150, 105); Nombre = "Salvia" },
            @{ Fondo1 = @(190, 155, 130); Fondo2 = @(155, 120, 95);  Nombre = "Arena" },
            @{ Fondo1 = @(130, 165, 210); Fondo2 = @(90, 130, 180);  Nombre = "Zafiro" },
            @{ Fondo1 = @(185, 160, 185); Fondo2 = @(150, 125, 150); Nombre = "Orquidea" },
            @{ Fondo1 = @(160, 195, 165); Fondo2 = @(120, 160, 130); Nombre = "Menta" }
        )

        # Seleccionar color basado en el numero de inventario (determinista)
        $Seed = [int]($NumInventario) % $PaletaColores.Count
        $Colores = $PaletaColores[$Seed]

        # =================================================================
        # AVATAR SOPORTE: Elegante, profesional con "S"
        # =================================================================
        $BmpS = New-Object System.Drawing.Bitmap $Size, $Size
        $GfxS = [System.Drawing.Graphics]::FromImage($BmpS)
        $GfxS.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $GfxS.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        # Fondo gradiente gris oscuro elegante
        $RectS = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)
        $GradS = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $RectS,
            [System.Drawing.Color]::FromArgb(55, 60, 70),
            [System.Drawing.Color]::FromArgb(35, 40, 50),
            [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal
        )
        $GfxS.FillRectangle($GradS, $RectS)

        # Circulo principal gris azulado
        $BrushCircle = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(75, 85, 100))
        $GfxS.FillEllipse($BrushCircle, 30, 30, ($Size - 60), ($Size - 60))

        # Anillo sutil
        $PenRing = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(100, 200, 200, 210)), 3
        $GfxS.DrawEllipse($PenRing, 30, 30, ($Size - 60), ($Size - 60))

        # Letra S
        $SFS = New-Object System.Drawing.StringFormat
        $SFS.Alignment = [System.Drawing.StringAlignment]::Center
        $SFS.LineAlignment = [System.Drawing.StringAlignment]::Center
        $FontS = New-Object System.Drawing.Font("Segoe UI Light", 140)
        $BrushText = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 225, 230))
        $GfxS.DrawString("S", $FontS, $BrushText, (New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)), $SFS)

        $AvatarSoporte = "$env:TEMP\avatar_soporte.jpg"
        $BmpS.Save($AvatarSoporte, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $FontS.Dispose(); $BrushText.Dispose(); $BrushCircle.Dispose()
        $PenRing.Dispose(); $GradS.Dispose(); $SFS.Dispose()
        $GfxS.Dispose(); $BmpS.Dispose()

        # =================================================================
        # AVATAR USUARIO: Minimalista, suave y unico por equipo
        # =================================================================
        $AvatarUsuario = "$env:TEMP\avatar_usuario.jpg"

        $BmpU = New-Object System.Drawing.Bitmap $Size, $Size
        $GfxU = [System.Drawing.Graphics]::FromImage($BmpU)
        $GfxU.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $GfxU.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        # Fondo gradiente suave con los colores del equipo
        $RectU = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)
        $Color1 = [System.Drawing.Color]::FromArgb($Colores.Fondo1[0], $Colores.Fondo1[1], $Colores.Fondo1[2])
        $Color2 = [System.Drawing.Color]::FromArgb($Colores.Fondo2[0], $Colores.Fondo2[1], $Colores.Fondo2[2])

        # Angulo del gradiente segun inventario
        $AnguloGrad = @(
            [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal,
            [System.Drawing.Drawing2D.LinearGradientMode]::BackwardDiagonal,
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical,
            [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal
        )
        $GradMode = $AnguloGrad[[int]($NumInventario) % $AnguloGrad.Count]
        $GradU = New-Object System.Drawing.Drawing2D.LinearGradientBrush($RectU, $Color1, $Color2, $GradMode)
        $GfxU.FillRectangle($GradU, $RectU)

        # Efecto de luz suave (circulo semi-transparente arriba-izquierda)
        $BrushGlow = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(25, 255, 255, 255))
        $GfxU.FillEllipse($BrushGlow, -100, -120, ($Size + 80), ($Size))

        # Anillo circular exterior sutil
        $PenMarco = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(60, 255, 255, 255)), 3
        $GfxU.DrawEllipse($PenMarco, 20, 20, ($Size - 40), ($Size - 40))

        # Numero de inventario centrado
        $SFU = New-Object System.Drawing.StringFormat
        $SFU.Alignment = [System.Drawing.StringAlignment]::Center
        $SFU.LineAlignment = [System.Drawing.StringAlignment]::Center

        $FontNum = New-Object System.Drawing.Font("Segoe UI Light", 90)
        # Sombra suave
        $BrushSombra = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(50, 0, 0, 0))
        $GfxU.DrawString($NumInventario, $FontNum, $BrushSombra, (New-Object System.Drawing.RectangleF(3, 3, $Size, $Size)), $SFU)
        # Texto blanco
        $BrushBlanco = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(240, 255, 255, 255))
        $GfxU.DrawString($NumInventario, $FontNum, $BrushBlanco, (New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)), $SFU)

        # Texto "HCG" abajo, pequeno y discreto
        $FontHCG = New-Object System.Drawing.Font("Segoe UI Light", 22)
        $BrushHCG = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(140, 255, 255, 255))
        $GfxU.DrawString("HCG", $FontHCG, $BrushHCG, (New-Object System.Drawing.RectangleF(0, ($Size - 100), $Size, 70)), $SFU)

        # Guardar
        $BmpU.Save($AvatarUsuario, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        Write-Log "Avatar personalizado: tono $($Colores.Nombre)" "OK"

        # Limpiar recursos
        $FontNum.Dispose(); $FontHCG.Dispose()
        $BrushBlanco.Dispose(); $BrushSombra.Dispose(); $BrushHCG.Dispose()
        $BrushGlow.Dispose(); $PenMarco.Dispose(); $GradU.Dispose(); $SFU.Dispose()
        $GfxU.Dispose(); $BmpU.Dispose()

        # =================================================================
        # APLICAR AVATARES CON API DE WINDOWS
        # =================================================================
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class UserTileAPI {
    [DllImport("shell32.dll", EntryPoint = "#262", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern void SetUserTile(string username, int reserved, string imagePath);
}
"@ -ErrorAction SilentlyContinue

        [UserTileAPI]::SetUserTile($UsuarioSoporte, 0, $AvatarSoporte)
        Write-Log "Imagen de perfil: '$UsuarioSoporte' configurada (Panel de Control)" "OK"

        [UserTileAPI]::SetUserTile($NumInventario, 0, $AvatarUsuario)
        Write-Log "Imagen de perfil: '$NumInventario' -> $($Colores.Nombre) (Panel de Control)" "OK"

        # =================================================================
        # APLICAR AVATARES A WINDOWS SETTINGS (App de Configuracion)
        # Windows Settings lee de AccountPicture en el registro HKLM
        # y no usa la API SetUserTile legacy de shell32.dll
        # =================================================================
        $TamanosAvatar = @(448, 240, 192, 96, 64, 48, 40, 32)

        foreach ($DatosUsuario in @(
            @{ Nombre = $UsuarioSoporte; Fuente = $AvatarSoporte },
            @{ Nombre = $NumInventario; Fuente = $AvatarUsuario }
        )) {
            try {
                $NombreUser = $DatosUsuario.Nombre
                $ImagenFuente = $DatosUsuario.Fuente

                # Obtener SID del usuario
                $SID = (New-Object System.Security.Principal.NTAccount($NombreUser)).Translate(
                    [System.Security.Principal.SecurityIdentifier]
                ).Value

                # Crear carpeta persistente para avatares
                $CarpetaAvatar = "C:\ProgramData\HCG\avatars\$SID"
                if (-not (Test-Path $CarpetaAvatar)) {
                    New-Item -ItemType Directory -Path $CarpetaAvatar -Force | Out-Null
                }

                # Cargar imagen fuente (448x448)
                $ImgOriginal = [System.Drawing.Image]::FromFile($ImagenFuente)

                # Crear clave de registro para AccountPicture
                $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users\$SID"
                if (-not (Test-Path $RegPath)) {
                    New-Item -Path $RegPath -Force | Out-Null
                }

                # Generar imagen en cada tamano requerido por Settings
                foreach ($Tam in $TamanosAvatar) {
                    $ArchivoDestino = "$CarpetaAvatar\Image$Tam.jpg"

                    $BmpResize = New-Object System.Drawing.Bitmap($Tam, $Tam)
                    $GfxResize = [System.Drawing.Graphics]::FromImage($BmpResize)
                    $GfxResize.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                    $GfxResize.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                    $GfxResize.DrawImage($ImgOriginal, 0, 0, $Tam, $Tam)
                    $BmpResize.Save($ArchivoDestino, [System.Drawing.Imaging.ImageFormat]::Jpeg)
                    $GfxResize.Dispose()
                    $BmpResize.Dispose()

                    # Registrar ruta en el registro de Windows
                    Set-ItemProperty -Path $RegPath -Name "Image$Tam" -Value $ArchivoDestino -Type String
                }

                $ImgOriginal.Dispose()
                Write-Log "Avatar Settings ($NombreUser): $($TamanosAvatar.Count) tamanos registrados" "OK"

            } catch {
                Write-Log "Avatar Settings ($($DatosUsuario.Nombre)): $($_.Exception.Message)" "WARN"
            }
        }

    } catch {
        Write-Log "No se pudieron configurar imagenes: $($_.Exception.Message)" "WARN"
    }
}

function Set-RedPrivada {
    Write-StepHeader -Step 7 -Title "CONFIGURANDO RED PRIVADA"
    Show-ProgressCosmos -Step 7
    Get-NetConnectionProfile | ForEach-Object { Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private -ErrorAction SilentlyContinue }
    netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes 2>$null
    netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes 2>$null
    Write-Log "Red configurada como privada" "OK"
    $Script:SoftwareInstalado += "Red privada"
}

function Set-HoraAutomatica {
    Write-StepHeader -Step 8 -Title "CONFIGURANDO HORA AUTOMATICA"
    Show-ProgressCosmos -Step 8
    Set-TimeZone -Id $ZonaHoraria -ErrorAction SilentlyContinue
    Set-Service -Name w32time -StartupType Automatic
    Start-Service w32time -ErrorAction SilentlyContinue
    w32tm /config /manualpeerlist:"time.windows.com" /syncfromflags:manual /reliable:yes /update 2>$null
    w32tm /resync /force 2>$null
    Write-Log "Hora automatica configurada (Guadalajara)" "OK"
    $Script:SoftwareInstalado += "Hora auto"
}

function Set-TemaOscuro {
    param([string]$NumInventario)

    Write-StepHeader -Step 9 -Title "CONFIGURANDO TEMA OSCURO PERSONALIZADO"
    Show-ProgressCosmos -Step 9

    # =================================================================
    # PALETA DE COLORES DE ACENTO (misma que avatares - suave y relajante)
    # Windows usa formato ABGR (invertido) para AccentColor en el registro
    # =================================================================
    $PaletaAcento = @(
        @{ R = 100; G = 149; B = 237; Nombre = "Cielo" },
        @{ R = 72;  G = 191; B = 170; Nombre = "Jade" },
        @{ R = 180; G = 136; B = 200; Nombre = "Lavanda" },
        @{ R = 210; G = 150; B = 100; Nombre = "Ambar" },
        @{ R = 95;  G = 170; B = 200; Nombre = "Oceano" },
        @{ R = 120; G = 190; B = 130; Nombre = "Esmeralda" },
        @{ R = 200; G = 130; B = 150; Nombre = "Coral" },
        @{ R = 150; G = 160; B = 200; Nombre = "Lila" },
        @{ R = 200; G = 175; B = 110; Nombre = "Miel" },
        @{ R = 110; G = 180; B = 190; Nombre = "Turquesa" },
        @{ R = 175; G = 140; B = 170; Nombre = "Amatista" },
        @{ R = 140; G = 185; B = 140; Nombre = "Salvia" },
        @{ R = 190; G = 155; B = 130; Nombre = "Arena" },
        @{ R = 130; G = 165; B = 210; Nombre = "Zafiro" },
        @{ R = 185; G = 160; B = 185; Nombre = "Orquidea" },
        @{ R = 160; G = 195; B = 165; Nombre = "Menta" }
    )

    # Seleccionar color basado en el numero de inventario (mismo que avatar)
    $Seed = [int]($NumInventario) % $PaletaAcento.Count
    $ColorAccent = $PaletaAcento[$Seed]

    # Windows AccentColor usa formato ABGR: 0xFF + BB + GG + RR
    # Usar BitConverter para obtener representacion Int32 correcta (DWord requiere signed int)
    $HexAccent = "0xFF{0:X2}{1:X2}{2:X2}" -f $ColorAccent.B, $ColorAccent.G, $ColorAccent.R
    $AccentABGR = [BitConverter]::ToInt32([BitConverter]::GetBytes([Convert]::ToUInt32($HexAccent, 16)), 0)
    # Version mas oscura para elementos inactivos
    $R2 = [Math]::Max(0, $ColorAccent.R - 30)
    $G2 = [Math]::Max(0, $ColorAccent.G - 30)
    $B2 = [Math]::Max(0, $ColorAccent.B - 30)
    $HexMenu = "0xFF{0:X2}{1:X2}{2:X2}" -f $B2, $G2, $R2
    $AccentMenuABGR = [BitConverter]::ToInt32([BitConverter]::GetBytes([Convert]::ToUInt32($HexMenu, 16)), 0)

    $RegPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    Set-ItemProperty -Path $RegPath -Name "AppsUseLightTheme" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $RegPath -Name "SystemUsesLightTheme" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $RegPath -Name "ColorPrevalence" -Value 1 -Type DWord -ErrorAction SilentlyContinue

    # Color de acento en barra de titulo y bordes
    $DWMPath = "HKCU:\SOFTWARE\Microsoft\Windows\DWM"
    Set-ItemProperty -Path $DWMPath -Name "ColorPrevalence" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $DWMPath -Name "AccentColor" -Value $AccentABGR -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $DWMPath -Name "AccentColorInactive" -Value $AccentMenuABGR -Type DWord -ErrorAction SilentlyContinue

    # Color de acento en Start menu y barra de tareas
    $ExplorerAccent = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent"
    if (-not (Test-Path $ExplorerAccent)) { New-Item -Path $ExplorerAccent -Force | Out-Null }
    Set-ItemProperty -Path $ExplorerAccent -Name "AccentColorMenu" -Value $AccentMenuABGR -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $ExplorerAccent -Name "StartColorMenu" -Value $AccentABGR -Type DWord -ErrorAction SilentlyContinue

    Write-Log "Color de acento: $($ColorAccent.Nombre)" "OK"

    # Aplicar tambien al perfil por defecto (para el usuario de inventario)
    try {
        if (-not (Test-Path "Registry::HKU\DefaultUser")) {
            reg load "HKU\DefaultUser" "C:\Users\Default\NTUSER.DAT" 2>$null
        }
        $DefPath = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        if (-not (Test-Path $DefPath)) { New-Item -Path $DefPath -Force | Out-Null }
        Set-ItemProperty -Path $DefPath -Name "AppsUseLightTheme" -Value 0 -Type DWord
        Set-ItemProperty -Path $DefPath -Name "SystemUsesLightTheme" -Value 0 -Type DWord
        Set-ItemProperty -Path $DefPath -Name "ColorPrevalence" -Value 1 -Type DWord

        $DefDWM = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\DWM"
        if (-not (Test-Path $DefDWM)) { New-Item -Path $DefDWM -Force | Out-Null }
        Set-ItemProperty -Path $DefDWM -Name "ColorPrevalence" -Value 1 -Type DWord
        Set-ItemProperty -Path $DefDWM -Name "AccentColor" -Value $AccentABGR -Type DWord
        Set-ItemProperty -Path $DefDWM -Name "AccentColorInactive" -Value $AccentMenuABGR -Type DWord

        $DefAccent = "Registry::HKU\DefaultUser\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent"
        if (-not (Test-Path $DefAccent)) { New-Item -Path $DefAccent -Force | Out-Null }
        Set-ItemProperty -Path $DefAccent -Name "AccentColorMenu" -Value $AccentMenuABGR -Type DWord
        Set-ItemProperty -Path $DefAccent -Name "StartColorMenu" -Value $AccentABGR -Type DWord

        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "Tema personalizado aplicado al perfil por defecto" "OK"
    } catch {
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\DefaultUser" 2>$null
    }

    Write-Log "Tema oscuro personalizado: $($ColorAccent.Nombre)" "OK"
    $Script:SoftwareInstalado += "Tema personalizado"
}

function Set-FondoPantalla {
    Write-StepHeader -Step 14 -Title "ESTABLECIENDO FONDO DE PANTALLA"
    Show-ProgressCosmos -Step 14

    # Copiar wallpaper_hcg.jpg tal cual (sin modificar) a Imagenes publicas
    $ImagenServidor = "$RutaWallpaper\wallpaper_hcg.jpg"
    $CarpetaDestino = "C:\Users\Public\Pictures"
    if (-not (Test-Path $CarpetaDestino)) { New-Item -ItemType Directory -Path $CarpetaDestino -Force | Out-Null }
    $FondoLocal = "$CarpetaDestino\wallpaper_hcg.jpg"

    if (Test-Path $ImagenServidor) {
        Copy-Item -Path $ImagenServidor -Destination $FondoLocal -Force
        Write-Log "Wallpaper copiado a $FondoLocal" "OK"
    } else {
        Write-Log "No se encontro wallpaper_hcg.jpg en: $RutaWallpaper" "ERROR"
        return
    }

    # Aplicar como fondo de pantalla (usuario actual)
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Value $FondoLocal
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value "10"
    rundll32.exe user32.dll, UpdatePerUserSystemParameters 1, True

    # Aplicar tambien al perfil por defecto (para el usuario de inventario)
    try {
        if (-not (Test-Path "Registry::HKU\DefaultUser")) {
            reg load "HKU\DefaultUser" "C:\Users\Default\NTUSER.DAT" 2>$null
        }
        $DefDesktop = "Registry::HKU\DefaultUser\Control Panel\Desktop"
        if (-not (Test-Path $DefDesktop)) { New-Item -Path $DefDesktop -Force | Out-Null }
        Set-ItemProperty -Path $DefDesktop -Name "Wallpaper" -Value $FondoLocal
        Set-ItemProperty -Path $DefDesktop -Name "WallpaperStyle" -Value "10"
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\DefaultUser" 2>$null
        Write-Log "Fondo aplicado al perfil por defecto" "OK"
    } catch {
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\DefaultUser" 2>$null
    }

    Write-Log "Fondo de pantalla establecido" "OK"
    $Script:SoftwareInstalado += "Fondo HCG"
}

function Set-LockScreenBackground {
    param([string]$NumInventario)

    Write-StepHeader -Step 14.5 -Title "GENERANDO FONDO DE PANTALLA DE BLOQUEO"
    Show-ProgressCosmos -Step 14

    try {
        Add-Type -AssemblyName System.Drawing

        # Detectar resolucion de pantalla (con fallback si falla)
        $Width = 1920
        $Height = 1080
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            $Screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
            if ($Screen.Width -ge 1920) {
                $Width = $Screen.Width
                $Height = $Screen.Height
            }
        } catch {
            # Fallback: intentar via WMI
            try {
                $VideoMode = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop | Select-Object -First 1
                if ($VideoMode.CurrentHorizontalResolution -ge 1920) {
                    $Width = $VideoMode.CurrentHorizontalResolution
                    $Height = $VideoMode.CurrentVerticalResolution
                }
            } catch {
                Write-Log "Usando resolucion por defecto (1920x1080)" "INFO"
            }
        }

        Write-Log "Generando fondo de bloqueo ($Width x $Height)..."

        # Usar inventario como semilla para generacion deterministica
        $Seed = [int]$NumInventario
        $Rnd = New-Object System.Random($Seed)

        # =================================================================
        # PALETAS DE COLORES PROFESIONALES
        # =================================================================
        $Paletas = @(
            @{ Base = @(15, 30, 60);    Accent1 = @(45, 80, 140);   Accent2 = @(80, 130, 180);  Accent3 = @(120, 170, 210); Nombre = "Oceano Profundo" },
            @{ Base = @(25, 25, 35);    Accent1 = @(60, 50, 90);    Accent2 = @(100, 80, 140);  Accent3 = @(150, 120, 180); Nombre = "Anochecer" },
            @{ Base = @(20, 35, 35);    Accent1 = @(40, 80, 80);    Accent2 = @(70, 130, 120);  Accent3 = @(100, 170, 160); Nombre = "Bosque" },
            @{ Base = @(35, 25, 30);    Accent1 = @(80, 50, 60);    Accent2 = @(130, 80, 100);  Accent3 = @(180, 120, 140); Nombre = "Atardecer" },
            @{ Base = @(25, 30, 40);    Accent1 = @(50, 70, 100);   Accent2 = @(90, 120, 150);  Accent3 = @(140, 170, 200); Nombre = "Acero" },
            @{ Base = @(30, 35, 25);    Accent1 = @(70, 90, 50);    Accent2 = @(110, 140, 80);  Accent3 = @(150, 180, 120); Nombre = "Pradera" },
            @{ Base = @(40, 30, 25);    Accent1 = @(100, 70, 50);   Accent2 = @(150, 110, 80);  Accent3 = @(200, 160, 120); Nombre = "Tierra" },
            @{ Base = @(20, 30, 45);    Accent1 = @(40, 70, 110);   Accent2 = @(70, 110, 160);  Accent3 = @(110, 160, 210); Nombre = "Cielo Nocturno" }
        )

        $Paleta = $Paletas[$Seed % $Paletas.Count]
        $Estilo = $Seed % 5  # 5 estilos diferentes

        # Crear bitmap
        $Bmp = New-Object System.Drawing.Bitmap($Width, $Height)
        $Gfx = [System.Drawing.Graphics]::FromImage($Bmp)
        $Gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $Gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $Gfx.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

        # Color base
        $ColorBase = [System.Drawing.Color]::FromArgb($Paleta.Base[0], $Paleta.Base[1], $Paleta.Base[2])
        $Gfx.Clear($ColorBase)

        # =================================================================
        # ESTILO 0: ONDAS SUAVES (tipo macOS)
        # =================================================================
        if ($Estilo -eq 0) {
            for ($layer = 0; $layer -lt 6; $layer++) {
                $ColorIdx = $layer % 3
                $ColorArr = if ($ColorIdx -eq 0) { $Paleta.Accent1 } elseif ($ColorIdx -eq 1) { $Paleta.Accent2 } else { $Paleta.Accent3 }
                $Alpha = 40 + ($layer * 15)
                $Color = [System.Drawing.Color]::FromArgb($Alpha, $ColorArr[0], $ColorArr[1], $ColorArr[2])
                $Brush = New-Object System.Drawing.SolidBrush($Color)

                $Points = New-Object System.Collections.ArrayList
                $BaseY = $Height * (0.3 + ($layer * 0.12))
                $Amplitude = 80 + ($Rnd.Next(60))
                $Frequency = 0.002 + ($Rnd.NextDouble() * 0.003)
                $Phase = $Rnd.NextDouble() * 6.28

                for ($x = 0; $x -le $Width; $x += 20) {
                    $y = $BaseY + [Math]::Sin(($x * $Frequency) + $Phase) * $Amplitude
                    $y += [Math]::Sin(($x * $Frequency * 2.5) + $Phase * 1.5) * ($Amplitude * 0.3)
                    [void]$Points.Add([System.Drawing.PointF]::new($x, $y))
                }
                [void]$Points.Add([System.Drawing.PointF]::new($Width, $Height))
                [void]$Points.Add([System.Drawing.PointF]::new(0, $Height))

                if ($Points.Count -ge 3) {
                    $Gfx.FillPolygon($Brush, $Points.ToArray())
                }
                $Brush.Dispose()
            }
        }

        # =================================================================
        # ESTILO 1: BOKEH / CIRCULOS SUAVES
        # =================================================================
        elseif ($Estilo -eq 1) {
            # Gradiente de fondo
            $GradBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                [System.Drawing.Point]::new(0, 0),
                [System.Drawing.Point]::new($Width, $Height),
                [System.Drawing.Color]::FromArgb($Paleta.Base[0], $Paleta.Base[1], $Paleta.Base[2]),
                [System.Drawing.Color]::FromArgb($Paleta.Accent1[0], $Paleta.Accent1[1], $Paleta.Accent1[2])
            )
            $Gfx.FillRectangle($GradBrush, 0, 0, $Width, $Height)
            $GradBrush.Dispose()

            # Circulos bokeh
            for ($i = 0; $i -lt 35; $i++) {
                $Size = $Rnd.Next(50, 350)
                $X = $Rnd.Next(-100, $Width + 100)
                $Y = $Rnd.Next(-100, $Height + 100)
                $Alpha = $Rnd.Next(8, 35)
                $ColorIdx = $Rnd.Next(3)
                $ColorArr = if ($ColorIdx -eq 0) { $Paleta.Accent1 } elseif ($ColorIdx -eq 1) { $Paleta.Accent2 } else { $Paleta.Accent3 }
                $Color = [System.Drawing.Color]::FromArgb($Alpha, $ColorArr[0], $ColorArr[1], $ColorArr[2])

                $Path = New-Object System.Drawing.Drawing2D.GraphicsPath
                $Path.AddEllipse($X, $Y, $Size, $Size)
                $CircleBrush = New-Object System.Drawing.Drawing2D.PathGradientBrush($Path)
                $CircleBrush.CenterColor = [System.Drawing.Color]::FromArgb($Alpha + 20, $ColorArr[0], $ColorArr[1], $ColorArr[2])
                $CircleBrush.SurroundColors = @([System.Drawing.Color]::FromArgb(0, $ColorArr[0], $ColorArr[1], $ColorArr[2]))
                $Gfx.FillEllipse($CircleBrush, $X, $Y, $Size, $Size)
                $CircleBrush.Dispose()
                $Path.Dispose()
            }
        }

        # =================================================================
        # ESTILO 2: AURORA / GRADIENTES FLUIDOS
        # =================================================================
        elseif ($Estilo -eq 2) {
            for ($layer = 0; $layer -lt 8; $layer++) {
                $ColorIdx = $layer % 3
                $ColorArr = if ($ColorIdx -eq 0) { $Paleta.Accent1 } elseif ($ColorIdx -eq 1) { $Paleta.Accent2 } else { $Paleta.Accent3 }
                $Alpha = 20 + ($Rnd.Next(25))

                $Points = New-Object System.Collections.ArrayList
                $CenterY = $Height * (0.2 + ($Rnd.NextDouble() * 0.6))
                $Spread = 100 + ($Rnd.Next(200))

                for ($x = 0; $x -le $Width; $x += 15) {
                    $noise1 = [Math]::Sin($x * 0.003 + $layer) * 150
                    $noise2 = [Math]::Sin($x * 0.007 + $layer * 2) * 80
                    $y = $CenterY + $noise1 + $noise2
                    [void]$Points.Add([System.Drawing.PointF]::new($x, $y - $Spread))
                }
                for ($x = $Width; $x -ge 0; $x -= 15) {
                    $noise1 = [Math]::Sin($x * 0.003 + $layer) * 150
                    $noise2 = [Math]::Sin($x * 0.007 + $layer * 2) * 80
                    $y = $CenterY + $noise1 + $noise2
                    [void]$Points.Add([System.Drawing.PointF]::new($x, $y + $Spread))
                }

                if ($Points.Count -ge 3) {
                    $Color = [System.Drawing.Color]::FromArgb($Alpha, $ColorArr[0], $ColorArr[1], $ColorArr[2])
                    $Brush = New-Object System.Drawing.SolidBrush($Color)
                    $Gfx.FillPolygon($Brush, $Points.ToArray())
                    $Brush.Dispose()
                }
            }
        }

        # =================================================================
        # ESTILO 3: GEOMETRICO MINIMAL
        # =================================================================
        elseif ($Estilo -eq 3) {
            # Lineas diagonales sutiles
            for ($i = 0; $i -lt 12; $i++) {
                $ColorArr = if ($i % 3 -eq 0) { $Paleta.Accent1 } elseif ($i % 3 -eq 1) { $Paleta.Accent2 } else { $Paleta.Accent3 }
                $Alpha = 15 + ($Rnd.Next(20))
                $Color = [System.Drawing.Color]::FromArgb($Alpha, $ColorArr[0], $ColorArr[1], $ColorArr[2])
                $Pen = New-Object System.Drawing.Pen($Color, (2 + $Rnd.Next(4)))

                $X1 = $Rnd.Next($Width)
                $Y1 = $Rnd.Next($Height)
                $Angle = $Rnd.NextDouble() * 3.14159
                $Length = 400 + $Rnd.Next(800)
                $X2 = $X1 + [Math]::Cos($Angle) * $Length
                $Y2 = $Y1 + [Math]::Sin($Angle) * $Length

                $Gfx.DrawLine($Pen, $X1, $Y1, $X2, $Y2)
                $Pen.Dispose()
            }

            # Formas geometricas
            for ($i = 0; $i -lt 8; $i++) {
                $ColorArr = if ($i % 3 -eq 0) { $Paleta.Accent2 } elseif ($i % 3 -eq 1) { $Paleta.Accent3 } else { $Paleta.Accent1 }
                $Alpha = 25 + ($Rnd.Next(30))
                $Color = [System.Drawing.Color]::FromArgb($Alpha, $ColorArr[0], $ColorArr[1], $ColorArr[2])
                $Brush = New-Object System.Drawing.SolidBrush($Color)

                $X = $Rnd.Next($Width)
                $Y = $Rnd.Next($Height)
                $Size = 100 + $Rnd.Next(300)
                $Shape = $Rnd.Next(3)

                if ($Shape -eq 0) {
                    $Gfx.FillEllipse($Brush, $X, $Y, $Size, $Size)
                } elseif ($Shape -eq 1) {
                    $Gfx.FillRectangle($Brush, $X, $Y, $Size, $Size * 0.6)
                } else {
                    $TriPoints = @(
                        [System.Drawing.PointF]::new($X + $Size/2, $Y),
                        [System.Drawing.PointF]::new($X, $Y + $Size),
                        [System.Drawing.PointF]::new($X + $Size, $Y + $Size)
                    )
                    $Gfx.FillPolygon($Brush, $TriPoints)
                }
                $Brush.Dispose()
            }
        }

        # =================================================================
        # ESTILO 4: MONTANAS / CAPAS
        # =================================================================
        elseif ($Estilo -eq 4) {
            for ($layer = 0; $layer -lt 5; $layer++) {
                $ColorArr = if ($layer -lt 2) { $Paleta.Accent1 } elseif ($layer -lt 4) { $Paleta.Accent2 } else { $Paleta.Accent3 }
                $Darkness = 1.0 - ($layer * 0.15)
                $R = [int]($ColorArr[0] * $Darkness)
                $G = [int]($ColorArr[1] * $Darkness)
                $B = [int]($ColorArr[2] * $Darkness)
                $Color = [System.Drawing.Color]::FromArgb(180, $R, $G, $B)
                $Brush = New-Object System.Drawing.SolidBrush($Color)

                $Points = New-Object System.Collections.ArrayList
                $BaseY = $Height * (0.4 + ($layer * 0.12))

                [void]$Points.Add([System.Drawing.PointF]::new(0, $Height))
                for ($x = 0; $x -le $Width; $x += 30) {
                    $Peak = [Math]::Sin($x * 0.005 + $layer * 2) * (150 - $layer * 20)
                    $Peak += [Math]::Sin($x * 0.012 + $layer) * (80 - $layer * 10)
                    $y = $BaseY - [Math]::Abs($Peak)
                    [void]$Points.Add([System.Drawing.PointF]::new($x, $y))
                }
                [void]$Points.Add([System.Drawing.PointF]::new($Width, $Height))

                if ($Points.Count -ge 3) {
                    $Gfx.FillPolygon($Brush, $Points.ToArray())
                }
                $Brush.Dispose()
            }
        }

        # =================================================================
        # GUARDAR Y APLICAR
        # =================================================================
        $ImagePath = $null
        $ImageSaved = $false

        # Intentar guardar en ubicacion del sistema (requiere admin)
        try {
            $LockScreenPath = "C:\Windows\Web\Wallpaper\HCG"
            if (-not (Test-Path $LockScreenPath)) {
                New-Item -ItemType Directory -Path $LockScreenPath -Force -ErrorAction Stop | Out-Null
            }
            $ImagePath = "$LockScreenPath\LockScreen_$NumInventario.jpg"
            $Bmp.Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            $ImageSaved = $true
            Write-Log "Imagen guardada en: $ImagePath" "OK"
        } catch {
            Write-Log "No se pudo guardar en Windows folder: $($_.Exception.Message)" "WARN"
        }

        # Fallback: guardar en ProgramData si fallo lo anterior
        if (-not $ImageSaved) {
            try {
                $LockScreenPath = "C:\ProgramData\HCG\LockScreen"
                if (-not (Test-Path $LockScreenPath)) {
                    New-Item -ItemType Directory -Path $LockScreenPath -Force -ErrorAction Stop | Out-Null
                }
                $ImagePath = "$LockScreenPath\LockScreen_$NumInventario.jpg"
                $Bmp.Save($ImagePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
                $ImageSaved = $true
                Write-Log "Imagen guardada en ubicacion alternativa: $ImagePath" "OK"
            } catch {
                Write-Log "No se pudo guardar imagen: $($_.Exception.Message)" "ERROR"
            }
        }

        # Liberar recursos GDI+
        try { $Gfx.Dispose() } catch {}
        try { $Bmp.Dispose() } catch {}

        if (-not $ImageSaved -or -not $ImagePath) {
            Write-Log "No se pudo generar el fondo de bloqueo" "ERROR"
            return
        }

        Write-Log "Imagen generada: $($Paleta.Nombre) (Estilo $Estilo)" "OK"

        # Aplicar como fondo de pantalla de bloqueo via registro (requiere admin)
        try {
            $RegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
            if (-not (Test-Path $RegPath)) { New-Item -Path $RegPath -Force -ErrorAction Stop | Out-Null }
            Set-ItemProperty -Path $RegPath -Name "LockScreenImage" -Value $ImagePath -Type String -ErrorAction Stop
            Set-ItemProperty -Path $RegPath -Name "NoChangingLockScreen" -Value 1 -Type DWord -ErrorAction Stop
            Write-Log "Registro Policies configurado" "OK"
        } catch {
            Write-Log "No se pudo configurar registro Policies: $($_.Exception.Message)" "WARN"
        }

        # Tambien configurar via PersonalizationCSP (Windows 10/11)
        try {
            $CSPPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
            if (-not (Test-Path $CSPPath)) { New-Item -Path $CSPPath -Force -ErrorAction Stop | Out-Null }
            Set-ItemProperty -Path $CSPPath -Name "LockScreenImagePath" -Value $ImagePath -Type String -ErrorAction Stop
            Set-ItemProperty -Path $CSPPath -Name "LockScreenImageStatus" -Value 1 -Type DWord -ErrorAction Stop
            Write-Log "Registro PersonalizationCSP configurado" "OK"
        } catch {
            Write-Log "No se pudo configurar registro CSP: $($_.Exception.Message)" "WARN"
        }

        Write-Log "Fondo de bloqueo aplicado: $($Paleta.Nombre)" "OK"
        $Script:SoftwareInstalado += "Lock Screen personalizado"

    } catch {
        # Asegurar limpieza de recursos GDI+ en caso de error
        try { if ($Gfx) { $Gfx.Dispose() } } catch {}
        try { if ($Bmp) { $Bmp.Dispose() } } catch {}
        Write-Log "Error al generar fondo de bloqueo: $($_.Exception.Message)" "WARN"
    }
}

function Install-WinRAR {
    Write-StepHeader -Step 10 -Title "INSTALANDO WINRAR CON LICENCIA"
    Show-ProgressCosmos -Step 10

    # PASO 1: Instalar WinRAR
    $YaInstalado = Test-Path "C:\Program Files\WinRAR\WinRAR.exe"

    if (-not $YaInstalado) {
        # Buscar el instalador (winrar-x64-*.exe)
        $Instalador = "$RutaWinRAR\winrar-x64-713.exe"
        if (-not (Test-Path $Instalador)) {
            # Buscar cualquier instalador winrar-x64
            $Instalador = Get-ChildItem -Path $RutaWinRAR -Filter "winrar-x64*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($Instalador) { $Instalador = $Instalador.FullName }
        }
        if (-not $Instalador -or -not (Test-Path $Instalador)) {
            # Buscar cualquier exe que NO sea la licencia
            $Instalador = Get-ChildItem -Path $RutaWinRAR -Filter "*.exe" -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*License*" } | Select-Object -First 1
            if ($Instalador) { $Instalador = $Instalador.FullName }
        }

        if ($Instalador -and (Test-Path $Instalador)) {
            Write-Log "Instalando WinRAR: $(Split-Path $Instalador -Leaf)"
            Start-Process -FilePath $Instalador -ArgumentList "/S" -Wait -NoNewWindow
            Start-Sleep -Seconds 3

            if (Test-Path "C:\Program Files\WinRAR\WinRAR.exe") {
                Write-Log "WinRAR instalado correctamente" "OK"
            } else {
                Write-Log "No se pudo verificar la instalacion de WinRAR" "WARN"
            }
        } else {
            Write-Log "No se encontro instalador de WinRAR en: $RutaWinRAR" "ERROR"
            return
        }
    } else {
        Write-Log "WinRAR ya esta instalado" "INFO"
    }

    # Verificar que WinRAR esta presente antes de registrar
    if (-not (Test-Path "C:\Program Files\WinRAR\WinRAR.exe")) { return }

    # PASO 2: Aplicar licencia
    $Licencia = "$RutaWinRAR\WinRAR License.exe"
    if (-not (Test-Path $Licencia)) {
        $Licencia = Get-ChildItem -Path $RutaWinRAR -Filter "*License*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($Licencia) { $Licencia = $Licencia.FullName }
    }

    if ($Licencia -and (Test-Path $Licencia)) {
        Write-Log "Aplicando licencia de WinRAR..."

        # Ejecutar licencia SIN esperar (para evitar bloqueo por ventana de confirmacion)
        Start-Process -FilePath $Licencia -ArgumentList "/S" -NoNewWindow

        # Esperar un momento para que se aplique la licencia
        Start-Sleep -Seconds 3

        # Cerrar automaticamente cualquier ventana de WinRAR que aparezca
        # (la ventana de confirmacion de licencia)
        $WinRARWindows = Get-Process -Name "WinRAR*" -ErrorAction SilentlyContinue
        foreach ($proc in $WinRARWindows) {
            try {
                # Enviar tecla Enter para cerrar dialogos de confirmacion
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                Start-Sleep -Milliseconds 500
            } catch { }
        }

        # Tambien cerrar procesos de licencia que puedan quedar abiertos
        Get-Process -Name "*License*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Get-Process -Name "WinRAR" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -ne "" } | ForEach-Object {
            # Si tiene ventana abierta, cerrarla
            $_.CloseMainWindow() | Out-Null
        }

        Start-Sleep -Seconds 1

        # Verificar que la licencia se aplico (buscar rarreg.key en la carpeta de WinRAR)
        if (Test-Path "C:\Program Files\WinRAR\rarreg.key") {
            Write-Log "Licencia de WinRAR aplicada correctamente" "OK"
        } else {
            Write-Log "Licencia ejecutada (verificar activacion manualmente)" "WARN"
        }
    } else {
        Write-Log "No se encontro archivo de licencia de WinRAR" "WARN"
    }

    $Script:SoftwareInstalado += "WinRAR"
}

function Install-DotNet35 {
    Write-StepHeader -Step 11 -Title "INSTALANDO .NET FRAMEWORK 3.5"
    Show-ProgressCosmos -Step 11

    $Estado = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State
    if ($Estado -eq "Enabled") {
        Write-Log ".NET 3.5 ya instalado" "INFO"
        $Script:SoftwareInstalado += ".NET 3.5"
        return
    }

    # Iniciar melodia ambiental durante la instalacion
    Start-BackgroundMelody -Mensaje "Melodia del Santuario - Instalando .NET 3.5..."

    # Intentar instalacion offline desde el servidor (sin descargar de internet)
    if (Test-Path $RutaDotNet) {
        Write-Log "Instalando .NET 3.5 offline desde servidor..."
        try {
            $DismResult = Dism /Online /Enable-Feature /FeatureName:NetFx3 /All /Source:"$RutaDotNet" /LimitAccess /NoRestart 2>&1
            $EstadoPost = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State
            if ($EstadoPost -eq "Enabled") {
                Stop-BackgroundMelody
                Write-Log ".NET 3.5 instalado (offline)" "OK"
                $Script:SoftwareInstalado += ".NET 3.5"
                return
            } else {
                Write-Log "DISM offline no completo la instalacion, intentando via Windows Update..." "WARN"
            }
        } catch {
            Write-Log "Error DISM offline: $($_.Exception.Message). Intentando via Windows Update..." "WARN"
        }
    } else {
        Write-Log "Carpeta offline no encontrada ($RutaDotNet). Intentando via Windows Update..." "WARN"
    }

    # Fallback: instalar via Windows Update (requiere internet)
    Write-Log "Instalando .NET 3.5 via Windows Update..."
    Enable-WindowsOptionalFeature -Online -FeatureName "NetFx3" -All -NoRestart -ErrorAction SilentlyContinue | Out-Null

    $EstadoFinal = (Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue).State

    # Detener melodia ambiental
    Stop-BackgroundMelody

    if ($EstadoFinal -eq "Enabled") {
        Write-Log ".NET 3.5 instalado (online)" "OK"
        $Script:SoftwareInstalado += ".NET 3.5"
    } else {
        Write-Log "No se pudo instalar .NET 3.5" "ERROR"
    }
}

function Install-AcrobatReader {
    Write-StepHeader -Step 12 -Title "INSTALANDO ACROBAT READER"
    Show-ProgressCosmos -Step 12

    # Verificar si ya esta instalado
    $AcrobatPaths = @(
        "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    )
    foreach ($path in $AcrobatPaths) {
        if (Test-Path $path) {
            Write-Log "Acrobat Reader ya esta instalado" "INFO"
            $Script:SoftwareInstalado += "Acrobat"
            return
        }
    }

    # Buscar instalador
    $Instalador = "$RutaAcrobat\Reader_es_install.exe"
    if (-not (Test-Path $Instalador)) {
        $Instalador = Find-Installer -Carpeta $RutaAcrobat -Filtro "*.exe"
    }

    if ($Instalador -and (Test-Path $Instalador)) {
        Write-Log "Instalando Acrobat Reader (sin McAfee)..."
        Start-Process -FilePath $Instalador -ArgumentList "/sAll /rs /msi EULA_ACCEPT=YES DISABLE_OPTIONAL_OFFER=YES SUPPRESS_APP_LAUNCH=YES" -Wait -NoNewWindow

        # Verificar instalacion
        $Instalado = $false
        foreach ($path in $AcrobatPaths) {
            if (Test-Path $path) { $Instalado = $true; break }
        }

        if ($Instalado) {
            Write-Log "Acrobat Reader instalado correctamente" "OK"
        } else {
            Write-Log "Acrobat Reader puede estar instalandose en segundo plano" "WARN"
        }

        # Desinstalar McAfee si se colo (usar registro en vez de Win32_Product que es muy lento)
        $McAfeeEntries = @()
        $RegPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
        foreach ($rp in $RegPaths) {
            Get-ItemProperty $rp -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*McAfee*" } | ForEach-Object { $McAfeeEntries += $_ }
        }
        if ($McAfeeEntries.Count -gt 0) {
            Write-Log "Detectado McAfee, desinstalando..." "INFO"
            foreach ($entry in $McAfeeEntries) {
                if ($entry.UninstallString) {
                    Start-Process "cmd.exe" -ArgumentList "/c $($entry.UninstallString) /quiet /norestart" -Wait -NoNewWindow -ErrorAction SilentlyContinue
                }
            }
            Write-Log "McAfee eliminado" "OK"
        }

        $Script:SoftwareInstalado += "Acrobat"
    } else {
        Write-Log "No se encontro instalador de Acrobat en: $RutaAcrobat" "ERROR"
    }
}

function Install-Chrome {
    Write-StepHeader -Step 13 -Title "INSTALANDO GOOGLE CHROME"
    Show-ProgressCosmos -Step 13

    $ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

    if (-not (Test-Path $ChromePath)) {
        # Buscar instalador
        $Instalador = "$RutaChrome\ChromeSetup.exe"
        if (-not (Test-Path $Instalador)) {
            $Instalador = Find-Installer -Carpeta $RutaChrome -Filtro "*.exe"
        }

        if ($Instalador -and (Test-Path $Instalador)) {
            Write-Log "Instalando Chrome..."
            Start-Process -FilePath $Instalador -ArgumentList "/silent /install" -Wait -NoNewWindow
            Start-Sleep -Seconds 5
            Write-Log "Chrome instalado" "OK"
        } else {
            Write-Log "No se encontro instalador de Chrome" "ERROR"
            return
        }
    } else {
        Write-Log "Chrome ya esta instalado" "INFO"
    }

    # Crear acceso directo en escritorio
    $Desktop = "C:\Users\Public\Desktop"
    $ShortcutPath = "$Desktop\Google Chrome.lnk"
    if (-not (Test-Path $ShortcutPath)) {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $ChromePath
        $Shortcut.WorkingDirectory = "C:\Program Files\Google\Chrome\Application"
        $Shortcut.IconLocation = "$ChromePath,0"
        $Shortcut.Save()
        Write-Log "Acceso directo de Chrome en escritorio" "OK"
    }

    # Anclar Chrome a la barra de tareas
    Pin-ToTaskbar -ShortcutPath $ShortcutPath -AppName "Google Chrome"

    # Establecer Chrome como navegador predeterminado
    Write-Log "Configurando Chrome como navegador predeterminado..."

    # Metodo 1: Usar el argumento de Chrome
    Start-Process -FilePath "$ChromePath" -ArgumentList "--make-default-browser" -NoNewWindow -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # Metodo 2: Configurar asociaciones en registro
    $ChromeProgId = "ChromeHTML"
    $Extensions = @(".htm", ".html", ".shtml", ".xht", ".xhtml")
    $Protocols = @("http", "https")

    foreach ($ext in $Extensions) {
        $RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\UserChoice"
        try {
            if (Test-Path $RegPath) { Remove-Item -Path $RegPath -Force -ErrorAction SilentlyContinue }
        } catch {}
    }

    # Configurar como navegador predeterminado via Settings
    Start-Process "ms-settings:defaultapps" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Stop-Process -Name "SystemSettings" -Force -ErrorAction SilentlyContinue

    Write-Log "Chrome configurado como navegador predeterminado" "OK"
    $Script:SoftwareInstalado += "Chrome (default)"
}

function Install-Office {
    Write-StepHeader -Step 15 -Title "INSTALANDO OFFICE 2007 (Word, Excel, PowerPoint)"
    Show-ProgressCosmos -Step 15

    # Check idempotencia: Office 2007 ya instalado
    $OfficeYaInstalado = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and ($_.DisplayName -like "*Office*2007*" -or $_.DisplayName -like "*Office Enterprise*") }
    if ($OfficeYaInstalado) {
        $NombreOffice = ($OfficeYaInstalado | Select-Object -First 1).DisplayName
        Write-Log "Office ya esta instalado: $NombreOffice" "INFO"
        $Script:SoftwareInstalado += "Office 2007 (ya instalado)"
        return
    }

    # Ruta exacta del setup de Office 2007
    $SetupPath = "$RutaOffice\Ofice2007\setup.exe"
    $SerialFile = "$RutaOffice\Ofice2007\SERIAL.txt"

    # Leer serial del archivo si existe (formato: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)
    $SerialOffice = ""
    if (Test-Path $SerialFile) {
        $ContenidoSerial = Get-Content $SerialFile -First 1
        # Extraer solo el serial (quitar espacios y numeros extra al final)
        $SerialOffice = ($ContenidoSerial -split '\s+')[0].Trim() -replace '-', ''
        $SerialPreview = if ($SerialOffice.Length -ge 5) { $SerialOffice.Substring(0,5) } else { $SerialOffice }
        Write-Log "Serial de Office: $SerialPreview..."
    }

    if (-not $SerialOffice) {
        Write-Log "No se encontro serial de Office. Instalacion omitida" "ERROR"
        return
    }

    if (Test-Path $SetupPath) {
        # Crear archivo de configuracion para instalar solo Word, Excel, PowerPoint
        $ConfigXML = @"
<Configuration Product="Enterprise">
    <Display Level="basic" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
    <PIDKEY Value="$SerialOffice" />
    <OptionState Id="WORDFiles" State="local" Children="force" />
    <OptionState Id="EXCELFiles" State="local" Children="force" />
    <OptionState Id="PPTFiles" State="local" Children="force" />
    <OptionState Id="OUTLOOKFiles" State="absent" Children="force" />
    <OptionState Id="ACCESSFiles" State="absent" Children="force" />
    <OptionState Id="PUBFiles" State="absent" Children="force" />
    <OptionState Id="ONOTEFiles" State="absent" Children="force" />
    <OptionState Id="GROOVEFiles" State="absent" Children="force" />
    <OptionState Id="INFOPATHFiles" State="absent" Children="force" />
</Configuration>
"@
        $ConfigPath = "$env:TEMP\Office2007Config.xml"
        [System.IO.File]::WriteAllText($ConfigPath, $ConfigXML, (New-Object System.Text.UTF8Encoding $false))

        Write-Log "Instalando Office 2007 (Word, Excel, PowerPoint)..."

        # Iniciar melodia ambiental durante la instalacion
        Start-BackgroundMelody -Mensaje "Melodia del Santuario - Instalando Office..."

        Start-Process -FilePath $SetupPath -ArgumentList "/config `"$ConfigPath`"" -Wait -NoNewWindow

        # Detener melodia ambiental
        Stop-BackgroundMelody

        Write-Log "Office 2007 instalado" "OK"
        $Script:SoftwareInstalado += "Office 2007"
    } else {
        Write-Log "No se encontro setup.exe de Office en: $SetupPath" "ERROR"
    }
}

function Install-Dedalus {
    Write-StepHeader -Step 16 -Title "INSTALANDO DEDALUS EXPEDIENTE CLINICO"
    Show-ProgressCosmos -Step 16

    # Check idempotencia: Dedalus ya instalado con contenido
    if ((Test-Path "C:\Dedalus\xHIS") -and (Get-ChildItem "C:\Dedalus\xHIS" -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0) {
        Write-Log "Dedalus ya esta instalado en C:\Dedalus\xHIS" "INFO"
        $Script:SoftwareInstalado += "Dedalus (ya instalado)"
        return
    }

    $Netlogon = "$RutaDedalus\netlogon6.bat"

    if (Test-Path $Netlogon) {
        # Crear carpeta Dedalus si no existe
        if (-not (Test-Path "C:\Dedalus")) {
            New-Item -ItemType Directory -Path "C:\Dedalus" -Force | Out-Null
        }

        Write-Log "Ejecutando netlogon6.bat..."

        # Iniciar melodia ambiental durante la instalacion
        Start-BackgroundMelody -Mensaje "Melodia del Santuario - Instalando Dedalus..."

        Start-Process cmd -ArgumentList "/c `"$Netlogon`"" -Wait -NoNewWindow

        # Detener melodia ambiental
        Stop-BackgroundMelody

        Write-Log "Dedalus instalado" "OK"
        $Script:SoftwareInstalado += "Dedalus"
    } else {
        Write-Log "No se encontro netlogon6.bat en: $RutaDedalus" "ERROR"
    }
}

function Add-DedalusSyncStartup {
    Write-StepHeader -Step 17 -Title "LIMPIEZA DE SINCRONIZADOR Y CONFIGURACION"
    Show-ProgressCosmos -Step 17

    # =========================================================================
    # NOTA IMPORTANTE: netlogon6.bat configura TODO automaticamente
    # Esta funcion SOLO hace limpieza de entradas manuales creadas anteriormente
    # NO creamos accesos directos ni entradas de registro - netlogon6.bat lo hace
    # =========================================================================

    Write-Host "  Limpiando entradas manuales anteriores del sincronizador..." -ForegroundColor Yellow

    # 1. Eliminar script wrapper si existe (versiones anteriores del configurador)
    $WrapperFiles = @(
        "C:\Dedalus\HCG_SyncVisual.ps1",
        "C:\Dedalus\HCG_SyncWrapper.ps1",
        "C:\Dedalus\SyncWrapper.ps1",
        "C:\Dedalus\DedalusSync.ps1"
    )
    foreach ($WrapperFile in $WrapperFiles) {
        if (Test-Path $WrapperFile) {
            Remove-Item -Path $WrapperFile -Force -ErrorAction SilentlyContinue
            Write-Log "Eliminado wrapper anterior: $WrapperFile" "OK"
        }
    }

    # 2. Limpiar accesos directos MANUALES de Dedalus Sync en carpetas Startup
    # (Solo eliminamos los que creamos nosotros, no los de netlogon6.bat)
    $StartupFolders = @(
        "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    )

    # Tambien limpiar para todos los usuarios
    $UsersPath = "C:\Users"
    Get-ChildItem -Path $UsersPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $UserStartup = "$($_.FullName)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        if (Test-Path $UserStartup) {
            $StartupFolders += $UserStartup
        }
    }

    foreach ($StartupFolder in $StartupFolders) {
        if (Test-Path $StartupFolder) {
            # Solo eliminar "Dedalus Sync.lnk" que creamos nosotros manualmente
            $ManualShortcut = "$StartupFolder\Dedalus Sync.lnk"
            if (Test-Path $ManualShortcut) {
                Remove-Item -Path $ManualShortcut -Force -ErrorAction SilentlyContinue
                Write-Log "Eliminado acceso directo manual: $ManualShortcut" "OK"
            }
        }
    }

    # 3. Limpiar entradas del registro que creamos nosotros (Run keys)
    $RegistryRunKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
    )
    foreach ($RunKey in $RegistryRunKeys) {
        # Solo eliminar entradas que creamos manualmente
        $EntriesToRemove = @("DedalusSync", "HCG_Sync", "SyncDedalus", "xHIS_Sync", "Dedalus Sync")
        foreach ($EntryName in $EntriesToRemove) {
            try {
                $CurrentValue = Get-ItemProperty -Path $RunKey -Name $EntryName -ErrorAction SilentlyContinue
                if ($CurrentValue) {
                    Remove-ItemProperty -Path $RunKey -Name $EntryName -Force -ErrorAction SilentlyContinue
                    Write-Log "Eliminada entrada de registro manual: $RunKey\$EntryName" "OK"
                }
            } catch { }
        }
    }

    Write-Log "Limpieza de entradas manuales completada" "OK"

    # =========================================================================
    # CONFIGURACION DEL SERVIDOR Y CREDENCIALES
    # =========================================================================

    # Guardar credenciales del servidor en Windows Credential Manager
    try {
        $ServerIP = "10.2.1.17"
        $CredUser = "distribucion"
        $CredPass = "distribucion"

        cmdkey /delete:$ServerIP 2>$null | Out-Null
        cmdkey /add:$ServerIP /user:$CredUser /pass:$CredPass | Out-Null

        Write-Log "Credenciales del servidor $ServerIP guardadas en Windows" "OK"
    } catch {
        Write-Log "No se pudo guardar credenciales del servidor (no critico)" "WARN"
    }

    # Agregar servidor a zona de Intranet (evita bloqueos de seguridad)
    try {
        $IntranetZone = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\10.2.1.17"
        if (-not (Test-Path $IntranetZone)) {
            New-Item -Path $IntranetZone -Force | Out-Null
        }
        Set-ItemProperty -Path $IntranetZone -Name "file" -Value 1 -Type DWord -ErrorAction SilentlyContinue

        Write-Log "Servidor agregado a zona de Intranet (sin bloqueos)" "OK"
    } catch {
        Write-Log "No se pudo configurar zona de Intranet (no critico)" "WARN"
    }

    # =========================================================================
    # COPIAR SINCRONIZADOR AL STARTUP (sync_xhis6.bat)
    # =========================================================================
    $SyncSource = "\\10.2.1.17\distribucion\dedalus\sincronizador\sync_xhis6.bat"
    $StartupFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
    $SyncDestino = "$StartupFolder\sync_xhis6.bat"

    Write-Host "  Copiando sincronizador al inicio de Windows..." -ForegroundColor Cyan

    try {
        if (Test-Path $SyncSource) {
            # Copiar el archivo sync_xhis6.bat al Startup
            Copy-Item -Path $SyncSource -Destination $SyncDestino -Force -ErrorAction Stop

            # Desbloquear el archivo copiado
            Unblock-File -Path $SyncDestino -ErrorAction SilentlyContinue

            Write-Log "Sincronizador copiado a Startup: $SyncDestino" "OK"
            $Script:SoftwareInstalado += "Sync Dedalus (sync_xhis6.bat)"
        } else {
            Write-Log "No se encontro el sincronizador en: $SyncSource" "WARN"
        }
    } catch {
        Write-Log "Error al copiar sincronizador: $($_.Exception.Message)" "ERROR"
    }

    # Desbloquear archivos de Dedalus si existen
    if (Test-Path "C:\Dedalus") {
        Get-ChildItem "C:\Dedalus" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue
        }
        Write-Log "Archivos Dedalus desbloqueados" "OK"
    }

    Write-Log "Configuracion de sincronizador completada" "OK"

    # NOTA: Los accesos directos de Dedalus (xHIS, xFARMA, Escritorio Clinico)
    # se crean automaticamente por el sincronizador sync_xhis6.bat
    # No los creamos manualmente para evitar duplicados
}

function Install-Antivirus {
    Write-StepHeader -Step 18 -Title "INSTALANDO ANTIVIRUS ESET"
    Show-ProgressCosmos -Step 18

    # Verificar si ya esta instalado
    $ESETService = Get-Service -Name "ekrn" -ErrorAction SilentlyContinue
    if ($ESETService) {
        Write-Log "ESET ya esta instalado" "INFO"
        $Script:SoftwareInstalado += "ESET"
        return
    }

    # Buscar instalador con nombre exacto
    $Instalador = "$RutaAntivirus\PROTECT_Installer_x64_es_CL 2.exe"
    if (-not (Test-Path $Instalador)) {
        $Instalador = Find-Installer -Carpeta $RutaAntivirus -Filtro "*.exe"
    }

    if ($Instalador -and (Test-Path $Instalador)) {
        Write-Log "Instalando ESET PROTECT... (puede tardar varios minutos)"
        Write-Host "  Archivo: $Instalador" -ForegroundColor Gray

        # Iniciar melodia ambiental durante la instalacion
        Start-BackgroundMelody -Mensaje "Melodia del Santuario - Instalando Antivirus..."

        $Process = Start-Process -FilePath $Instalador -ArgumentList "--silent --accepteula" -Wait -NoNewWindow -PassThru

        # Detener melodia ambiental
        Stop-BackgroundMelody

        if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
            Write-Log "ESET instalado correctamente" "OK"
            $Script:SoftwareInstalado += "ESET"
        } else {
            # Verificar de todas formas
            Start-Sleep -Seconds 5
            if (Get-Service -Name "ekrn" -ErrorAction SilentlyContinue) {
                Write-Log "ESET instalado (verificado)" "OK"
                $Script:SoftwareInstalado += "ESET"
            } else {
                Write-Log "Verificar manualmente la instalacion de ESET (codigo: $($Process.ExitCode))" "WARN"
            }
        }
    } else {
        Write-Log "No se encontro instalador de ESET en: $RutaAntivirus" "ERROR"
    }
}

function Copy-AccesosDirectos {
    Write-StepHeader -Step 19 -Title "COPIANDO ACCESOS DIRECTOS"
    Show-ProgressCosmos -Step 19

    $Desktop = "C:\Users\Public\Desktop"

    if (Test-Path $RutaAccesos) {
        $Accesos = Get-ChildItem -Path $RutaAccesos -Include "*.lnk","*.url" -Recurse
        $Contador = 0
        foreach ($Acceso in $Accesos) {
            Copy-Item -Path $Acceso.FullName -Destination $Desktop -Force -ErrorAction SilentlyContinue
            $Contador++
        }
        Write-Log "$Contador accesos directos copiados" "OK"
        $Script:SoftwareInstalado += "Accesos"
    } else {
        Write-Log "No se encontro carpeta de accesos: $RutaAccesos" "WARN"
    }
}

function Remove-DuplicateDesktopIcons {
    param([switch]$SilentMode)

    if (-not $SilentMode) {
        Write-StepHeader -Step 20 -Title "LIMPIANDO ICONOS DUPLICADOS DEL ESCRITORIO"
        Show-ProgressCosmos -Step 20
    }

    # Nombres EXACTOS de iconos de Dedalus (case insensitive)
    # El script va a buscar cualquier icono que CONTENGA estas palabras clave
    $PalabrasClave = @(
        "xfarma",
        "x-farma",
        "xhis",
        "x-his",
        "escritorio clinico",
        "escritorioclinico",
        "hpresc",
        "dedalus"
    )

    # Rutas de escritorio
    $DesktopPaths = @(
        "C:\Users\Public\Desktop"
    )
    Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $UserDesktop = "$($_.FullName)\Desktop"
        if (Test-Path $UserDesktop) {
            $DesktopPaths += $UserDesktop
        }
    }

    $IconosEliminados = 0

    foreach ($DesktopPath in $DesktopPaths) {
        if (-not (Test-Path $DesktopPath)) { continue }

        # Obtener TODOS los .lnk del escritorio
        $AllShortcuts = Get-ChildItem -Path $DesktopPath -Filter "*.lnk" -ErrorAction SilentlyContinue

        foreach ($PalabraClave in $PalabrasClave) {
            # Buscar iconos que contengan la palabra clave (ignorando mayusculas)
            $MatchingIcons = $AllShortcuts | Where-Object {
                $_.BaseName.ToLower() -like "*$PalabraClave*"
            }

            if ($MatchingIcons.Count -gt 1) {
                # Hay duplicados! Ordenar para quedarnos con el "mejor" nombre
                # Prioridad: nombre mas corto y sin numeros al final
                $Sorted = $MatchingIcons | Sort-Object {
                    $name = $_.BaseName
                    # Penalizar nombres con (2), (3), - Copia, etc.
                    $penalty = 0
                    if ($name -match '\(\d+\)') { $penalty += 100 }
                    if ($name -match ' - Copia') { $penalty += 100 }
                    if ($name -match ' - Copy') { $penalty += 100 }
                    if ($name -match ' \d+$') { $penalty += 50 }
                    $penalty + $name.Length
                }

                # Mantener el primero (mejor), eliminar el resto
                $Keep = $Sorted | Select-Object -First 1
                $ToDelete = $Sorted | Select-Object -Skip 1

                foreach ($Dup in $ToDelete) {
                    Remove-Item -Path $Dup.FullName -Force -ErrorAction SilentlyContinue
                    if (-not $SilentMode) {
                        Write-Log "Eliminado duplicado: $($Dup.Name)" "OK"
                    }
                    $IconosEliminados++
                }
            }
        }
    }

    if (-not $SilentMode) {
        if ($IconosEliminados -gt 0) {
            Write-Log "Se eliminaron $IconosEliminados iconos duplicados del escritorio" "OK"
        } else {
            Write-Log "No se encontraron iconos duplicados" "OK"
        }
    }

    return $IconosEliminados
}

function Enable-WindowsUpdate {
    Write-StepHeader -Step 21 -Title "HABILITANDO ACTUALIZACIONES AUTOMATICAS DE WINDOWS"
    Show-ProgressCosmos -Step 21

    try {
        # Habilitar servicio de Windows Update
        $WUService = Get-Service -Name "wuauserv" -ErrorAction SilentlyContinue
        if ($WUService) {
            Set-Service -Name "wuauserv" -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue
            Write-Log "Servicio Windows Update habilitado" "OK"
        }

        # Habilitar servicio BITS (Background Intelligent Transfer)
        $BITSService = Get-Service -Name "BITS" -ErrorAction SilentlyContinue
        if ($BITSService) {
            Set-Service -Name "BITS" -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name "BITS" -ErrorAction SilentlyContinue
            Write-Log "Servicio BITS habilitado" "OK"
        }

        # Configurar politicas de Windows Update via registro
        $WUPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

        # Crear la ruta si no existe
        if (-not (Test-Path $WUPolicyPath)) {
            New-Item -Path $WUPolicyPath -Force | Out-Null
        }

        # Configurar actualizaciones automaticas
        # AUOptions: 4 = Auto download and schedule install
        Set-ItemProperty -Path $WUPolicyPath -Name "NoAutoUpdate" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $WUPolicyPath -Name "AUOptions" -Value 4 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $WUPolicyPath -Name "ScheduledInstallDay" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue  # 0 = todos los dias
        Set-ItemProperty -Path $WUPolicyPath -Name "ScheduledInstallTime" -Value 3 -Type DWord -Force -ErrorAction SilentlyContinue  # 3 AM

        Write-Log "Politicas de actualizacion automatica configuradas" "OK"

        # =========================================================================
        # HABILITAR FEATURE UPDATES (Windows 11 y actualizaciones mayores)
        # =========================================================================
        $WUPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        if (-not (Test-Path $WUPath)) {
            New-Item -Path $WUPath -Force | Out-Null
        }

        # Permitir Feature Updates (actualizaciones a nuevas versiones como Windows 11)
        # TargetReleaseVersion = 0 significa "no restringir version"
        Remove-ItemProperty -Path $WUPath -Name "TargetReleaseVersion" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $WUPath -Name "TargetReleaseVersionInfo" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path $WUPath -Name "ProductVersion" -Force -ErrorAction SilentlyContinue

        # Habilitar "Obtener las ultimas actualizaciones tan pronto esten disponibles"
        # Esto acelera la recepcion de Feature Updates
        $WUSettingsPath = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
        if (-not (Test-Path $WUSettingsPath)) {
            New-Item -Path $WUSettingsPath -Force | Out-Null
        }
        Set-ItemProperty -Path $WUSettingsPath -Name "IsContinuousInnovationOptedIn" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $WUSettingsPath -Name "AllowMUUpdateService" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue

        Write-Log "Feature Updates habilitados (Windows 11 disponible)" "OK"

        # Remover restricciones de GPO si existen (que bloquean updates)
        $RestrictionsPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        $RestrictionsToRemove = @("DisableWindowsUpdateAccess", "DoNotConnectToWindowsUpdateInternetLocations", "WUServer", "WUStatusServer", "DeferFeatureUpdates", "DeferFeatureUpdatesPeriodInDays")
        foreach ($Restriction in $RestrictionsToRemove) {
            Remove-ItemProperty -Path $RestrictionsPath -Name $Restriction -Force -ErrorAction SilentlyContinue
        }

        # Limpiar configuracion de usuario que pueda bloquear
        $UserWUPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WindowsUpdate"
        if (Test-Path $UserWUPath) {
            Remove-ItemProperty -Path $UserWUPath -Name "DisableWindowsUpdateAccess" -Force -ErrorAction SilentlyContinue
        }

        Write-Log "Restricciones de Windows Update eliminadas" "OK"

        # Forzar deteccion de actualizaciones (no esperar el resultado)
        Start-Process -FilePath "UsoClient.exe" -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Process -FilePath "UsoClient.exe" -ArgumentList "StartDownload" -WindowStyle Hidden -ErrorAction SilentlyContinue

        Write-Log "Windows Update habilitado - actualizaciones automaticas y Feature Updates (Win11)" "OK"
        $Script:SoftwareInstalado += "Windows Update Auto"

    } catch {
        Write-Log "Error al configurar Windows Update: $($_.Exception.Message)" "WARN"
    }
}

function Install-AllWindowsUpdates {
    Write-StepHeader -Step 21.5 -Title "INSTALANDO TODAS LAS ACTUALIZACIONES DE WINDOWS"
    Show-ProgressCosmos -Step 21

    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "  |  Instalando TODAS las actualizaciones pendientes:        |" -ForegroundColor Cyan
    Write-Host "  |  - Windows 11 25H2 (si disponible)                       |" -ForegroundColor Cyan
    Write-Host "  |  - Drivers (Lenovo, Intel, Realtek, etc.)                |" -ForegroundColor Cyan
    Write-Host "  |  - .NET Framework                                        |" -ForegroundColor Cyan
    Write-Host "  |  - Actualizaciones de seguridad                          |" -ForegroundColor Cyan
    Write-Host "  |                                                          |" -ForegroundColor Cyan
    Write-Host "  |  Esto puede tardar 15-45 minutos...                      |" -ForegroundColor Cyan
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""

    # Iniciar melodia ambiental
    Start-BackgroundMelody -Mensaje "Melodia del Santuario - Actualizando Windows..."

    try {
        # Metodo 1: Usar Windows Update via COM Object (nativo, no requiere modulos)
        Write-Log "Buscando actualizaciones disponibles..." "INFO"

        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

        Write-Host "  Escaneando actualizaciones (puede tardar varios minutos)..." -ForegroundColor Yellow
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

        $UpdatesToInstall = $SearchResult.Updates
        $TotalUpdates = $UpdatesToInstall.Count

        if ($TotalUpdates -eq 0) {
            Write-Log "No hay actualizaciones pendientes - El sistema esta al dia" "OK"
            Stop-BackgroundMelody
            $Script:SoftwareInstalado += "Windows Actualizado"
            return
        }

        Write-Log "Se encontraron $TotalUpdates actualizaciones pendientes" "INFO"
        Write-Host ""

        # Mostrar lista de actualizaciones
        $UpdateCounter = 0
        foreach ($Update in $UpdatesToInstall) {
            $UpdateCounter++
            $SizeMB = [math]::Round($Update.MaxDownloadSize / 1MB, 1)
            Write-Host "  [$UpdateCounter/$TotalUpdates] $($Update.Title) ($SizeMB MB)" -ForegroundColor Gray
        }
        Write-Host ""

        # Crear coleccion para descargar
        $UpdatesCollection = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($Update in $UpdatesToInstall) {
            $UpdatesCollection.Add($Update) | Out-Null
        }

        # Descargar actualizaciones
        Write-Log "Descargando $TotalUpdates actualizaciones..." "INFO"
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesCollection

        try {
            $DownloadResult = $Downloader.Download()
            if ($DownloadResult.ResultCode -eq 2) {
                Write-Log "Descarga completada exitosamente" "OK"
            } else {
                Write-Log "Descarga completada con codigo: $($DownloadResult.ResultCode)" "WARN"
            }
        } catch {
            Write-Log "Error en descarga: $($_.Exception.Message)" "WARN"
        }

        # Instalar actualizaciones
        Write-Log "Instalando $TotalUpdates actualizaciones..." "INFO"
        Write-Host "  Este proceso puede tardar 15-45 minutos. NO apagues el equipo." -ForegroundColor Yellow
        Write-Host ""

        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $UpdatesCollection

        try {
            $InstallResult = $Installer.Install()

            $Installed = 0
            $Failed = 0
            $RebootRequired = $false

            for ($i = 0; $i -lt $UpdatesCollection.Count; $i++) {
                $UpdateResult = $InstallResult.GetUpdateResult($i)
                if ($UpdateResult.ResultCode -eq 2) {
                    $Installed++
                } else {
                    $Failed++
                }
                if ($UpdateResult.RebootRequired) {
                    $RebootRequired = $true
                }
            }

            Write-Host ""
            Write-Log "$Installed actualizaciones instaladas correctamente" "OK"
            if ($Failed -gt 0) {
                Write-Log "$Failed actualizaciones fallaron (se pueden reintentar despues)" "WARN"
            }

            if ($RebootRequired) {
                Write-Host ""
                Write-Host "  ============================================================" -ForegroundColor Yellow
                Write-Host "  |  REINICIO REQUERIDO                                      |" -ForegroundColor Yellow
                Write-Host "  |  Algunas actualizaciones se completaran tras reiniciar   |" -ForegroundColor Yellow
                Write-Host "  ============================================================" -ForegroundColor Yellow
                $Script:SoftwareInstalado += "Updates ($Installed - Reinicio pendiente)"
            } else {
                $Script:SoftwareInstalado += "Updates ($Installed instaladas)"
            }

        } catch {
            Write-Log "Error durante instalacion: $($_.Exception.Message)" "WARN"
            $Script:SoftwareInstalado += "Updates (parcial)"
        }

    } catch {
        Write-Log "Error en Windows Update: $($_.Exception.Message)" "WARN"

        # Metodo alternativo: usar usoclient
        Write-Log "Intentando metodo alternativo..." "INFO"
        try {
            Start-Process -FilePath "usoclient.exe" -ArgumentList "StartInteractiveScan" -Wait -NoNewWindow -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            Start-Process -FilePath "usoclient.exe" -ArgumentList "StartDownload" -Wait -NoNewWindow -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            Start-Process -FilePath "usoclient.exe" -ArgumentList "StartInstall" -Wait -NoNewWindow -ErrorAction SilentlyContinue
            Write-Log "Windows Update iniciado via usoclient" "OK"
            $Script:SoftwareInstalado += "Updates (usoclient)"
        } catch {
            Write-Log "No se pudo iniciar Windows Update automatico" "WARN"
        }
    }

    Stop-BackgroundMelody
    Write-Host ""
}

function Remove-AdminUsuarioActual {
    Write-StepHeader -Step 22 -Title "QUITANDO PRIVILEGIOS DE ADMINISTRADOR"
    Show-ProgressCosmos -Step 22

    $UsuarioActual = $Script:UsuarioOriginal

    # No quitar admin si es el usuario Soporte o Administrator
    if ($UsuarioActual -eq $UsuarioSoporte -or $UsuarioActual -eq "Administrator" -or $UsuarioActual -eq "Administrador") {
        Write-Log "Usuario '$UsuarioActual' mantiene privilegios de administrador" "INFO"
        return
    }

    # Verificar si el usuario actual es administrador
    $EsAdmin = Get-LocalGroupMember -Group $GrupoAdmin -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*\$UsuarioActual" }

    if ($EsAdmin) {
        try {
            Remove-LocalGroupMember -Group $GrupoAdmin -Member $UsuarioActual -ErrorAction Stop
            Write-Log "Privilegios de administrador removidos de '$UsuarioActual'" "OK"
            Write-Host "  IMPORTANTE: El usuario '$UsuarioActual' ya no es administrador" -ForegroundColor Yellow
            Write-Host "  Solo el usuario '$UsuarioSoporte' tiene privilegios de administrador" -ForegroundColor Yellow
            $Script:SoftwareInstalado += "Admin removido"
        } catch {
            Write-Log "No se pudieron remover privilegios de '$UsuarioActual': $($_.Exception.Message)" "WARN"
        }
    } else {
        Write-Log "Usuario '$UsuarioActual' no era administrador" "INFO"
    }
}

function Install-ReporteIP {
    Write-StepHeader -Step 23 -Title "CONFIGURANDO REPORTE AUTOMATICO DE IP"
    Show-ProgressCosmos -Step 23

    $ScriptPath = "C:\HCG_Logs\report_ip.ps1"

    # Crear el script de reporte de IP
    $ScriptContent = @'
# HCG - Reporte automatico de IP (cada 3 horas + al iniciar sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $WiFiAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Wi-Fi*" -or $_.Name -like "*WiFi*" -or $_.Name -like "*Wireless*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }
    $IPEthernet = ""
    if ($EthAdapter -and $EthAdapter.Status -eq "Up") {
        $IPEthernet = (Get-NetIPAddress -InterfaceIndex $EthAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
    }
    $IPWiFi = ""
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq "Up") {
        $IPWiFi = (Get-NetIPAddress -InterfaceIndex $WiFiAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1).IPAddress
    }
    $SSIDWiFi = ""
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq "Up") {
        $NetshOutput = netsh wlan show interfaces 2>$null
        if ($NetshOutput) {
            $SSIDLine = $NetshOutput | Select-String "^\s+SSID\s+:" | Select-Object -First 1
            if ($SSIDLine) { $SSIDWiFi = ($SSIDLine.ToString() -replace "^\s+SSID\s+:\s+", "").Trim() }
        }
    }
    if (-not $IPEthernet -and -not $IPWiFi) { exit }
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }
    $Body = @{
        Accion = "ip"; MACEthernet = $MACEthernet; IPEthernet = $IPEthernet
        IPWiFi = $IPWiFi; SSIDWiFi = $SSIDWiFi; NombreEquipo = $env:COMPUTERNAME
        FechaReporte = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 30 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) } }
    }
} catch { exit }
'@

    Initialize-HCGLogsFolder | Out-Null
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de reporte de IP creado en $ScriptPath" "OK"

    # Crear launcher VBS para ejecucion completamente oculta
    $VbsLauncher = New-HiddenLauncher -PowerShellScriptPath $ScriptPath
    Write-Log "Launcher VBS creado: $VbsLauncher" "OK"

    # Crear tarea programada: cada 3 horas + al iniciar sesion
    $TaskName = "HCG_ReporteIP"

    try {
        # Usar wscript.exe con VBS para evitar cualquier ventana visible
        $Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsLauncher`""

        # Trigger 1: Al iniciar sesion (cualquier usuario), delay 60 segundos
        $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
        $TriggerLogon.Delay = "PT60S"

        # Trigger 2: Cada 3 horas, indefinidamente
        $TriggerRepeat = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 3)

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
            -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $TaskName -Action $Action `
            -Trigger @($TriggerLogon, $TriggerRepeat) `
            -Settings $Settings `
            -Description "HCG - Reporte automatico de IP cada 3 horas" `
            -RunLevel Limited -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (cada 3 horas + al iniciar sesion)" "OK"
        } else {
            Write-Log "No se pudo crear la tarea programada" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "Reporte IP"
}

function Install-ReporteSistema {
    Write-StepHeader -Step 24 -Title "CONFIGURANDO REPORTE DE SISTEMA"
    Show-ProgressCosmos -Step 24

    $ScriptPath = "C:\HCG_Logs\report_system.ps1"

    # Crear el script de reporte de sistema
    $ScriptContent = @'
# HCG - Reporte de sistema y limpieza automatica (cada inicio de sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 180)

    # --- Identificar equipo por MAC Ethernet ---
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }

    # === 1. LIMPIEZA DE ARCHIVOS TEMPORALES (solo archivos > 1 dia) ===
    $BytesLimpiados = 0
    $TempFolders = @("$env:TEMP", "C:\Windows\Temp", "$env:LOCALAPPDATA\Temp")
    foreach ($folder in $TempFolders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                ForEach-Object {
                    $sz = $_.Length; $fp = $_.FullName
                    Remove-Item $fp -Force -ErrorAction SilentlyContinue
                    if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
                }
        }
    }
    if (Test-Path "C:\Windows\Prefetch") {
        Get-ChildItem "C:\Windows\Prefetch" -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
            ForEach-Object {
                $sz = $_.Length; $fp = $_.FullName
                Remove-Item $fp -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
            }
    }
    $WUFolder = "C:\Windows\SoftwareDistribution\Download"
    if (Test-Path $WUFolder) {
        Get-ChildItem $WUFolder -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
            ForEach-Object {
                $sz = $_.Length; $fp = $_.FullName
                Remove-Item $fp -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
            }
    }
    $MBLimpiados = [math]::Round($BytesLimpiados / 1MB, 1)

    # === 2. IMPRESORAS INSTALADAS ===
    $PrinterList = @()
    Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
        $Name = $_.Name; $PortName = $_.PortName; $Type = "Local"; $IP = ""
        if ($PortName -like "*USB*") { $Type = "USB" }
        elseif ($PortName -match "\d+\.\d+\.\d+\.\d+") {
            $Type = "Red"; $IP = [regex]::Match($PortName, "\d+\.\d+\.\d+\.\d+").Value
        } else {
            $Port = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
            if ($Port -and $Port.PrinterHostAddress) { $Type = "Red"; $IP = $Port.PrinterHostAddress }
        }
        if ($IP) { $PrinterList += "$Name [$Type - $IP]" } else { $PrinterList += "$Name [$Type]" }
    }

    # === 3. USUARIOS DEL SISTEMA ===
    $UserList = @()
    $AdminMembers = @()
    try {
        $AdminGroupName = (Get-LocalGroup -ErrorAction SilentlyContinue | Where-Object { $_.SID.Value -eq 'S-1-5-32-544' }).Name
        if ($AdminGroupName) {
            $AdminMembers = Get-LocalGroupMember -Group $AdminGroupName -ErrorAction SilentlyContinue |
                ForEach-Object { ($_.Name -split '\\')[-1] }
        }
    } catch {}
    Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
        $UName = $_.Name; $IsAdmin = $AdminMembers -contains $UName
        if ($IsAdmin) { $UserList += "$UName [Admin]" } else { $UserList += $UName }
    }

    # === 4. APLICACIONES INSTALADAS ===
    $Apps = @()
    $RegPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" -and $_.DisplayName -notlike "Update for*" -and $_.DisplayName -notlike "Security Update*" } |
            ForEach-Object { $Apps += $_.DisplayName.Trim() }
    }
    $Apps = $Apps | Select-Object -Unique | Sort-Object

    # === 5. ACCESOS DIRECTOS DEL ESCRITORIO ===
    $Shortcuts = @()
    $DesktopPaths = @("C:\Users\Public\Desktop")
    if ($env:USERPROFILE -and (Test-Path "$env:USERPROFILE\Desktop")) { $DesktopPaths += "$env:USERPROFILE\Desktop" }
    foreach ($dp in $DesktopPaths) {
        if (Test-Path $dp) {
            Get-ChildItem $dp -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object { $Shortcuts += $_.BaseName }
            Get-ChildItem $dp -Filter "*.url" -ErrorAction SilentlyContinue | ForEach-Object { $Shortcuts += $_.BaseName + " (web)" }
        }
    }
    $Shortcuts = $Shortcuts | Select-Object -Unique | Sort-Object

    # === 6. ESPACIO LIBRE EN DISCO ===
    $Disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $EspacioLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $EspacioTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # === 7. VERSION DE WINDOWS ===
    $WinVer = ""
    try {
        $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        $Build = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue)
        $DisplayVersion = $Build.DisplayVersion
        $WinVer = "$($OS.Caption) $DisplayVersion (Build $($Build.CurrentBuild))"
    } catch { $WinVer = "Desconocida" }

    # === 8. VERIFICAR INTERNET Y ENVIAR (con reintentos) ===
    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }

    # Agregar version de Windows al inicio de las apps instaladas
    $AppsConWindows = @($WinVer) + $Apps

    $Body = @{
        Accion            = "sistema"
        MACEthernet       = $MACEthernet
        NombreEquipo      = $env:COMPUTERNAME
        Impresoras        = ($PrinterList -join " | ")
        Usuarios          = ($UserList -join " | ")
        AppsInstaladas    = ($AppsConWindows -join " | ")
        AccesosEscritorio = ($Shortcuts -join " | ")
        EspacioLibreGB    = "$EspacioLibreGB / $EspacioTotalGB GB"
        MBLimpiados       = $MBLimpiados
        FechaReporte      = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json -Depth 3

    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 60 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45) } }
    }
} catch { exit }
'@

    Initialize-HCGLogsFolder | Out-Null
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de reporte de sistema creado en $ScriptPath" "OK"

    # Crear launcher VBS para ejecucion completamente oculta
    $VbsLauncher = New-HiddenLauncher -PowerShellScriptPath $ScriptPath
    Write-Log "Launcher VBS creado: $VbsLauncher" "OK"

    # Crear tarea programada: solo al iniciar sesion, delay 2 minutos
    $TaskName = "HCG_ReporteSistema"

    try {
        # Usar wscript.exe con VBS para evitar cualquier ventana visible
        $Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsLauncher`""

        # Trigger: Al iniciar sesion (cualquier usuario), delay 120 segundos
        $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
        $TriggerLogon.Delay = "PT120S"

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
            -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $TaskName -Action $Action `
            -Trigger $TriggerLogon `
            -Settings $Settings `
            -Description "HCG - Reporte de sistema y limpieza al iniciar sesion" `
            -RunLevel Limited -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (al iniciar sesion, delay 2 min)" "OK"
        } else {
            Write-Log "No se pudo crear la tarea programada de sistema" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "Reporte Sistema"
}

function Install-ReporteDiagnostico {
    Write-StepHeader -Step 25 -Title "CONFIGURANDO REPORTE DE DIAGNOSTICO DE SALUD"
    Show-ProgressCosmos -Step 25

    $ScriptPath = "C:\HCG_Logs\report_diagnostico.ps1"

    # Crear el script de diagnostico de salud
    $ScriptContent = @'
# HCG - Reporte de diagnostico de salud (cada 4 horas + al iniciar sesion)
$ErrorActionPreference = "SilentlyContinue"
try {
    $GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)

    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*Ethernet*" } | Select-Object -First 1
    $MACEthernet = ""
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace "-", "").ToUpper() }
    if (-not $MACEthernet) { exit }

    $TestOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName "dns.google" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri "https://www.google.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }

    # --- RAM ---
    $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $RAMTotalGB = if ($OS.TotalVisibleMemorySize) { [math]::Round($OS.TotalVisibleMemorySize / 1MB, 1) } else { 0 }
    $RAMLibreGB = if ($OS.FreePhysicalMemory) { [math]::Round($OS.FreePhysicalMemory / 1MB, 1) } else { 0 }
    $RAMUsadaGB = [math]::Round($RAMTotalGB - $RAMLibreGB, 1)
    $RAMPct = if ($RAMTotalGB -gt 0) { [math]::Round(($RAMUsadaGB / $RAMTotalGB) * 100, 0) } else { 0 }

    # --- Top 5 procesos ---
    $Top5 = Get-Process -ErrorAction SilentlyContinue |
        Sort-Object WorkingSet64 -Descending | Select-Object -First 5 |
        ForEach-Object { "$($_.ProcessName) ($([math]::Round($_.WorkingSet64 / 1MB, 0)) MB)" }
    $Top5Str = $Top5 -join " | "

    # --- Chrome ---
    $ChromeProcs = Get-Process -Name "chrome" -ErrorAction SilentlyContinue
    $ChromeMB = 0; $ChromeCount = 0
    if ($ChromeProcs) {
        $ChromeCount = @($ChromeProcs).Count
        $ChromeMB = [math]::Round(($ChromeProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    # --- Dedalus ---
    $DedalusProcs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like "*dedalus*" }
    $DedalusMB = 0; $DedalusCount = 0
    if ($DedalusProcs) {
        $DedalusCount = @($DedalusProcs).Count
        $DedalusMB = [math]::Round(($DedalusProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    $TotalProcs = @(Get-Process -ErrorAction SilentlyContinue).Count
    $CPU = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $CPUPct = if ($CPU) { $CPU.LoadPercentage } else { 0 }
    if (-not $CPUPct) { $CPUPct = 0 }

    $PageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue | Select-Object -First 1
    $PageFileUsadoMB = if ($PageFile) { $PageFile.CurrentUsage } else { 0 }
    $PageFileTotalMB = if ($PageFile) { $PageFile.AllocatedBaseSize } else { 0 }

    $LastBoot = $OS.LastBootUpTime
    $UptimeDias = if ($LastBoot) { [math]::Round(((Get-Date) - $LastBoot).TotalDays, 1) } else { 0 }

    $Disco = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
    $DiscoLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $DiscoTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # --- Estado ---
    $Estado = "OK"; $Recomendaciones = @()
    if ($RAMPct -gt 85) {
        $Estado = "Critico"; $Recomendaciones += "RAM critica ($RAMPct%). Se recomienda ampliar memoria"
    } elseif ($RAMPct -gt 70) {
        $Estado = "Atencion"; $Recomendaciones += "RAM elevada ($RAMPct%). Monitorear uso. Considerar ampliacion"
    } else {
        $Recomendaciones += "Equipo operando con recursos suficientes"
    }
    if ($ChromeMB -gt 1500) { $Recomendaciones += "Chrome consumiendo $ChromeMB MB. Reducir pestanas" }
    if ($UptimeDias -gt 15) { $Recomendaciones += "Sin reinicio hace $UptimeDias dias. Reiniciar pronto" }
    if ($DiscoLibreGB -lt 20) { $Recomendaciones += "Disco bajo: $DiscoLibreGB GB libres. Liberar espacio" }
    $RecomendacionStr = $Recomendaciones -join " | "

    $Body = @{
        Accion = "diagnostico"; MACEthernet = $MACEthernet; NombreEquipo = $env:COMPUTERNAME
        RAMTotalGB = $RAMTotalGB; RAMUsadaGB = $RAMUsadaGB; RAMLibreGB = $RAMLibreGB; RAMPct = $RAMPct
        Top5Procesos = $Top5Str; ChromeMB = $ChromeMB; ChromeProcs = $ChromeCount
        DedalusMB = $DedalusMB; DedalusProcs = $DedalusCount; TotalProcs = $TotalProcs
        CPUPct = $CPUPct; PageFileUsado = $PageFileUsadoMB; PageFileTotal = $PageFileTotalMB
        UptimeDias = $UptimeDias; DiscoLibreGB = "$DiscoLibreGB / $DiscoTotalGB GB"
        Estado = $Estado; Recomendacion = $RecomendacionStr
        FechaReporte = (Get-Date -Format "dd/MM/yyyy HH:mm")
    } | ConvertTo-Json
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType "application/json; charset=utf-8" -TimeoutSec 30 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) } }
    }
} catch { exit }
'@

    Initialize-HCGLogsFolder | Out-Null
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de diagnostico creado en $ScriptPath" "OK"

    # Crear launcher VBS para ejecucion completamente oculta
    $VbsLauncher = New-HiddenLauncher -PowerShellScriptPath $ScriptPath
    Write-Log "Launcher VBS creado: $VbsLauncher" "OK"

    # Crear tarea programada: cada 4 horas + al iniciar sesion (delay 3 min)
    $TaskName = "HCG_ReporteDiagnostico"

    try {
        # Usar wscript.exe con VBS para evitar cualquier ventana visible
        $Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsLauncher`""

        # Trigger 1: Al iniciar sesion (cualquier usuario), delay 180 segundos
        $TriggerLogon = New-ScheduledTaskTrigger -AtLogOn
        $TriggerLogon.Delay = "PT180S"

        # Trigger 2: Cada 4 horas, indefinidamente
        $TriggerRepeat = New-ScheduledTaskTrigger -Once -At "00:00" -RepetitionInterval (New-TimeSpan -Hours 4)

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
            -MultipleInstances IgnoreNew

        Register-ScheduledTask -TaskName $TaskName -Action $Action `
            -Trigger @($TriggerLogon, $TriggerRepeat) `
            -Settings $Settings `
            -Description "HCG - Diagnostico de salud cada 4 horas" `
            -RunLevel Limited -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (cada 4 horas + al iniciar sesion, delay 3 min)" "OK"
        } else {
            Write-Log "No se pudo crear la tarea programada de diagnostico" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "Reporte Diagnostico"
}

function Install-AutoUpdate {
    Write-StepHeader -Step 25.5 -Title "CONFIGURANDO ACTUALIZACIONES AUTOMATICAS (CADA 3 MESES)"
    Show-ProgressCosmos -Step 25

    # Crear script simple que usa usoclient (comando nativo de Windows)
    $ScriptPath = "C:\HCG_Logs\auto_update.ps1"

    $ScriptContent = @'
# HCG - Actualizaciones automaticas cada 3 meses
# Usa usoclient (comando nativo de Windows) - seguro y estable
$ErrorActionPreference = "SilentlyContinue"

# Log
$LogFile = "C:\HCG_Logs\auto_update.log"
$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $LogFile -Value "[$Date] Iniciando actualizaciones automaticas..."

try {
    # Buscar actualizaciones
    Start-Process -FilePath "usoclient.exe" -ArgumentList "StartScan" -Wait -WindowStyle Hidden
    Start-Sleep -Seconds 60

    # Descargar actualizaciones
    Start-Process -FilePath "usoclient.exe" -ArgumentList "StartDownload" -Wait -WindowStyle Hidden
    Start-Sleep -Seconds 120

    # Instalar actualizaciones
    Start-Process -FilePath "usoclient.exe" -ArgumentList "StartInstall" -Wait -WindowStyle Hidden

    Add-Content -Path $LogFile -Value "[$Date] Actualizaciones completadas"

    # Si hay reinicio pendiente, reiniciar en 5 minutos
    $RebootPending = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue)
    if ($RebootPending) {
        Add-Content -Path $LogFile -Value "[$Date] Reinicio programado en 5 minutos"
        shutdown /r /t 300 /c "Actualizaciones de Windows instaladas. El equipo se reiniciara en 5 minutos."
    }
} catch {
    Add-Content -Path $LogFile -Value "[$Date] Error: $($_.Exception.Message)"
}
'@

    Initialize-HCGLogsFolder | Out-Null
    $ScriptContent | Out-File -FilePath $ScriptPath -Encoding UTF8 -Force
    Write-Log "Script de actualizaciones automaticas creado" "OK"

    # Crear launcher VBS para ejecucion oculta
    $VbsLauncher = New-HiddenLauncher -PowerShellScriptPath $ScriptPath
    Write-Log "Launcher VBS creado" "OK"

    # Crear tarea programada: cada 3 meses, a las 2:00 AM
    $TaskName = "HCG_AutoUpdate"

    try {
        # Eliminar tarea anterior si existe
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

        $Action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$VbsLauncher`""

        # Trigger: cada 3 meses (90 dias), a las 2:00 AM
        $Trigger = New-ScheduledTaskTrigger -Daily -At "02:00AM" -DaysInterval 90

        $Settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -WakeToRun

        $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

        Register-ScheduledTask -TaskName $TaskName `
            -Action $Action `
            -Trigger $Trigger `
            -Settings $Settings `
            -Principal $Principal `
            -Description "HCG - Actualizaciones automaticas de Windows cada 3 meses" `
            -Force | Out-Null

        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            Write-Log "Tarea programada '$TaskName' creada (cada 3 meses a las 2:00 AM)" "OK"
        } else {
            Write-Log "No se pudo verificar la tarea programada" "WARN"
        }
    } catch {
        Write-Log "Error al crear tarea programada: $($_.Exception.Message)" "WARN"
    }

    $Script:SoftwareInstalado += "AutoUpdate 3M"
}

function Verify-Configuracion {
    param([string]$NumInventario)

    $NombreUsuarioEsperado = if ($Script:EsOPD) { "OPD" } else { $NumInventario }

    # Limpieza final de iconos duplicados (silenciosa)
    Write-Host "  Limpiando iconos duplicados del escritorio..." -ForegroundColor Gray
    $DupsRemoved = Remove-DuplicateDesktopIcons -SilentMode
    if ($DupsRemoved -gt 0) {
        Write-Host "  Se eliminaron $DupsRemoved iconos duplicados" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  ======================================================" -ForegroundColor DarkYellow
    Write-Host "    VERIFICACION FINAL DE CONFIGURACION" -ForegroundColor Yellow
    Write-Host "  ======================================================" -ForegroundColor DarkYellow
    Write-Host ""

    $Checks = @()
    $Errores = 0

    # --- Usuarios ---
    $SoporteUser = Get-LocalUser -Name $UsuarioSoporte -ErrorAction SilentlyContinue
    if ($SoporteUser -and $SoporteUser.Enabled) {
        $EsAdmin = Get-LocalGroupMember -Group $GrupoAdmin -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*\$UsuarioSoporte" }
        if ($EsAdmin) {
            $Checks += @{ Status = "OK"; Msg = "Usuario $UsuarioSoporte existe (Administrador)" }
        } else {
            $Checks += @{ Status = "WARN"; Msg = "Usuario $UsuarioSoporte existe pero NO es Administrador" }
            Add-LocalGroupMember -Group $GrupoAdmin -Member $UsuarioSoporte -ErrorAction SilentlyContinue
        }
    } else {
        $Checks += @{ Status = "FAIL"; Msg = "Usuario $UsuarioSoporte NO existe o esta deshabilitado" }
        $Errores++
    }

    $NormalUser = Get-LocalUser -Name $NombreUsuarioEsperado -ErrorAction SilentlyContinue
    if ($NormalUser -and $NormalUser.Enabled) {
        $Checks += @{ Status = "OK"; Msg = "Usuario $NombreUsuarioEsperado existe (Estandar)" }
    } else {
        $Checks += @{ Status = "FAIL"; Msg = "Usuario $NombreUsuarioEsperado NO existe - intentando crear..." }
        $Errores++
        try {
            $DescUsuario = if ($Script:EsOPD) { "Usuario OPD - HCG" } else { "Usuario Equipo $NumInventario - HCG FAA" }
            New-LocalUser -Name $NombreUsuarioEsperado -NoPassword -Description $DescUsuario -PasswordNeverExpires -UserMayNotChangePassword -ErrorAction Stop | Out-Null
            Add-LocalGroupMember -Group $GrupoUsuarios -Member $NombreUsuarioEsperado -ErrorAction SilentlyContinue
            Enable-LocalUser -Name $NombreUsuarioEsperado -ErrorAction SilentlyContinue
            $Checks += @{ Status = "OK"; Msg = "Usuario $NombreUsuarioEsperado creado exitosamente" }
            $Errores--
        } catch {
            $Checks += @{ Status = "FAIL"; Msg = "No se pudo crear usuario: $($_.Exception.Message)" }
        }
    }

    # --- Software ---
    $RegPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
    $AppsInstaladas = @()
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | ForEach-Object { $AppsInstaladas += $_.DisplayName }
    }

    # WinRAR
    if ($AppsInstaladas | Where-Object { $_ -like "*WinRAR*" }) {
        $Checks += @{ Status = "OK"; Msg = "WinRAR instalado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "WinRAR NO detectado" }
    }

    # .NET 3.5
    $DotNet35 = Get-WindowsOptionalFeature -Online -FeatureName "NetFx3" -ErrorAction SilentlyContinue
    if ($DotNet35 -and $DotNet35.State -eq "Enabled") {
        $Checks += @{ Status = "OK"; Msg = ".NET 3.5 habilitado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = ".NET 3.5 NO habilitado" }
    }

    # Chrome
    if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
        $Checks += @{ Status = "OK"; Msg = "Chrome instalado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "Chrome NO detectado" }
    }

    # Acrobat Reader
    if ($AppsInstaladas | Where-Object { $_ -like "*Acrobat*Reader*" -or $_ -like "*Adobe*Reader*" }) {
        $Checks += @{ Status = "OK"; Msg = "Acrobat Reader instalado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "Acrobat Reader NO detectado" }
    }

    # Office
    if ($AppsInstaladas | Where-Object { $_ -like "*Office*" }) {
        $Checks += @{ Status = "OK"; Msg = "Office detectado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "Office NO detectado" }
    }

    # ESET Antivirus
    $ESETService = Get-Service -Name "ekrn" -ErrorAction SilentlyContinue
    if ($ESETService -and $ESETService.Status -eq "Running") {
        $Checks += @{ Status = "OK"; Msg = "ESET Antivirus activo" }
    } elseif ($ESETService) {
        $Checks += @{ Status = "WARN"; Msg = "ESET Antivirus instalado pero no corriendo" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "ESET Antivirus NO detectado" }
    }

    # --- Tareas programadas ---
    foreach ($tarea in @("HCG_ReporteIP", "HCG_ReporteSistema", "HCG_ReporteDiagnostico")) {
        if (Get-ScheduledTask -TaskName $tarea -ErrorAction SilentlyContinue) {
            $Checks += @{ Status = "OK"; Msg = "Tarea $tarea configurada" }
        } else {
            $Checks += @{ Status = "WARN"; Msg = "Tarea $tarea NO encontrada" }
        }
    }

    # --- Equipo renombrado ---
    $NombreEsperado = "PC-$NumInventario"
    if ($env:COMPUTERNAME -eq $NombreEsperado) {
        $Checks += @{ Status = "OK"; Msg = "Equipo renombrado: $NombreEsperado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "Equipo aun se llama '$env:COMPUTERNAME' (se renombrara a '$NombreEsperado' tras reinicio)" }
    }

    # --- Fondo de pantalla ---
    $FondoHCG = "C:\Users\Public\Pictures\wallpaper_hcg.jpg"
    if (Test-Path $FondoHCG) {
        $Checks += @{ Status = "OK"; Msg = "Fondo de pantalla configurado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "Fondo de pantalla NO encontrado" }
    }

    # --- Fondo de bloqueo ---
    $LockScreenHCG = "C:\Windows\Web\Wallpaper\HCG\LockScreen_$NumInventario.jpg"
    $LockScreenAlt = "C:\ProgramData\HCG\LockScreen\LockScreen_$NumInventario.jpg"
    if ((Test-Path $LockScreenHCG) -or (Test-Path $LockScreenAlt)) {
        $Checks += @{ Status = "OK"; Msg = "Fondo de pantalla de bloqueo generado" }
    } else {
        $Checks += @{ Status = "WARN"; Msg = "Fondo de pantalla de bloqueo NO encontrado" }
    }

    # --- Mostrar resultados ---
    foreach ($check in $Checks) {
        $Icono = switch ($check.Status) {
            "OK"   { "[OK]" }
            "WARN" { "[!!]" }
            "FAIL" { "[XX]" }
        }
        $Color = switch ($check.Status) {
            "OK"   { "Green" }
            "WARN" { "Yellow" }
            "FAIL" { "Red" }
        }
        Write-Host "  $Icono $($check.Msg)" -ForegroundColor $Color
    }

    $TotalOK = ($Checks | Where-Object { $_.Status -eq "OK" }).Count
    $TotalWarn = ($Checks | Where-Object { $_.Status -eq "WARN" }).Count
    $TotalFail = ($Checks | Where-Object { $_.Status -eq "FAIL" }).Count

    Write-Host ""
    Write-Host "  Resultado: $TotalOK OK, $TotalWarn advertencias, $TotalFail errores" -ForegroundColor $(if ($TotalFail -gt 0) { "Red" } elseif ($TotalWarn -gt 0) { "Yellow" } else { "Green" })
    Write-Host "  ======================================================" -ForegroundColor DarkYellow
    Write-Host ""
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Show-Banner

Write-Host "  $([char]0x2605) Ingresa el numero de inventario ST (5 digitos)" -ForegroundColor DarkYellow
$NumInventario = Read-Host "  Numero"

if ($NumInventario -notmatch '^\d{5}$') {
    Write-Host "`n  $([char]0x2716) [ERROR] Debe ser de 5 digitos" -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit
}

Show-CosmosAnimation -Message "Preparando configuracion cosmica..."
Write-Host ""
Write-Host "  $([char]0x2734) Iniciando configuracion para equipo: $NumInventario" -ForegroundColor Cyan
Write-Separator

# Obtener datos del equipo
$Datos = Get-DatosEquipo

# PASO 1: Registrar inicio
Send-DatosInicio -InvST $NumInventario -Datos $Datos
Play-StepSound

# PASO 2: Conectar al servidor
if (-not (Connect-Servidor)) {
    Write-Host "`n  $([char]0x2716) [ERROR] No se pudo conectar al servidor. Abortando." -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit
}
Play-StepSound

# PASOS 3-22: Instalacion y configuracion
Remove-OfficePrevio; Play-StepSound
New-UsuarioSoporte; Play-StepSound
New-UsuarioEquipo -NumInventario $NumInventario; Play-StepSound
Set-ImagenesUsuarios -NumInventario $NumInventario; Play-StepSound
Set-RedPrivada; Play-StepSound
Set-HoraAutomatica; Play-StepSound
Set-TemaOscuro -NumInventario $NumInventario; Play-StepSound
Install-WinRAR; Play-StepSound
Install-DotNet35; Play-StepSound
Install-AcrobatReader; Play-StepSound
Install-Chrome; Play-StepSound
Set-FondoPantalla; Play-StepSound
Set-LockScreenBackground -NumInventario $NumInventario; Play-StepSound
Install-Office; Play-StepSound
Install-Dedalus; Play-StepSound
Add-DedalusSyncStartup; Play-StepSound
Install-Antivirus; Play-StepSound
Copy-AccesosDirectos; Play-StepSound
Remove-DuplicateDesktopIcons; Play-StepSound
Enable-WindowsUpdate; Play-StepSound
Install-AllWindowsUpdates; Play-StepSound
Remove-AdminUsuarioActual; Play-StepSound
Install-ReporteIP; Play-StepSound
Install-ReporteSistema; Play-StepSound
Install-ReporteDiagnostico; Play-StepSound
Install-AutoUpdate; Play-StepSound

# Renombrar equipo
$NuevoNombre = "PC-$NumInventario"
try {
    Rename-Computer -NewName $NuevoNombre -Force -ErrorAction Stop
    Write-Log "Equipo renombrado a: $NuevoNombre" "OK"
} catch {
    Write-Log "No se pudo renombrar equipo a '$NuevoNombre': $($_.Exception.Message)" "WARN"
}

# Verificacion final
Verify-Configuracion -NumInventario $NumInventario

# PASO 24: Actualizar Google Sheets
Send-DatosFin -InvST $NumInventario
Play-StepSound

# PASO 25: Registrar inventario de software
Send-SoftwareInfo -InvST $NumInventario
Play-StepSound

# Mostrar resumen final cosmico
$Star = [char]0x2605
$Spark = [char]0x2734
$Arrow = [char]0x2192

# Animacion de celebracion
Show-CosmosAnimation -Message "Victoria cosmica alcanzada!"

# Melodia de victoria
Play-VictorySound

Show-ProgressCosmos -Step 27 -Total 27

Write-Host ""
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "        $Spark  CONFIGURACION COSMICA COMPLETADA  $Spark" -ForegroundColor Green
Write-Host "        $Spark  EL COSMO HA SIDO ENCENDIDO!       $Spark" -ForegroundColor Cyan
Write-Host ""
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "  $Spark DATOS DEL EQUIPO:" -ForegroundColor DarkYellow
Write-Host "  $Arrow Nombre:        $NuevoNombre" -ForegroundColor White
Write-Host "  $Arrow No. Serie:     $($Datos.Serie)" -ForegroundColor White
Write-Host "  $Arrow MAC Ethernet:  $($Datos.MACEthernet)" -ForegroundColor White
Write-Host "  $Arrow MAC WiFi:      $($Datos.MACWiFi)" -ForegroundColor White
Write-Host ""
Write-Host "  $Spark SOFTWARE INSTALADO:" -ForegroundColor DarkYellow
Write-Host "  $Arrow $($Script:SoftwareInstalado -join ', ')" -ForegroundColor Cyan
Write-Host ""
$NombreUsuarioFinal = if ($Script:EsOPD) { "OPD" } else { $NumInventario }
Write-Host "  $Spark SEGURIDAD:" -ForegroundColor DarkYellow
Write-Host "  $Arrow Usuario normal: $NombreUsuarioFinal (auto-login, estandar)" -ForegroundColor White
Write-Host "  $Arrow Usuario admin: $UsuarioSoporte (solo soporte tecnico)" -ForegroundColor White
Write-Host "  $Arrow Usuario '$($Script:UsuarioOriginal)' sin privilegios de admin" -ForegroundColor White
Write-Host ""
Write-Host "  $Spark REPORTES AUTOMATICOS:" -ForegroundColor DarkYellow
Write-Host "  $Arrow IP: cada 3 horas + al iniciar sesion" -ForegroundColor White
Write-Host "  $Arrow Sistema: impresoras, apps, usuarios, limpieza (al iniciar sesion)" -ForegroundColor White
Write-Host "  $Arrow Diagnostico: RAM, CPU, procesos, disco (cada 4 horas)" -ForegroundColor White
Write-Host ""
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "  $Spark IMPORTANTE: Reinicia el equipo para aplicar todos los cambios" -ForegroundColor Yellow
Write-Host ""
Write-Host "  $Star  Los Caballeros de Informatica protegen este equipo  $Star" -ForegroundColor Magenta
Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor DarkYellow
Write-Host ""

# PASO FINAL OPCIONAL: Actualizar a Windows 11
$WinVersion = [System.Environment]::OSVersion.Version
$WinBuild = $WinVersion.Build
$EsWindows11 = $WinBuild -ge 22000

if (-not $EsWindows11) {
    Write-Host ""
    Write-Host "  $Star $Star $Star ACTUALIZACION A WINDOWS 11 DISPONIBLE $Star $Star $Star" -ForegroundColor Cyan
    Write-Host ""
    $RespW11 = Read-Host "  $Star Deseas actualizar este equipo a Windows 11? (S/N)"

    if ($RespW11 -eq "S" -or $RespW11 -eq "s") {
        Write-Host ""
        Write-Host "  $Spark Preparando actualizacion cosmica a Windows 11..." -ForegroundColor Cyan

        # Ruta del ISO en el servidor
        $IsoServidor = "\\10.2.1.13\soportefaa\pack_installer_iA\windows_11_act\Win11_25H2_Spanish_Mexico_x64.iso"
        $IsoLocal = "$env:TEMP\Win11_Upgrade.iso"

        # Funcion para tocar melodia triste (Blue Dream - Saint Seiya)
        function Play-SadMelody {
            # Melodia triste inspirada en Blue Dream de Saint Seiya
            # Notas melancolicas con reverberacion emocional
            $notas = @(
                # Intro melancolico
                @{F=440; D=600; V=0.4},   # La4 - inicio suave
                @{F=392; D=600; V=0.45},  # Sol4
                @{F=349; D=800; V=0.5},   # Fa4 - sostenido
                @{F=330; D=400; V=0.45},  # Mi4
                @{F=294; D=600; V=0.4},   # Re4
                @{F=262; D=1000; V=0.5},  # Do4 - pausa emotiva

                # Secuencia nostalgica
                @{F=330; D=500; V=0.45},  # Mi4
                @{F=349; D=500; V=0.5},   # Fa4
                @{F=392; D=700; V=0.55},  # Sol4
                @{F=440; D=600; V=0.5},   # La4
                @{F=392; D=400; V=0.45},  # Sol4
                @{F=349; D=800; V=0.5},   # Fa4 - sostenido

                # Climax melancolico
                @{F=523; D=800; V=0.6},   # Do5 - punto alto
                @{F=494; D=500; V=0.55},  # Si4
                @{F=440; D=600; V=0.5},   # La4
                @{F=392; D=700; V=0.45},  # Sol4
                @{F=349; D=500; V=0.4},   # Fa4
                @{F=330; D=900; V=0.45},  # Mi4 - sostenido
                @{F=294; D=600; V=0.4},   # Re4
                @{F=262; D=1200; V=0.5},  # Do4 - final melancolico

                # Coda esperanzadora
                @{F=330; D=400; V=0.4},   # Mi4
                @{F=392; D=500; V=0.45},  # Sol4
                @{F=523; D=1000; V=0.55}  # Do5 - esperanza
            )

            foreach ($n in $notas) {
                try { [PegasusWavPlayer]::PlayTone($n.F, $n.D, $n.V) } catch { }
                Start-Sleep -Milliseconds 50
            }
        }

        # Funcion para mostrar progreso durante upgrade
        function Show-UpgradeProgress {
            param([string]$Message, [int]$Percent)
            $barWidth = 40
            $filled = [math]::Floor($barWidth * $Percent / 100)
            $empty = $barWidth - $filled
            $bar = ("$Star" * $filled) + ("-" * $empty)
            Write-Host "`r  [$bar] $Percent% - $Message" -NoNewline -ForegroundColor Cyan
        }

        try {
            # Verificar que el ISO existe en el servidor
            Write-Host "  $Arrow Verificando ISO en servidor..." -ForegroundColor Yellow
            if (-not (Test-Path $IsoServidor)) {
                Write-Host "  $Arrow ERROR: No se encontro el ISO en: $IsoServidor" -ForegroundColor Red
                throw "ISO no encontrado"
            }
            Write-Host "  $Arrow ISO encontrado! Copiando al equipo local..." -ForegroundColor Green

            # Copiar ISO al equipo local (con progreso)
            $IsoSize = (Get-Item $IsoServidor).Length
            $IsoSizeGB = [math]::Round($IsoSize / 1GB, 2)
            Write-Host "  $Arrow Tamano del ISO: $IsoSizeGB GB" -ForegroundColor Cyan
            Write-Host "  $Arrow Copiando ISO... (esto tomara unos minutos)" -ForegroundColor Yellow
            Write-Host ""

            # Iniciar melodia ambiental durante la copia
            Start-BackgroundMelody -Mensaje "Melodia del Santuario - Copiando Windows 11..."

            # Copiar ISO
            Copy-Item -Path $IsoServidor -Destination $IsoLocal -Force

            # Detener melodia ambiental
            Stop-BackgroundMelody

            Write-Host ""
            Write-Host "  $Arrow ISO copiado exitosamente!" -ForegroundColor Green

            # Montar ISO
            Write-Host "  $Arrow Montando imagen ISO..." -ForegroundColor Yellow
            $MountResult = Mount-DiskImage -ImagePath $IsoLocal -PassThru
            $DriveLetter = ($MountResult | Get-Volume).DriveLetter

            if (-not $DriveLetter) {
                throw "No se pudo montar el ISO"
            }

            Write-Host "  $Arrow ISO montado en unidad: ${DriveLetter}:" -ForegroundColor Green

            # Ruta del setup
            $SetupPath = "${DriveLetter}:\setup.exe"

            if (-not (Test-Path $SetupPath)) {
                throw "No se encontro setup.exe en el ISO"
            }

            # Mostrar mensaje epico antes del upgrade
            Write-Host ""
            Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor Magenta
            Write-Host ""
            Write-Host "         $Spark  INICIANDO TRANSFORMACION COSMICA  $Spark" -ForegroundColor Cyan
            Write-Host "         $Spark  WINDOWS 11 - EL NUEVO COSMOS      $Spark" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star $Star" -ForegroundColor Magenta
            Write-Host ""
            Write-Host "  $Arrow El proceso de actualizacion comenzara ahora..." -ForegroundColor Yellow
            Write-Host "  $Arrow El equipo se reiniciara automaticamente varias veces" -ForegroundColor Yellow
            Write-Host "  $Arrow NO apagues el equipo durante la actualizacion!" -ForegroundColor Red
            Write-Host ""

            # Tocar melodia triste antes del upgrade
            Write-Host "  $Spark Melodia de despedida a Windows 10..." -ForegroundColor DarkCyan
            Play-SadMelody

            Start-Sleep -Seconds 2

            # Ejecutar upgrade silencioso
            Write-Host ""
            Write-Host "  $Arrow Ejecutando Windows 11 Setup..." -ForegroundColor Green

            # Parametros para upgrade silencioso
            # /auto upgrade = actualizacion automatica
            # /quiet = sin interaccion del usuario
            # /eula accept = aceptar EULA
            # /DynamicUpdate disable = no descargar actualizaciones durante setup
            $SetupArgs = "/auto upgrade /quiet /eula accept /DynamicUpdate disable /copylogs $env:TEMP\Win11UpgradeLogs"

            Start-Process -FilePath $SetupPath -ArgumentList $SetupArgs -Wait

            # Si llegamos aqui, el upgrade se inicio (puede reiniciar)
            Write-Host ""
            Write-Host "  $Star $Star $Star ACTUALIZACION INICIADA EXITOSAMENTE $Star $Star $Star" -ForegroundColor Green
            Write-Host "  $Arrow El equipo se reiniciara para completar la actualizacion" -ForegroundColor Cyan

        } catch {
            Write-Host ""
            Write-Host "  $Arrow Error durante actualizacion: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  $Arrow La configuracion basica se completo correctamente" -ForegroundColor Yellow

            # Limpiar ISO local si existe
            if (Test-Path $IsoLocal) {
                try {
                    Dismount-DiskImage -ImagePath $IsoLocal -ErrorAction SilentlyContinue
                    Remove-Item $IsoLocal -Force -ErrorAction SilentlyContinue
                } catch { }
            }
        }
    }
} else {
    Write-Host ""
    Write-Host "  $Star Este equipo ya tiene Windows 11 (Build $WinBuild)" -ForegroundColor Green
}

# =============================================================================
# REINICIO AUTOMATICO - AVE FENIX
# =============================================================================
Write-Host ""
Write-Host "  $Star Preparando reinicio automatico para aplicar actualizaciones..." -ForegroundColor Yellow
Start-Sleep -Seconds 3

# Mostrar cuenta regresiva del Fenix (60 segundos) y reiniciar
Show-PhoenixRebootCountdown -Seconds 60
Restart-Computer -Force
