# =============================================================================
# HCG - INSTALAR PROTECTOR DE PANTALLA MATRIX
# =============================================================================
# Compila el .scr, lo copia a System32, y lo activa
# Se eleva automaticamente como Administrador
# =============================================================================

# --- Auto-elevacion como Administrador ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  HCG - INSTALAR PROTECTOR DE PANTALLA MATRIX" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""

$ScriptDir = Split-Path -Parent $PSCommandPath

# --- 1. Buscar el compilador de C# ---
Write-Host "  [1/4] Buscando compilador C#..." -ForegroundColor Yellow

$cscPath = $null
$fwPaths = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
foreach ($p in $fwPaths) {
    if (Test-Path $p) { $cscPath = $p; break }
}

if (-not $cscPath) {
    Write-Host "  [ERROR] No se encontro csc.exe" -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit 1
}
Write-Host "         $cscPath" -ForegroundColor Gray

# --- 2. Compilar el .scr ---
Write-Host "  [2/4] Compilando HCG Matrix.scr..." -ForegroundColor Yellow

$csSource = "$ScriptDir\OOBE\HCG_Matrix_Screensaver.cs"
if (-not (Test-Path $csSource)) {
    # Si no esta en OOBE, buscar junto al script
    $csSource = "$ScriptDir\HCG_Matrix_Screensaver.cs"
}

$scrOutput = "$env:TEMP\HCG Matrix.scr"

if (-not (Test-Path $csSource)) {
    Write-Host "  [ERROR] No se encontro HCG_Matrix_Screensaver.cs" -ForegroundColor Red
    Write-Host "         Buscado en: $ScriptDir\OOBE\" -ForegroundColor Gray
    Read-Host "  Presiona Enter para salir"
    exit 1
}

$compileResult = & $cscPath /target:winexe /out:"$scrOutput" /reference:System.Windows.Forms.dll "$csSource" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [ERROR] Fallo la compilacion:" -ForegroundColor Red
    Write-Host "  $compileResult" -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit 1
}
Write-Host "         Compilado OK" -ForegroundColor Gray

# --- 3. Copiar archivos a System32 ---
Write-Host "  [3/4] Instalando en System32..." -ForegroundColor Yellow

# Copiar el .scr
Copy-Item -Path $scrOutput -Destination "$env:WINDIR\System32\HCG Matrix.scr" -Force
Write-Host "         HCG Matrix.scr -> System32" -ForegroundColor Gray

# Copiar el script PowerShell del Matrix
$matrixScript = "$ScriptDir\ScreenSaver_Matrix.ps1"
if (-not (Test-Path $matrixScript)) {
    Write-Host "  [ERROR] No se encontro ScreenSaver_Matrix.ps1" -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit 1
}
Copy-Item -Path $matrixScript -Destination "$env:WINDIR\System32\HCG_Matrix.ps1" -Force
Write-Host "         HCG_Matrix.ps1 -> System32" -ForegroundColor Gray

# --- 4. Activar como protector de pantalla (opcional) ---
Write-Host "  [4/4] Configurando como protector de pantalla activo..." -ForegroundColor Yellow

# Establecer como screensaver activo con 5 minutos de espera
$regPath = "HKCU:\Control Panel\Desktop"
Set-ItemProperty -Path $regPath -Name "SCRNSAVE.EXE" -Value "$env:WINDIR\System32\HCG Matrix.scr"
Set-ItemProperty -Path $regPath -Name "ScreenSaveActive" -Value "1"
Set-ItemProperty -Path $regPath -Name "ScreenSaveTimeOut" -Value "300"

Write-Host "         Activado (5 minutos de inactividad)" -ForegroundColor Gray

# --- Limpiar temp ---
Remove-Item $scrOutput -Force -ErrorAction SilentlyContinue

# --- Resumen ---
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  INSTALADO CORRECTAMENTE" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Archivos instalados:" -ForegroundColor White
Write-Host "    $env:WINDIR\System32\HCG Matrix.scr" -ForegroundColor Gray
Write-Host "    $env:WINDIR\System32\HCG_Matrix.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "  Ya aparece en:" -ForegroundColor White
Write-Host "    Configuracion > Personalizacion > Protector de pantalla" -ForegroundColor Gray
Write-Host "    O: clic derecho escritorio > Personalizar > Protector de pantalla" -ForegroundColor Gray
Write-Host ""
Write-Host "  Se activa despues de 5 minutos de inactividad." -ForegroundColor Yellow
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""

Read-Host "  Presiona Enter para salir"
