# Verificar si estamos como admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-Host "Ejecutando como Admin: $isAdmin" -ForegroundColor $(if ($isAdmin) { "Green" } else { "Red" })

# Verificar si hay procesos ESET corriendo
$esetProcs = Get-Process -Name "*eset*","*ekrn*","*egui*" -ErrorAction SilentlyContinue
if ($esetProcs) {
    Write-Host "Procesos ESET encontrados:" -ForegroundColor Yellow
    $esetProcs | ForEach-Object { Write-Host "  - $($_.Name) (PID: $($_.Id))" }
} else {
    Write-Host "No hay procesos ESET corriendo" -ForegroundColor Yellow
}

# Verificar si algo se instalo parcialmente
$paths = @(
    "C:\Program Files\ESET",
    "C:\ProgramData\ESET"
)
foreach ($p in $paths) {
    if (Test-Path $p) {
        Write-Host "Carpeta encontrada: $p" -ForegroundColor Cyan
        Get-ChildItem $p | ForEach-Object { Write-Host "  - $($_.Name)" }
    }
}

# Ver si el instalador tiene ayuda
Write-Host ""
Write-Host "Probando parametros del instalador..." -ForegroundColor Yellow
$instalador = "\\10.2.1.13\soportefaa\ANTIVIRUS\ESET 2025\PROTECT_Installer_x64_es_CL 2.exe"
$p = Start-Process -FilePath $instalador -ArgumentList "/?" -Wait -NoNewWindow -PassThru
Write-Host "Exit code con /?: $($p.ExitCode)"
