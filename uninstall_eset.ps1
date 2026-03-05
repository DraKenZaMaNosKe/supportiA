# Script para desinstalar ESET e instalar la nueva version

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  DESINSTALADOR DE ESET" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Buscar productos ESET instalados
Write-Host "Buscando productos ESET instalados..." -ForegroundColor Yellow
$esetProducts = Get-WmiObject Win32_Product | Where-Object { $_.Name -like '*ESET*' }

if ($esetProducts) {
    Write-Host "Productos ESET encontrados:" -ForegroundColor Green
    foreach ($product in $esetProducts) {
        Write-Host "  - $($product.Name)" -ForegroundColor White
    }
    Write-Host ""

    # Desinstalar cada producto
    foreach ($product in $esetProducts) {
        Write-Host "Desinstalando: $($product.Name)..." -ForegroundColor Yellow
        try {
            $result = $product.Uninstall()
            if ($result.ReturnValue -eq 0) {
                Write-Host "  OK - Desinstalado correctamente" -ForegroundColor Green
            } else {
                Write-Host "  Codigo de retorno: $($result.ReturnValue)" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "  Error: $_" -ForegroundColor Red
        }
    }
} else {
    Write-Host "No se encontraron productos ESET mediante WMI" -ForegroundColor Yellow
    Write-Host "Intentando metodo alternativo..." -ForegroundColor Yellow
}

# Metodo alternativo: buscar en registro y usar uninstaller
Write-Host ""
Write-Host "Buscando en el registro de Windows..." -ForegroundColor Yellow

$uninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $uninstallPaths) {
    $esetApps = Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like '*ESET*' }
    foreach ($app in $esetApps) {
        if ($app.UninstallString) {
            Write-Host "Encontrado: $($app.DisplayName)" -ForegroundColor Cyan
            Write-Host "  Comando: $($app.UninstallString)" -ForegroundColor Gray
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Proceso completado." -ForegroundColor Green
Write-Host ""
Write-Host "Si ESET sigue instalado, usa el desinstalador oficial:" -ForegroundColor Yellow
Write-Host "  Panel de Control > Programas > Desinstalar ESET" -ForegroundColor White
Write-Host ""
Write-Host "Despues de desinstalar, reinicia y ejecuta:" -ForegroundColor Yellow
Write-Host '  \\10.2.1.13\soportefaa\pack_installer_iA\antivirus\PROTECT_Installer_x64_es_CL 2.exe' -ForegroundColor White
