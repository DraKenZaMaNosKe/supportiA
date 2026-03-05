# Script temporal para leer netlogon6.bat
$ErrorActionPreference = "Stop"

try {
    # Conectar al servidor
    $result = net use \\10.2.1.17\distribucion /user:distribucion distribucion /persistent:no 2>&1
    Write-Host "Conexion: $result"

    # Listar archivos .bat
    Write-Host "`n=== ARCHIVOS .BAT EN SINCRONIZADOR ===" -ForegroundColor Cyan
    Get-ChildItem "\\10.2.1.17\distribucion\dedalus\sincronizador\*.bat" | ForEach-Object {
        Write-Host $_.Name -ForegroundColor Yellow
    }

    # Leer netlogon6.bat
    Write-Host "`n=== CONTENIDO DE netlogon6.bat ===" -ForegroundColor Green
    $content = Get-Content "\\10.2.1.17\distribucion\dedalus\sincronizador\netlogon6.bat" -Raw
    Write-Host $content

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Desconectar
    net use \\10.2.1.17\distribucion /delete /y 2>$null
}
