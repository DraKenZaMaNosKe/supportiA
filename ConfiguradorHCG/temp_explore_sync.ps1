# Script para explorar archivos del sincronizador
$ErrorActionPreference = "SilentlyContinue"

# Conectar al servidor
net use \\10.2.1.17\distribucion /user:distribucion distribucion /persistent:no 2>$null

$basePath = "\\10.2.1.17\distribucion\dedalus\sincronizador"

Write-Host "=== ESTRUCTURA DE CARPETAS ===" -ForegroundColor Cyan
Get-ChildItem $basePath -Directory | ForEach-Object {
    Write-Host "  [DIR] $($_.Name)" -ForegroundColor Yellow
}

Write-Host "`n=== ARCHIVOS .EXE EN SINCRONIZADOR ===" -ForegroundColor Cyan
Get-ChildItem $basePath -Filter "*.exe" -Recurse | ForEach-Object {
    Write-Host "  $($_.FullName)" -ForegroundColor Green
}

Write-Host "`n=== CONTENIDO DE CARPETA CreateSynch ===" -ForegroundColor Cyan
Get-ChildItem "$basePath\CreateSynch" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "  $($_.Name)" -ForegroundColor Yellow
}

Write-Host "`n=== BUSCANDO REFERENCIAS A .EXE EN sync_xhis6.bat ===" -ForegroundColor Cyan
$content = Get-Content "\\10.2.1.17\distribucion\dedalus\sincronizador\sync_xhis6.bat" -Raw
$matches = [regex]::Matches($content, '[^\s"]+\.exe[^\s"]*|"[^"]+\.exe[^"]*"')
foreach ($match in $matches) {
    Write-Host "  $($match.Value)" -ForegroundColor Magenta
}

Write-Host "`n=== CONTENIDO COMPLETO DE sync_xhis6.bat ===" -ForegroundColor Cyan
Write-Host $content

# Desconectar
net use \\10.2.1.17\distribucion /delete /y 2>$null
