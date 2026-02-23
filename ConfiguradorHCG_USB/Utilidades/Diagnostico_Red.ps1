# =============================================================================
# DIAGNOSTICO DE RED - CONEXION A GOOGLE SHEETS
# =============================================================================

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "  DIAGNOSTICO DE CONEXION A GOOGLE" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""

# 1. Verificar conexion a internet basica
Write-Host "  [1] Verificando conexion a internet..." -ForegroundColor Yellow
$Ping = Test-Connection -ComputerName 8.8.8.8 -Count 2 -Quiet
if ($Ping) {
    Write-Host "      [OK] Hay conexion a internet (ping a 8.8.8.8)" -ForegroundColor Green
} else {
    Write-Host "      [ERROR] No hay conexion a internet" -ForegroundColor Red
}

# 2. Verificar DNS
Write-Host ""
Write-Host "  [2] Verificando resolucion DNS..." -ForegroundColor Yellow
try {
    $DNS = Resolve-DnsName -Name "script.google.com" -ErrorAction Stop
    Write-Host "      [OK] DNS resuelve script.google.com" -ForegroundColor Green
    Write-Host "      IP: $($DNS[0].IPAddress)" -ForegroundColor Gray
} catch {
    Write-Host "      [ERROR] No puede resolver script.google.com" -ForegroundColor Red
}

# 3. Verificar proxy
Write-Host ""
Write-Host "  [3] Verificando configuracion de proxy..." -ForegroundColor Yellow
$Proxy = [System.Net.WebRequest]::GetSystemWebProxy()
$ProxyUri = $Proxy.GetProxy("https://script.google.com")
if ($ProxyUri.Host -ne "script.google.com") {
    Write-Host "      [INFO] Proxy detectado: $ProxyUri" -ForegroundColor Yellow
} else {
    Write-Host "      [OK] No hay proxy configurado" -ForegroundColor Green
}

# Verificar proxy en registro
$RegProxy = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -ErrorAction SilentlyContinue
if ($RegProxy.ProxyEnable -eq 1) {
    Write-Host "      [INFO] Proxy en registro: $($RegProxy.ProxyServer)" -ForegroundColor Yellow
}

# 4. Verificar TLS
Write-Host ""
Write-Host "  [4] Verificando protocolos TLS..." -ForegroundColor Yellow
$TLS = [Net.ServicePointManager]::SecurityProtocol
Write-Host "      Protocolo actual: $TLS" -ForegroundColor Gray

# Forzar TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Host "      [OK] TLS 1.2 habilitado" -ForegroundColor Green

# 5. Probar conexion HTTPS a Google
Write-Host ""
Write-Host "  [5] Probando conexion HTTPS a Google..." -ForegroundColor Yellow
try {
    $Response = Invoke-WebRequest -Uri "https://www.google.com" -UseBasicParsing -TimeoutSec 10
    Write-Host "      [OK] Conexion a google.com exitosa (Status: $($Response.StatusCode))" -ForegroundColor Green
} catch {
    Write-Host "      [ERROR] No puede conectar a google.com: $($_.Exception.Message)" -ForegroundColor Red
}

# 6. Probar conexion al Apps Script
Write-Host ""
Write-Host "  [6] Probando conexion al Apps Script..." -ForegroundColor Yellow
$URL = "https://script.google.com/macros/s/AKfycbwDAZERvkP6V8sZczvpJgh6LvoXOqwuWymCauicX5dyYGRh5Iym4J5czjHg3tDHbPtP/exec"

try {
    # Primero solo GET para ver si responde
    $TestResponse = Invoke-WebRequest -Uri $URL -UseBasicParsing -TimeoutSec 30 -Method Get
    Write-Host "      [OK] Apps Script responde (Status: $($TestResponse.StatusCode))" -ForegroundColor Green
} catch {
    Write-Host "      [ERROR] No puede conectar al Apps Script" -ForegroundColor Red
    Write-Host "      Detalle: $($_.Exception.Message)" -ForegroundColor Gray
}

# 7. Probar envio POST
Write-Host ""
Write-Host "  [7] Probando envio de datos (POST)..." -ForegroundColor Yellow
try {
    $TestData = @{ Test = "Diagnostico"; Fecha = (Get-Date -Format "dd/MM/yyyy HH:mm:ss") } | ConvertTo-Json

    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $PostResponse = Invoke-RestMethod -Uri $URL -Method Post -Body $TestData -ContentType "application/json" -TimeoutSec 60
    $Stopwatch.Stop()

    Write-Host "      [OK] POST exitoso en $($Stopwatch.ElapsedMilliseconds)ms" -ForegroundColor Green
    Write-Host "      Respuesta: $($PostResponse | ConvertTo-Json -Compress)" -ForegroundColor Gray
} catch {
    Write-Host "      [ERROR] POST fallido" -ForegroundColor Red
    Write-Host "      Detalle: $($_.Exception.Message)" -ForegroundColor Gray

    if ($_.Exception.Message -like "*timeout*" -or $_.Exception.Message -like "*tiempo*") {
        Write-Host ""
        Write-Host "      CAUSA PROBABLE: Timeout" -ForegroundColor Yellow
        Write-Host "      La red del hospital puede estar bloqueando o ralentizando" -ForegroundColor Yellow
        Write-Host "      las conexiones a servidores de Google." -ForegroundColor Yellow
    }
}

# 8. Informacion de red
Write-Host ""
Write-Host "  [8] Informacion de red..." -ForegroundColor Yellow
$Adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
foreach ($Adapter in $Adapters) {
    Write-Host "      Adaptador: $($Adapter.Name) - $($Adapter.InterfaceDescription)" -ForegroundColor Gray
    $IP = Get-NetIPAddress -InterfaceIndex $Adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
    if ($IP) {
        Write-Host "      IP: $($IP.IPAddress)" -ForegroundColor Gray
    }
}

# Resumen
Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "  RESUMEN" -ForegroundColor Cyan
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Si el paso 7 falla con TIMEOUT, las opciones son:" -ForegroundColor Yellow
Write-Host "  1. Usar el sistema de respaldo (JSON local + EnviarPendientes.ps1)" -ForegroundColor White
Write-Host "  2. Pedir a TI que permita conexiones a *.google.com" -ForegroundColor White
Write-Host "  3. Configurar proxy si es necesario" -ForegroundColor White
Write-Host ""

Read-Host "  Presiona Enter para salir"
