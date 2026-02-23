# =============================================================================
# GUARDAR CREDENCIALES - HOSPITAL CIVIL FAA
# =============================================================================
# Ejecuta este script UNA VEZ en cada equipo para guardar las credenciales
# de red de forma segura en Windows Credential Manager
# =============================================================================

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║        GUARDAR CREDENCIALES DE RED - HCG                     ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Servidor 1: Software (10.2.1.13)
Write-Host "  Configurando credenciales para servidor de Software (10.2.1.13)" -ForegroundColor Yellow
Write-Host ""

$Usuario = Read-Host "  Usuario"
$Password = Read-Host "  Contrasena" -AsSecureString

# Convertir SecureString a texto para cmdkey
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
$PasswordPlano = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Guardar credenciales para ambos servidores
Write-Host ""
Write-Host "  Guardando credenciales..." -ForegroundColor Gray

# Servidor de Software
cmdkey /add:10.2.1.13 /user:$Usuario /pass:$PasswordPlano | Out-Null
Write-Host "  [OK] Credenciales guardadas para 10.2.1.13" -ForegroundColor Green

# Servidor de Dedalus
cmdkey /add:10.2.1.17 /user:$Usuario /pass:$PasswordPlano | Out-Null
Write-Host "  [OK] Credenciales guardadas para 10.2.1.17" -ForegroundColor Green

# Limpiar password de memoria
$PasswordPlano = $null
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

Write-Host ""
Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Credenciales guardadas en Windows Credential Manager" -ForegroundColor Green
Write-Host "  Ya no necesitas ingresarlas cada vez que ejecutes el script" -ForegroundColor Gray
Write-Host "  ════════════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""

# Verificar conexion
Write-Host "  Verificando conexion a servidores..." -ForegroundColor Yellow

$Test1 = Test-Path "\\10.2.1.13\soportefaa" -ErrorAction SilentlyContinue
$Test2 = Test-Path "\\10.2.1.17\distribucion" -ErrorAction SilentlyContinue

if ($Test1) {
    Write-Host "  [OK] Conexion a \\10.2.1.13\soportefaa exitosa" -ForegroundColor Green
} else {
    Write-Host "  [X] No se pudo conectar a \\10.2.1.13\soportefaa" -ForegroundColor Red
}

if ($Test2) {
    Write-Host "  [OK] Conexion a \\10.2.1.17\distribucion exitosa" -ForegroundColor Green
} else {
    Write-Host "  [X] No se pudo conectar a \\10.2.1.17\distribucion" -ForegroundColor Red
}

Write-Host ""
Read-Host "  Presiona Enter para salir"
