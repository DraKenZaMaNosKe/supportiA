# Test de conexion a servidores de red

Write-Host ""
Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  PRUEBA DE CONEXION A SERVIDORES - HCG" -ForegroundColor Cyan
Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Pedir credenciales
Write-Host "  Ingresa tus credenciales de red:" -ForegroundColor Yellow
$Usuario = Read-Host "  Usuario"
$Password = Read-Host "  Password"

Write-Host ""
Write-Host "  Conectando a servidores..." -ForegroundColor Gray

# Desconectar conexiones previas
net use \\10.2.1.13 /delete /y 2>$null | Out-Null
net use \\10.2.1.17 /delete /y 2>$null | Out-Null

# Conectar a servidor de software
Write-Host ""
Write-Host "  [1] Servidor de Software (10.2.1.13)" -ForegroundColor Cyan
$Resultado1 = net use \\10.2.1.13\soportefaa /user:$Usuario $Password 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "      [OK] Conectado" -ForegroundColor Green

    # Listar carpeta de software
    Write-Host ""
    Write-Host "      Contenido de: \\10.2.1.13\soportefaa\Software Equipo Nuevo\Programas" -ForegroundColor Yellow
    if (Test-Path "\\10.2.1.13\soportefaa\Software Equipo Nuevo\Programas") {
        Get-ChildItem "\\10.2.1.13\soportefaa\Software Equipo Nuevo\Programas" | ForEach-Object {
            Write-Host "        - $($_.Name)" -ForegroundColor White
        }
    }

    # Verificar Office 2007
    Write-Host ""
    Write-Host "      Contenido de: ...\Ofice2007" -ForegroundColor Yellow
    if (Test-Path "\\10.2.1.13\soportefaa\Software Equipo Nuevo\Programas\Ofice2007") {
        Get-ChildItem "\\10.2.1.13\soportefaa\Software Equipo Nuevo\Programas\Ofice2007" | ForEach-Object {
            Write-Host "        - $($_.Name)" -ForegroundColor White
        }
    } else {
        Write-Host "        [!] Carpeta no encontrada" -ForegroundColor Red
    }

    # Verificar Antivirus
    Write-Host ""
    Write-Host "      Contenido de: ...\ANTIVIRUS\ESET 2025" -ForegroundColor Yellow
    if (Test-Path "\\10.2.1.13\soportefaa\ANTIVIRUS\ESET 2025") {
        Get-ChildItem "\\10.2.1.13\soportefaa\ANTIVIRUS\ESET 2025" | ForEach-Object {
            Write-Host "        - $($_.Name)" -ForegroundColor White
        }
    } else {
        Write-Host "        [!] Carpeta no encontrada" -ForegroundColor Red
    }

    # Verificar Acrobat
    Write-Host ""
    Write-Host "      Contenido de: ...\AcroRdrDC2100120155_es_ES" -ForegroundColor Yellow
    $RutaAcrobat = "\\10.2.1.13\soportefaa\Software Equipo Nuevo\Programas AcroRdrDC2100120155_es_ES"
    if (Test-Path $RutaAcrobat) {
        Get-ChildItem $RutaAcrobat -File | Select-Object -First 5 | ForEach-Object {
            Write-Host "        - $($_.Name)" -ForegroundColor White
        }
    } else {
        Write-Host "        [!] Carpeta no encontrada" -ForegroundColor Red
    }

} else {
    Write-Host "      [ERROR] No se pudo conectar" -ForegroundColor Red
    Write-Host "      $Resultado1" -ForegroundColor Gray
}

# Conectar a servidor de Dedalus
Write-Host ""
Write-Host "  [2] Servidor Dedalus (10.2.1.17)" -ForegroundColor Cyan
$Resultado2 = net use \\10.2.1.17\distribucion /user:distribucion distribucion 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "      [OK] Conectado" -ForegroundColor Green

    # Listar carpeta de Dedalus
    Write-Host ""
    Write-Host "      Contenido de: \\10.2.1.17\distribucion\dedalus\sincronizador" -ForegroundColor Yellow
    if (Test-Path "\\10.2.1.17\distribucion\dedalus\sincronizador") {
        Get-ChildItem "\\10.2.1.17\distribucion\dedalus\sincronizador" | Select-Object -First 10 | ForEach-Object {
            Write-Host "        - $($_.Name)" -ForegroundColor White
        }
    } else {
        Write-Host "        [!] Carpeta no encontrada" -ForegroundColor Red
    }
} else {
    Write-Host "      [ERROR] No se pudo conectar" -ForegroundColor Red
    Write-Host "      $Resultado2" -ForegroundColor Gray
}

Write-Host ""
Write-Host "  ══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
