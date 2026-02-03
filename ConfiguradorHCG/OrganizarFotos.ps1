# =============================================================================
# ORGANIZADOR DE FOTOS POR INVENTARIO
# =============================================================================
# Copia fotos desde una carpeta de entrada y las organiza por numero de inventario
# =============================================================================

$RutaFotosEntrada = "C:\HCG_Fotos\Entrada"
$RutaFotosOrganizadas = "C:\HCG_Fotos\Equipos"

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "       ORGANIZADOR DE FOTOS POR INVENTARIO" -ForegroundColor Cyan
    Write-Host "       Hospital Civil de Guadalajara - FAA" -ForegroundColor Cyan
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Initialize-Carpetas {
    if (-not (Test-Path $RutaFotosEntrada)) {
        New-Item -ItemType Directory -Path $RutaFotosEntrada -Force | Out-Null
        Write-Host "  [INFO] Carpeta de entrada creada: $RutaFotosEntrada" -ForegroundColor Yellow
    }
    if (-not (Test-Path $RutaFotosOrganizadas)) {
        New-Item -ItemType Directory -Path $RutaFotosOrganizadas -Force | Out-Null
        Write-Host "  [INFO] Carpeta de salida creada: $RutaFotosOrganizadas" -ForegroundColor Yellow
    }
}

function Get-FotosDisponibles {
    $Extensiones = @("*.jpg", "*.jpeg", "*.png", "*.heic", "*.bmp")
    $Fotos = @()
    foreach ($Ext in $Extensiones) {
        $Fotos += Get-ChildItem -Path $RutaFotosEntrada -Filter $Ext -ErrorAction SilentlyContinue
    }
    return $Fotos | Sort-Object LastWriteTime
}

function Move-FotoAEquipo {
    param($Foto, $InvST, $TipoFoto)

    $CarpetaEquipo = Join-Path $RutaFotosOrganizadas "PC-$InvST"

    if (-not (Test-Path $CarpetaEquipo)) {
        New-Item -ItemType Directory -Path $CarpetaEquipo -Force | Out-Null
    }

    # Determinar nombre segun tipo
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $Extension = [System.IO.Path]::GetExtension($Foto.Name)

    $NombreNuevo = switch ($TipoFoto) {
        "1" { "Caja_$Timestamp$Extension" }
        "2" { "Etiqueta_Serie_$Timestamp$Extension" }
        "3" { "Etiqueta_Specs_$Timestamp$Extension" }
        "4" { "Garantia_$Timestamp$Extension" }
        "5" { "Otra_$Timestamp$Extension" }
        default { "$($Foto.BaseName)_$Timestamp$Extension" }
    }

    $RutaDestino = Join-Path $CarpetaEquipo $NombreNuevo
    Move-Item -Path $Foto.FullName -Destination $RutaDestino -Force

    return $RutaDestino
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Show-Banner
Initialize-Carpetas

Write-Host "  Carpeta de entrada:  $RutaFotosEntrada" -ForegroundColor Gray
Write-Host "  Carpeta de salida:   $RutaFotosOrganizadas" -ForegroundColor Gray
Write-Host ""

# Verificar fotos disponibles
$Fotos = Get-FotosDisponibles

if ($Fotos.Count -eq 0) {
    Write-Host "  [INFO] No hay fotos en la carpeta de entrada." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Instrucciones:" -ForegroundColor Cyan
    Write-Host "  1. Copia las fotos del celular a: $RutaFotosEntrada" -ForegroundColor White
    Write-Host "  2. Ejecuta este script nuevamente" -ForegroundColor White
    Write-Host "  3. Asigna cada foto a su numero de inventario" -ForegroundColor White
    Write-Host ""
    Write-Host "  Tip: Puedes conectar el celular por USB o usar:" -ForegroundColor Gray
    Write-Host "       - Google Photos (descarga)" -ForegroundColor Gray
    Write-Host "       - OneDrive" -ForegroundColor Gray
    Write-Host "       - Compartir por WiFi (AirDrop, Nearby Share)" -ForegroundColor Gray
    Write-Host ""

    $AbrirCarpeta = Read-Host "  Abrir carpeta de entrada? (S/N)"
    if ($AbrirCarpeta -eq "S" -or $AbrirCarpeta -eq "s") {
        Start-Process explorer.exe -ArgumentList $RutaFotosEntrada
    }

    Read-Host "`n  Presiona Enter para salir"
    exit
}

Write-Host "  Fotos encontradas: $($Fotos.Count)" -ForegroundColor Green
Write-Host ""

# Procesar cada foto
$FotoActual = 0
foreach ($Foto in $Fotos) {
    $FotoActual++

    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "  FOTO $FotoActual de $($Fotos.Count): $($Foto.Name)" -ForegroundColor Cyan
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "  Fecha: $($Foto.LastWriteTime)" -ForegroundColor Gray
    Write-Host "  Tamano: $([math]::Round($Foto.Length / 1KB, 0)) KB" -ForegroundColor Gray
    Write-Host ""

    # Abrir la foto para verla
    Write-Host "  Abriendo foto para visualizar..." -ForegroundColor Yellow
    Start-Process -FilePath $Foto.FullName

    Write-Host ""
    Write-Host "  Ingresa el numero de inventario para esta foto (5 digitos)" -ForegroundColor Yellow
    Write-Host "  O escribe 'S' para saltar, 'X' para terminar" -ForegroundColor Gray
    $Input = Read-Host "  Inventario"

    if ($Input -eq "X" -or $Input -eq "x") {
        Write-Host "`n  Proceso cancelado por el usuario." -ForegroundColor Yellow
        break
    }

    if ($Input -eq "S" -or $Input -eq "s") {
        Write-Host "  [INFO] Foto saltada." -ForegroundColor Gray
        continue
    }

    if ($Input -notmatch '^\d{5}$') {
        Write-Host "  [WARN] Numero invalido. Foto saltada." -ForegroundColor Yellow
        continue
    }

    $NumInventario = $Input

    # Preguntar tipo de foto
    Write-Host ""
    Write-Host "  Tipo de foto:" -ForegroundColor Cyan
    Write-Host "  1. Foto de la caja completa" -ForegroundColor White
    Write-Host "  2. Etiqueta con numero de serie" -ForegroundColor White
    Write-Host "  3. Etiqueta con especificaciones (RAM, disco, etc.)" -ForegroundColor White
    Write-Host "  4. Etiqueta de garantia" -ForegroundColor White
    Write-Host "  5. Otra" -ForegroundColor White
    $TipoFoto = Read-Host "  Selecciona (1-5)"

    # Mover foto
    $RutaFinal = Move-FotoAEquipo -Foto $Foto -InvST $NumInventario -TipoFoto $TipoFoto
    Write-Host "  [OK] Foto guardada en: $RutaFinal" -ForegroundColor Green
    Write-Host ""
}

# Resumen final
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  PROCESO COMPLETADO" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Las fotos fueron organizadas en: $RutaFotosOrganizadas" -ForegroundColor White
Write-Host ""

# Mostrar equipos con fotos
$Carpetas = Get-ChildItem -Path $RutaFotosOrganizadas -Directory -ErrorAction SilentlyContinue
if ($Carpetas) {
    Write-Host "  Equipos con fotos:" -ForegroundColor Cyan
    foreach ($Carpeta in $Carpetas) {
        $NumFotos = (Get-ChildItem -Path $Carpeta.FullName -File).Count
        Write-Host "    $($Carpeta.Name): $NumFotos foto(s)" -ForegroundColor White
    }
}

Write-Host ""
$AbrirCarpeta = Read-Host "  Abrir carpeta de equipos? (S/N)"
if ($AbrirCarpeta -eq "S" -or $AbrirCarpeta -eq "s") {
    Start-Process explorer.exe -ArgumentList $RutaFotosOrganizadas
}

Write-Host ""
Read-Host "  Presiona Enter para salir"
