# =============================================================================
# VERIFICACION DE DATOS - FOTOS vs SISTEMA
# =============================================================================
# Genera un reporte detallado del equipo para comparar con las fotos
# =============================================================================

#Requires -RunAsAdministrator

$RutaVerificacion = "C:\HCG_Logs\Verificacion"

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "       VERIFICACION DE DATOS - FOTOS vs SISTEMA" -ForegroundColor Cyan
    Write-Host "       Hospital Civil de Guadalajara - FAA" -ForegroundColor Cyan
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Get-DatosCompletos {
    Write-Host "  Extrayendo datos del equipo..." -ForegroundColor Yellow
    Write-Host ""

    # BIOS / Sistema
    $Bios = Get-WmiObject Win32_BIOS
    $Sistema = Get-WmiObject Win32_ComputerSystem
    $BaseBoard = Get-WmiObject Win32_BaseBoard

    # Procesador
    $CPU = Get-WmiObject Win32_Processor | Select-Object -First 1

    # Memoria RAM
    $RAM = Get-WmiObject Win32_PhysicalMemory
    $RAMTotal = [math]::Round(($RAM | Measure-Object -Property Capacity -Sum).Sum / 1GB, 0)
    $RAMSlots = ($RAM | Measure-Object).Count
    $RAMDetalle = $RAM | ForEach-Object {
        $Capacidad = [math]::Round($_.Capacity / 1GB, 0)
        $Velocidad = $_.Speed
        "$($Capacidad)GB @ $($Velocidad)MHz"
    }

    # Discos
    $Discos = Get-WmiObject Win32_DiskDrive
    $DiscosDetalle = $Discos | ForEach-Object {
        $Tamano = [math]::Round($_.Size / 1GB, 0)
        $Tipo = if ($_.MediaType -like "*SSD*" -or $_.Model -like "*SSD*" -or $_.Model -like "*NVMe*") { "SSD" } else { "HDD" }
        @{
            Modelo = $_.Model
            Serial = $_.SerialNumber
            Tamano = "$($Tamano) GB"
            Tipo = $Tipo
            Interfaz = $_.InterfaceType
        }
    }

    # Redes
    $MACEth = (Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Ethernet" -and $_.MACAddress } | Select-Object -First 1).MACAddress
    $MACWifi = (Get-WmiObject Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Wi-Fi" -and $_.MACAddress } | Select-Object -First 1).MACAddress

    # Product Key
    $ProductKey = ""
    try {
        $ProductKey = (Get-WmiObject -Query "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService" | Where-Object { $_.OA3xOriginalProductKey }).OA3xOriginalProductKey
    } catch {}

    # Sistema Operativo
    $OS = Get-WmiObject Win32_OperatingSystem

    return @{
        # Identificacion
        NombreEquipo = $env:COMPUTERNAME
        Fabricante = $Sistema.Manufacturer
        Modelo = $Sistema.Model
        NumeroSerie = $Bios.SerialNumber

        # BIOS
        VersionBIOS = $Bios.SMBIOSBIOSVersion
        FechaBIOS = $Bios.ReleaseDate

        # Placa base
        FabricantePlaca = $BaseBoard.Manufacturer
        ModeloPlaca = $BaseBoard.Product
        SeriePlaca = $BaseBoard.SerialNumber

        # Procesador
        Procesador = $CPU.Name
        Nucleos = $CPU.NumberOfCores
        Threads = $CPU.NumberOfLogicalProcessors

        # Memoria
        RAMTotal = "$RAMTotal GB"
        RAMSlots = $RAMSlots
        RAMDetalle = $RAMDetalle -join " | "

        # Almacenamiento
        Discos = $DiscosDetalle

        # Redes
        MACEthernet = $MACEth
        MACWiFi = $MACWifi

        # Windows
        ProductKey = $ProductKey
        SistemaOperativo = $OS.Caption
        VersionOS = $OS.Version

        # Fecha extraccion
        FechaExtraccion = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    }
}

function Show-DatosParaVerificar {
    param($Datos)

    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "  DATOS EXTRAIDOS DEL SISTEMA (Comparar con etiqueta/foto)" -ForegroundColor Green
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""

    Write-Host "  --- IDENTIFICACION DEL EQUIPO ---" -ForegroundColor Cyan
    Write-Host "  Fabricante:     $($Datos.Fabricante)" -ForegroundColor White
    Write-Host "  Modelo:         $($Datos.Modelo)" -ForegroundColor White
    Write-Host "  No. Serie:      $($Datos.NumeroSerie)" -ForegroundColor Yellow
    Write-Host ""

    Write-Host "  --- PROCESADOR ---" -ForegroundColor Cyan
    Write-Host "  CPU:            $($Datos.Procesador)" -ForegroundColor White
    Write-Host "  Nucleos/Threads: $($Datos.Nucleos) / $($Datos.Threads)" -ForegroundColor White
    Write-Host ""

    Write-Host "  --- MEMORIA RAM ---" -ForegroundColor Cyan
    Write-Host "  Total:          $($Datos.RAMTotal)" -ForegroundColor Yellow
    Write-Host "  Modulos:        $($Datos.RAMSlots)" -ForegroundColor White
    Write-Host "  Detalle:        $($Datos.RAMDetalle)" -ForegroundColor White
    Write-Host ""

    Write-Host "  --- ALMACENAMIENTO ---" -ForegroundColor Cyan
    foreach ($Disco in $Datos.Discos) {
        Write-Host "  Disco:          $($Disco.Modelo)" -ForegroundColor White
        Write-Host "  Serial Disco:   $($Disco.Serial)" -ForegroundColor White
        Write-Host "  Capacidad:      $($Disco.Tamano)" -ForegroundColor Yellow
        Write-Host "  Tipo:           $($Disco.Tipo)" -ForegroundColor White
        Write-Host ""
    }

    Write-Host "  --- CONECTIVIDAD ---" -ForegroundColor Cyan
    Write-Host "  MAC Ethernet:   $($Datos.MACEthernet)" -ForegroundColor Yellow
    Write-Host "  MAC WiFi:       $($Datos.MACWiFi)" -ForegroundColor White
    Write-Host ""

    Write-Host "  --- LICENCIA WINDOWS ---" -ForegroundColor Cyan
    Write-Host "  Product Key:    $($Datos.ProductKey)" -ForegroundColor Yellow
    Write-Host "  Sistema:        $($Datos.SistemaOperativo)" -ForegroundColor White
    Write-Host ""
}

function Export-ReporteVerificacion {
    param($Datos, $InvST)

    if (-not (Test-Path $RutaVerificacion)) {
        New-Item -ItemType Directory -Path $RutaVerificacion -Force | Out-Null
    }

    $NombreArchivo = "Verificacion_$InvST`_$($Datos.NumeroSerie).txt"
    $RutaArchivo = Join-Path $RutaVerificacion $NombreArchivo

    $Contenido = @"
================================================================================
FICHA DE VERIFICACION - EQUIPO $InvST
================================================================================
Fecha de verificacion: $($Datos.FechaExtraccion)
Nombre equipo:         $($Datos.NombreEquipo)

================================================================================
DATOS PARA COMPARAR CON ETIQUETA DE LA CAJA
================================================================================

IDENTIFICACION:
  Fabricante:          $($Datos.Fabricante)
  Modelo:              $($Datos.Modelo)
  Numero de Serie:     $($Datos.NumeroSerie)

PROCESADOR:
  CPU:                 $($Datos.Procesador)
  Nucleos / Threads:   $($Datos.Nucleos) / $($Datos.Threads)

MEMORIA RAM:
  Total:               $($Datos.RAMTotal)
  Modulos instalados:  $($Datos.RAMSlots)
  Detalle:             $($Datos.RAMDetalle)

ALMACENAMIENTO:

"@

    foreach ($Disco in $Datos.Discos) {
        $Contenido += @"
  Disco:               $($Disco.Modelo)
  Serial del disco:    $($Disco.Serial)
  Capacidad:           $($Disco.Tamano)
  Tipo:                $($Disco.Tipo)
  Interfaz:            $($Disco.Interfaz)

"@
    }

    $Contenido += @"
CONECTIVIDAD:
  MAC Ethernet:        $($Datos.MACEthernet)
  MAC WiFi:            $($Datos.MACWiFi)

LICENCIA WINDOWS:
  Product Key:         $($Datos.ProductKey)
  Sistema Operativo:   $($Datos.SistemaOperativo)
  Version:             $($Datos.VersionOS)

================================================================================
CHECKLIST DE VERIFICACION (Marcar con X si coincide)
================================================================================

[ ] Numero de serie coincide con etiqueta
[ ] Modelo del equipo coincide
[ ] Capacidad RAM coincide con especificaciones
[ ] Capacidad disco duro coincide
[ ] Tipo de disco (SSD/HDD) coincide

NOTAS DE VERIFICACION:
_________________________________________________________________
_________________________________________________________________
_________________________________________________________________

Verificado por: ___________________________  Fecha: _______________

"@

    $Contenido | Out-File -FilePath $RutaArchivo -Encoding UTF8

    Write-Host "  [OK] Reporte guardado en:" -ForegroundColor Green
    Write-Host "       $RutaArchivo" -ForegroundColor Cyan

    return $RutaArchivo
}

function Export-DatosJSON {
    param($Datos, $InvST)

    $NombreArchivo = "DatosCompletos_$InvST`_$($Datos.NumeroSerie).json"
    $RutaArchivo = Join-Path $RutaVerificacion $NombreArchivo

    # Convertir discos a formato serializable
    $DatosExport = @{
        Inventario = $InvST
        NombreEquipo = $Datos.NombreEquipo
        Fabricante = $Datos.Fabricante
        Modelo = $Datos.Modelo
        NumeroSerie = $Datos.NumeroSerie
        Procesador = $Datos.Procesador
        Nucleos = $Datos.Nucleos
        Threads = $Datos.Threads
        RAMTotal = $Datos.RAMTotal
        RAMSlots = $Datos.RAMSlots
        RAMDetalle = $Datos.RAMDetalle
        MACEthernet = $Datos.MACEthernet
        MACWiFi = $Datos.MACWiFi
        ProductKey = $Datos.ProductKey
        SistemaOperativo = $Datos.SistemaOperativo
        VersionOS = $Datos.VersionOS
        FechaExtraccion = $Datos.FechaExtraccion
        Discos = @()
    }

    foreach ($Disco in $Datos.Discos) {
        $DatosExport.Discos += @{
            Modelo = $Disco.Modelo
            Serial = $Disco.Serial
            Tamano = $Disco.Tamano
            Tipo = $Disco.Tipo
            Interfaz = $Disco.Interfaz
        }
    }

    $DatosExport | ConvertTo-Json -Depth 10 | Out-File -FilePath $RutaArchivo -Encoding UTF8

    Write-Host "  [OK] Datos JSON guardados en:" -ForegroundColor Green
    Write-Host "       $RutaArchivo" -ForegroundColor Cyan

    return $RutaArchivo
}

# =============================================================================
# EJECUCION PRINCIPAL
# =============================================================================

Show-Banner

Write-Host "  Ingresa el numero de inventario ST (5 digitos)" -ForegroundColor Yellow
$NumInventario = Read-Host "  Numero"

if ($NumInventario -notmatch '^\d{5}$') {
    Write-Host "`n  [ERROR] Debe ser de 5 digitos" -ForegroundColor Red
    Read-Host "  Presiona Enter para salir"
    exit
}

Write-Host ""

# Extraer datos
$Datos = Get-DatosCompletos

# Mostrar en pantalla
Show-DatosParaVerificar -Datos $Datos

# Guardar reportes
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "  GUARDANDO REPORTES" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

$RutaTXT = Export-ReporteVerificacion -Datos $Datos -InvST $NumInventario
$RutaJSON = Export-DatosJSON -Datos $Datos -InvST $NumInventario

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "  VERIFICACION LISTA" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Archivos generados en: $RutaVerificacion" -ForegroundColor White
Write-Host ""
Write-Host "  1. Abre la foto de la caja del equipo" -ForegroundColor Yellow
Write-Host "  2. Compara los datos de la etiqueta con los mostrados arriba" -ForegroundColor Yellow
Write-Host "  3. Si hay diferencias, anotarlas en el reporte TXT" -ForegroundColor Yellow
Write-Host ""

$AbrirCarpeta = Read-Host "  Abrir carpeta de verificacion? (S/N)"
if ($AbrirCarpeta -eq "S" -or $AbrirCarpeta -eq "s") {
    Start-Process explorer.exe -ArgumentList $RutaVerificacion
}

Write-Host ""
Read-Host "  Presiona Enter para salir"
