# =============================================================================
# CONFIGURACION DEL SISTEMA - HOSPITAL CIVIL FAA
# =============================================================================
# Modifica estas variables segun tu entorno

# --- SERVIDORES DE RED ---
$Script:ServidorSoftware = "\\10.2.1.13\soportefaa"
$Script:ServidorDedalus = "\\10.2.1.17\distribucion"

# --- RUTAS DE SOFTWARE ---
$Script:RutaSoftware = "$ServidorSoftware\Software Equipo Nuevo\Programas"
$Script:RutaOffice2007 = "$RutaSoftware\Ofice2007"
$Script:RutaAcrobat = "$RutaSoftware\AcroRdrDC2100120155_es_ES.exe"
$Script:RutaChrome = "$RutaSoftware\49.0.2623.112_chrome_installer-bueno.exe"
$Script:RutaDotNet35 = "$RutaSoftware\dotnetfx35.exe"
$Script:RutaAccesos = "$RutaSoftware\Accesos"
$Script:RutaAntivirus = "$ServidorSoftware\ANTIVIRUS\ESET 2025\PROTECT_Installer_x64_es_CL 2.exe"
$Script:RutaDedalus = "$ServidorDedalus\dedalus"
$Script:RutaSincronizador = "$ServidorDedalus\dedalus\sincronizador"

# --- CREDENCIALES DE RED ---
$Script:UsuarioRed = "2010201"
# La contrasena se guarda en Windows Credential Manager, no aqui

# --- SERIAL DE OFFICE 2007 ---
$Script:SerialOffice2007 = "DR268-PQJVQ-X9JMJ-BTK78-DYVD8"

# --- USUARIO DE SOPORTE ---
$Script:NombreUsuarioSoporte = "Soporte"
$Script:PasswordSoporte = "*TIsoporte"
$Script:DescripcionSoporte = "Cuenta para soporte tecnico remoto - Ext. 54425"

# --- CONFIGURACION DEL EQUIPO ---
$Script:PrefijoNombreEquipo = "PC-"  # Ej: PC-13655
$Script:ZonaHoraria = "Central Standard Time (Mexico)"
$Script:ServidorNTP = "time.windows.com"

# --- GOOGLE SHEETS ---
$Script:GoogleSheetId = "17xKz71wrBHs0T54G3XGYFWCkGDW0AJqz"
$Script:GoogleSheetGid = "1975223929"
$Script:GoogleSheetUrl = "https://docs.google.com/spreadsheets/d/$GoogleSheetId/edit?gid=$GoogleSheetGid"

# --- RUTAS LOCALES ---
$Script:RutaReportesLocal = "C:\HCG_Reportes"
$Script:RutaLogs = "C:\HCG_Reportes\Logs"

# --- ESPECIFICACIONES DEL MODELO (para el registro) ---
$Script:EspecsEquipo = @{
    Marca = "Lenovo"
    Modelo = "ThinkCentre M70s Gen 5"
    Procesador = "Intel Core i5-14500 vPro"
    Nucleos = "14"
    RAM = "8 GB DDR5"
    Disco = "512 GB SSD NVMe"
    Graficos = "Intel UHD 770"
    WiFi = "Wi-Fi 6"
    Bluetooth = "5.1"
    SistemaOp = "Win 11 Pro"
    GarantiaAnios = 3
}

# --- APLICACIONES A INSTALAR ---
$Script:AplicacionesInstalar = @(
    @{ Nombre = "Office 2007"; Habilitado = $true }
    @{ Nombre = "ESET Antivirus"; Habilitado = $true }
    @{ Nombre = "Google Chrome"; Habilitado = $true }
    @{ Nombre = "Acrobat Reader DC"; Habilitado = $true }
    @{ Nombre = ".NET Framework 3.5"; Habilitado = $true }
    @{ Nombre = "Dedalus"; Habilitado = $true }
)

# --- LISTA OFICIAL DE SERIES (para verificacion) ---
$Script:RutaSeriesFAA = "C:\Users\lalog\Mi unidad (ejcontreras@hcg.gob.mx)\installerIA\seriesFAA.xlsx"

Write-Host "Configuracion cargada correctamente" -ForegroundColor Green
