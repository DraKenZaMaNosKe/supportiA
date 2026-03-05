$file = "E:\ConfiguradorHCG_USB\ActualizarWindows11.ps1"
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "USB SCRIPT: SIN ERRORES - LISTO PARA USAR" -ForegroundColor Green
} else {
    Write-Host "USB SCRIPT: $($errors.Count) ERRORES" -ForegroundColor Red
}
