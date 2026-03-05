$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile('E:\ConfiguradorHCG_USB\ConfigurarEquipoHCG.ps1', [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "USB SCRIPT PRINCIPAL: SIN ERRORES - LISTO PARA USAR" -ForegroundColor Green
} else {
    Write-Host "USB SCRIPT: $($errors.Count) ERRORES" -ForegroundColor Red
}
