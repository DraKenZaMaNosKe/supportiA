$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile('E:\ConfiguradorHCG_USB\ActualizarWindows11.ps1', [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "ACTUALIZAR WINDOWS 11: SIN ERRORES - LISTO PARA USAR" -ForegroundColor Green
} else {
    Write-Host "ACTUALIZAR WINDOWS 11: $($errors.Count) ERRORES" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "  Linea $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
    }
}
