$file = "C:\supportiA\ConfiguradorHCG\ConfigurarEquipoHCG.ps1"
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "SCRIPT PRINCIPAL: SIN ERRORES" -ForegroundColor Green
} else {
    Write-Host "SCRIPT PRINCIPAL: $($errors.Count) ERRORES" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "  Linea $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
    }
}
