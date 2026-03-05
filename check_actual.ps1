$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile('C:\supportiA\ConfiguradorHCG\ActualizarWindows11.ps1', [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "SCRIPT ACTUALIZADO: SIN ERRORES" -ForegroundColor Green
} else {
    Write-Host "SCRIPT: $($errors.Count) ERRORES" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "  Linea $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
    }
}
