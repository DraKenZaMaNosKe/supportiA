$file = "C:\supportiA\ConfiguradorHCG\ActualizarWindows11.ps1"
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "ORIGINAL: SIN ERRORES" -ForegroundColor Green
} else {
    Write-Host "ORIGINAL: $($errors.Count) ERRORES" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "  Linea $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
    }
}
