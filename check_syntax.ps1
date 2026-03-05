$file = "C:\supportiA\ConfiguradorHCG_USB\ActualizarWindows11.ps1"
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$null, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "SIN ERRORES DE SINTAXIS" -ForegroundColor Green
} else {
    Write-Host "ERRORES ENCONTRADOS: $($errors.Count)" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "  Linea $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Yellow
    }
}
