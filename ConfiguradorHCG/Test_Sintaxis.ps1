# Verificar sintaxis
$scriptPath = "C:\supportiA\ConfiguradorHCG\ConfigurarEquipoHCG.ps1"
$parseErrors = $null
$tokens = $null

[System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)

if ($parseErrors.Count -gt 0) {
    Write-Host "Errores encontrados:" -ForegroundColor Red
    foreach ($err in $parseErrors) {
        Write-Host "  Linea $($err.Extent.StartLineNumber): $($err.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "Sin errores de sintaxis" -ForegroundColor Green
}
