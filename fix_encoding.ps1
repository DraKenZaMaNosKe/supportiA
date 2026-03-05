$source = "C:\supportiA\ConfiguradorHCG\ActualizarWindows11.ps1"
$dest = "C:\supportiA\ConfiguradorHCG_USB\ActualizarWindows11.ps1"
$content = Get-Content $source -Raw
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($dest, $content, $utf8NoBom)
Write-Host "Archivo guardado con UTF-8 sin BOM"
