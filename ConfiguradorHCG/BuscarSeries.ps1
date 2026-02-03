$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Usar archivo de Downloads
$filePath = "C:\Users\lalog\Downloads\series equipo FAA.xlsx"
Write-Host "Abriendo: $filePath"
$workbook = $excel.Workbooks.Open($filePath)
$sheet = $workbook.Sheets.Item(1)
Write-Host "Archivo abierto correctamente"

$seriales = @('MZ02WAT1', 'MZ02W5K4', 'MZ02W5PB')

foreach ($serial in $seriales) {
    $found = $sheet.UsedRange.Find($serial)
    if ($found) {
        $row = $found.Row
        $numSI = $sheet.Cells.Item($row, 1).Text
        Write-Host "Serial: $serial | Fila Excel: $row | SI #$numSI"
    } else {
        Write-Host "Serial: $serial | NO ENCONTRADO"
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
