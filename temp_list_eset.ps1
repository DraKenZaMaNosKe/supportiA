$path = "\\10.2.1.13\soportefaa\pack_installer_iA\antivirus"
if (Test-Path $path) {
    Get-ChildItem $path -Recurse | Select-Object FullName, Length | Format-Table -AutoSize
} else {
    Write-Host "Ruta no encontrada: $path"
    # Intentar listar la carpeta padre
    $parent = "\\10.2.1.13\soportefaa\pack_installer_iA"
    if (Test-Path $parent) {
        Write-Host "Contenido de $parent :"
        Get-ChildItem $parent | Select-Object Name | Format-Table -AutoSize
    }
}
