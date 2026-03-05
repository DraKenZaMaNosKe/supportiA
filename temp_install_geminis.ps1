# Copiar script PS1 a System32
Copy-Item "C:\supportiA\ConfiguradorHCG\ScreenSaver_Geminis.ps1" "$env:WINDIR\System32\HCG_Geminis.ps1" -Force
Write-Host "[OK] HCG_Geminis.ps1 copiado a System32" -ForegroundColor Green

# Crear wrapper C#
$csSource = @'
using System;
using System.Diagnostics;
using System.Windows.Forms;

class GeminisScreenSaver
{
    [STAThread]
    static void Main(string[] args)
    {
        string mode = "/s";
        if (args.Length > 0)
            mode = args[0].ToLower().Substring(0, 2);

        switch (mode)
        {
            case "/s":
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = "-ExecutionPolicy Bypass -NoProfile -STA -File \"C:\\Windows\\System32\\HCG_Geminis.ps1\"",
                        UseShellExecute = false
                    };
                    var proc = Process.Start(psi);
                    proc.WaitForExit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Saga de Geminis",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                break;

            case "/c":
                MessageBox.Show(
                    "Saga de Geminis\n" +
                    "Caballero Dorado de la Tercera Casa\n\n" +
                    "No hay opciones configurables.",
                    "Saga de Geminis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;

            case "/p":
                break;
        }
    }
}
'@

$tempCs = "$env:TEMP\HCG_Geminis.cs"
$scrOutput = "$env:WINDIR\System32\HCG_Geminis.scr"
$csSource | Out-File -FilePath $tempCs -Encoding UTF8 -Force

# Compilar
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" }

$result = & $csc /target:winexe /out:"$scrOutput" /reference:System.Windows.Forms.dll "$tempCs" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] HCG_Geminis.scr compilado" -ForegroundColor Green
} else {
    Write-Host "[ERROR] Compilacion: $result" -ForegroundColor Red
    exit 1
}
Remove-Item $tempCs -Force -ErrorAction SilentlyContinue

# Configurar registro - 3 minutos
$regPath = "HKCU:\Control Panel\Desktop"
Set-ItemProperty -Path $regPath -Name "SCRNSAVE.EXE" -Value $scrOutput -Force
Set-ItemProperty -Path $regPath -Name "ScreenSaveActive" -Value "1" -Force
Set-ItemProperty -Path $regPath -Name "ScreenSaveTimeOut" -Value "180" -Force
Write-Host "[OK] Registro configurado - 3 minutos de espera" -ForegroundColor Green

Write-Host "`nProtector de pantalla SAGA DE GEMINIS instalado!" -ForegroundColor Yellow
