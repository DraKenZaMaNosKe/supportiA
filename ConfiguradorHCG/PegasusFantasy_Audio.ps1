# Saint Seiya - Pegasus Fantasy usando audio real de Windows
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Threading;

public class AudioTone {
    [DllImport("winmm.dll")]
    public static extern uint mciSendString(string command, IntPtr returnString, uint returnSize, IntPtr callback);

    [DllImport("kernel32.dll")]
    public static extern bool Beep(uint frequency, uint duration);
}
"@

Write-Host ""
Write-Host "  PEGASUS FANTASY - Saint Seiya" -ForegroundColor Cyan
Write-Host "  Los Caballeros del Zodiaco" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Reproduciendo..." -ForegroundColor Green
Write-Host ""

# Usar Beep de kernel32 (mas compatible)
$notes = @(
    @{f=659; d=200},  # E5
    @{f=659; d=200},  # E5
    @{f=587; d=200},  # D5
    @{f=659; d=400},  # E5
    @{f=0; d=100},    # pausa
    @{f=784; d=200},  # G5
    @{f=784; d=200},  # G5
    @{f=698; d=200},  # F5
    @{f=659; d=400},  # E5
    @{f=0; d=100},    # pausa
    @{f=659; d=200},  # E5
    @{f=659; d=200},  # E5
    @{f=587; d=200},  # D5
    @{f=659; d=300},  # E5
    @{f=587; d=150},  # D5
    @{f=523; d=150},  # C5
    @{f=494; d=400},  # B4
    @{f=0; d=200},    # pausa
    @{f=440; d=200},  # A4
    @{f=494; d=200},  # B4
    @{f=523; d=200},  # C5
    @{f=587; d=200},  # D5
    @{f=659; d=600}   # E5
)

foreach ($note in $notes) {
    if ($note.f -gt 0) {
        [AudioTone]::Beep($note.f, $note.d)
    } else {
        Start-Sleep -Milliseconds $note.d
    }
}

Write-Host "  Seiya! Seiya!" -ForegroundColor Magenta
