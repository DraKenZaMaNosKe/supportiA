# Melodia Triste - Composicion Original
# Escala menor, tempo lento, notas descendentes

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class AudioTone {
    [DllImport("kernel32.dll")]
    public static extern bool Beep(uint frequency, uint duration);
}
"@

Write-Host ""
Write-Host "  ~ Melodia Triste ~" -ForegroundColor DarkCyan
Write-Host "  (cuando el codigo no compila...)" -ForegroundColor DarkGray
Write-Host ""

# Notas en escala menor (La menor / A minor)
$A3 = 220; $B3 = 247; $C4 = 262; $D4 = 294; $E4 = 330; $F4 = 349; $G4 = 392
$A4 = 440; $B4 = 494; $C5 = 523; $D5 = 587; $E5 = 659

# Melodia triste - lenta y descendente
$notes = @(
    # Inicio melancolico
    @{f=$E5; d=500},
    @{f=$D5; d=400},
    @{f=$C5; d=600},
    @{f=0; d=300},

    # Descenso triste
    @{f=$B4; d=400},
    @{f=$A4; d=500},
    @{f=$G4; d=400},
    @{f=$E4; d=700},
    @{f=0; d=400},

    # Suspenso emotivo
    @{f=$A4; d=300},
    @{f=$C5; d=300},
    @{f=$B4; d=600},
    @{f=$A4; d=400},
    @{f=0; d=200},

    # Caida final
    @{f=$G4; d=400},
    @{f=$F4; d=400},
    @{f=$E4; d=500},
    @{f=$D4; d=400},
    @{f=$C4; d=600},
    @{f=0; d=300},

    # Resolucion melancolica
    @{f=$A3; d=800},
    @{f=0; d=200},
    @{f=$A3; d=1000}
)

foreach ($note in $notes) {
    if ($note.f -gt 0) {
        [AudioTone]::Beep($note.f, $note.d) | Out-Null
    } else {
        Start-Sleep -Milliseconds $note.d
    }
}

Write-Host ""
Write-Host "  *sniff* :'(" -ForegroundColor DarkBlue
