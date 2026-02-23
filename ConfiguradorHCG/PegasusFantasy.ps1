# Saint Seiya - Pegasus Fantasy (Intro)
# Melodia aproximada del opening

function Play-Note {
    param($freq, $duration)
    [Console]::Beep($freq, $duration)
}

# Frecuencias de notas
$E4 = 330; $F4 = 349; $G4 = 392; $A4 = 440; $B4 = 494
$C5 = 523; $D5 = 587; $E5 = 659; $F5 = 698; $G5 = 784; $A5 = 880

Write-Host ""
Write-Host "  ★ PEGASUS FANTASY - Saint Seiya ★" -ForegroundColor Cyan
Write-Host "  Los Caballeros del Zodiaco" -ForegroundColor Yellow
Write-Host ""

# Intro iconico
Play-Note $E5 200
Play-Note $E5 200
Play-Note $D5 200
Play-Note $E5 400
Start-Sleep -Milliseconds 100

Play-Note $G5 200
Play-Note $G5 200
Play-Note $F5 200
Play-Note $E5 400
Start-Sleep -Milliseconds 100

# Segunda parte
Play-Note $E5 200
Play-Note $E5 200
Play-Note $D5 200
Play-Note $E5 300
Play-Note $D5 150
Play-Note $C5 150
Play-Note $B4 400
Start-Sleep -Milliseconds 200

# Remate epico
Play-Note $A4 200
Play-Note $B4 200
Play-Note $C5 200
Play-Note $D5 200
Play-Note $E5 600

Write-Host ""
Write-Host "  Seiya! Seiya!" -ForegroundColor Magenta
