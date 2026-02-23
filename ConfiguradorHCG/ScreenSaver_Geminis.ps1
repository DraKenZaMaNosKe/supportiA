# =============================================================================
# SAGA DE GEMINIS - PROTECTOR DE PANTALLA
# =============================================================================
# Caballero Dorado de la Tercera Casa
# Constelacion de Geminis + Galaxian Explosion
# =============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Configuracion ---
$frameMs = 33
$numStars = 220
$numParticles = 80

# --- Ventana fullscreen ---
$window = New-Object System.Windows.Window
$window.WindowStyle = 'None'
$window.WindowState = 'Maximized'
$window.Background = [System.Windows.Media.Brushes]::Black
$window.Topmost = $true
$window.Cursor = [System.Windows.Input.Cursors]::None

$canvas = New-Object System.Windows.Controls.Canvas
$canvas.Background = [System.Windows.Media.Brushes]::Black
$canvas.ClipToBounds = $true
$window.Content = $canvas

# --- Variables globales ---
$script:screenW = 0
$script:screenH = 0
$script:mouseStartX = -1
$script:mouseStartY = -1
$script:startTime = [DateTime]::Now
$script:frameCount = 0
$script:stars = @()
$script:conLines = @()
$script:conStarElements = @()
$script:particles = @()
$script:burstActive = $false
$script:burstTime = 0
$script:lastBurstTime = 0
$script:burstIsPurple = $false

# --- Colores ---
$script:gold = [System.Windows.Media.Color]::FromRgb(255, 215, 0)
$script:goldBright = [System.Windows.Media.Color]::FromRgb(255, 240, 120)
$script:goldDim = [System.Windows.Media.Color]::FromRgb(180, 150, 0)
$script:purple = [System.Windows.Media.Color]::FromRgb(160, 50, 255)
$script:purpleBright = [System.Windows.Media.Color]::FromRgb(200, 120, 255)
$script:cosmicBlue = [System.Windows.Media.Color]::FromRgb(100, 140, 255)

# --- Constelacion de Geminis (coordenadas normalizadas) ---
$script:conPoints = @(
    @{ X = 0.38; Y = 0.15; S = 7 },   # 0: Castor
    @{ X = 0.62; Y = 0.13; S = 8 },   # 1: Pollux
    @{ X = 0.35; Y = 0.30; S = 4 },   # 2: Mebsuta
    @{ X = 0.44; Y = 0.44; S = 4 },   # 3: Wasat
    @{ X = 0.56; Y = 0.48; S = 4 },   # 4: Mekbuda
    @{ X = 0.64; Y = 0.64; S = 5 },   # 5: Alhena
    @{ X = 0.28; Y = 0.56; S = 4 },   # 6: Propus
    @{ X = 0.26; Y = 0.47; S = 3 },   # 7: Tejat
    @{ X = 0.60; Y = 0.32; S = 4 }    # 8: Kappa
)

$script:conEdges = @(
    @(0, 2), @(2, 3), @(3, 7), @(7, 6),
    @(1, 8), @(8, 4), @(4, 5),
    @(3, 4)
)

$script:attackPhrases = @(
    "GALAXIAN EXPLOSION",
    "ANOTHER DIMENSION",
    "GALAXIAN EXPLOSION"
)

# --- INICIALIZACION ---
$window.Add_Loaded({
    $script:screenW = $canvas.ActualWidth
    $script:screenH = $canvas.ActualHeight

    # === ESTRELLAS DE FONDO ===
    for ($i = 0; $i -lt $numStars; $i++) {
        $size = (Get-Random -Minimum 1 -Maximum 4)
        $star = New-Object System.Windows.Shapes.Ellipse
        $star.Width = $size
        $star.Height = $size

        $colorRoll = Get-Random -Minimum 0 -Maximum 100
        if ($colorRoll -lt 12) {
            $star.Fill = New-Object System.Windows.Media.SolidColorBrush($script:gold)
        } elseif ($colorRoll -lt 22) {
            $star.Fill = New-Object System.Windows.Media.SolidColorBrush($script:cosmicBlue)
        } else {
            $b = Get-Random -Minimum 170 -Maximum 255
            $star.Fill = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromRgb($b, $b, [Math]::Min(255, $b + 20)))
        }

        $star.Opacity = (Get-Random -Minimum 15 -Maximum 90) / 100.0

        [System.Windows.Controls.Canvas]::SetLeft($star, (Get-Random -Minimum 0 -Maximum ([int]$script:screenW)))
        [System.Windows.Controls.Canvas]::SetTop($star, (Get-Random -Minimum 0 -Maximum ([int]$script:screenH)))
        $canvas.Children.Add($star) | Out-Null

        $script:stars += @{
            El = $star
            BaseOp = $star.Opacity
            Speed = (Get-Random -Minimum 5 -Maximum 30) / 10.0
            Phase = (Get-Random -Minimum 0 -Maximum 628) / 100.0
        }
    }

    # === CONSTELACION: LINEAS ===
    $conScale = [Math]::Min($script:screenW, $script:screenH) * 0.75
    $conOX = ($script:screenW - $conScale) / 2
    $conOY = ($script:screenH - $conScale) / 2 - $script:screenH * 0.05

    foreach ($edge in $script:conEdges) {
        $p1 = $script:conPoints[$edge[0]]
        $p2 = $script:conPoints[$edge[1]]

        $line = New-Object System.Windows.Shapes.Line
        $line.X1 = $conOX + $p1.X * $conScale
        $line.Y1 = $conOY + $p1.Y * $conScale
        $line.X2 = $conOX + $p2.X * $conScale
        $line.Y2 = $conOY + $p2.Y * $conScale
        $line.Stroke = New-Object System.Windows.Media.SolidColorBrush($script:goldDim)
        $line.StrokeThickness = 1.5
        $line.Opacity = 0
        $canvas.Children.Add($line) | Out-Null
        $script:conLines += @{ El = $line }
    }

    # === CONSTELACION: ESTRELLAS ===
    foreach ($pt in $script:conPoints) {
        $cs = New-Object System.Windows.Shapes.Ellipse
        $cs.Width = $pt.S
        $cs.Height = $pt.S
        $cs.Fill = New-Object System.Windows.Media.SolidColorBrush($script:goldBright)
        $cs.Opacity = 0

        $gl = New-Object System.Windows.Media.Effects.DropShadowEffect
        $gl.Color = $script:gold
        $gl.BlurRadius = 15
        $gl.ShadowDepth = 0
        $gl.Opacity = 0.8
        $cs.Effect = $gl

        [System.Windows.Controls.Canvas]::SetLeft($cs, $conOX + $pt.X * $conScale - $pt.S / 2)
        [System.Windows.Controls.Canvas]::SetTop($cs, $conOY + $pt.Y * $conScale - $pt.S / 2)
        $canvas.Children.Add($cs) | Out-Null
        $script:conStarElements += @{ El = $cs }
    }

    # === PARTICULAS (Galaxian Explosion / Another Dimension) ===
    for ($i = 0; $i -lt $numParticles; $i++) {
        $pe = New-Object System.Windows.Shapes.Ellipse
        $pSize = Get-Random -Minimum 2 -Maximum 6
        $pe.Width = $pSize
        $pe.Height = $pSize
        $pe.Fill = New-Object System.Windows.Media.SolidColorBrush($script:goldBright)
        $pe.Opacity = 0

        $pg = New-Object System.Windows.Media.Effects.DropShadowEffect
        $pg.Color = $script:gold
        $pg.BlurRadius = 6
        $pg.ShadowDepth = 0
        $pg.Opacity = 0.5
        $pe.Effect = $pg

        [System.Windows.Controls.Canvas]::SetLeft($pe, $script:screenW / 2)
        [System.Windows.Controls.Canvas]::SetTop($pe, $script:screenH / 2)
        $canvas.Children.Add($pe) | Out-Null

        $script:particles += @{
            El = $pe
            VX = 0; VY = 0
            PX = 0; PY = 0
            Life = 999; MaxLife = 60
        }
    }

    # === SIMBOLO GEMINIS (fondo, muy sutil) ===
    $gemSym = New-Object System.Windows.Controls.TextBlock
    $gemSym.Text = [char]0x264A
    $gemSym.FontSize = 200
    $gemSym.Foreground = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromArgb(25, 255, 215, 0))
    $gemSym.Opacity = 0
    $gemSym.TextAlignment = "Center"
    $canvas.Children.Add($gemSym) | Out-Null

    $gemSym.Measure([System.Windows.Size]::new($script:screenW, $script:screenH))
    [System.Windows.Controls.Canvas]::SetLeft($gemSym, ($script:screenW - $gemSym.DesiredSize.Width) / 2)
    [System.Windows.Controls.Canvas]::SetTop($gemSym, $script:screenH * 0.30)
    $script:gemSym = $gemSym

    # === TITULO ===
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = "SAGA DE G$([char]0xC9)MINIS"
    $title.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $title.FontSize = 54
    $title.FontWeight = [System.Windows.FontWeights]::Bold
    $title.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:gold)
    $title.Opacity = 0
    $title.TextAlignment = "Center"

    $tGlow = New-Object System.Windows.Media.Effects.DropShadowEffect
    $tGlow.Color = $script:gold
    $tGlow.BlurRadius = 35
    $tGlow.ShadowDepth = 0
    $tGlow.Opacity = 0.9
    $title.Effect = $tGlow
    $canvas.Children.Add($title) | Out-Null

    # Subtitulo
    $sub = New-Object System.Windows.Controls.TextBlock
    $sub.Text = "CABALLERO DORADO DE LA TERCERA CASA"
    $sub.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $sub.FontSize = 18
    $sub.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:goldDim)
    $sub.Opacity = 0
    $sub.TextAlignment = "Center"
    $canvas.Children.Add($sub) | Out-Null

    # Texto de ataque
    $atkText = New-Object System.Windows.Controls.TextBlock
    $atkText.Text = ""
    $atkText.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $atkText.FontSize = 36
    $atkText.FontWeight = [System.Windows.FontWeights]::Bold
    $atkText.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:goldBright)
    $atkText.Opacity = 0
    $atkText.TextAlignment = "Center"

    $aGlow = New-Object System.Windows.Media.Effects.DropShadowEffect
    $aGlow.Color = $script:goldBright
    $aGlow.BlurRadius = 25
    $aGlow.ShadowDepth = 0
    $aGlow.Opacity = 1
    $atkText.Effect = $aGlow
    $canvas.Children.Add($atkText) | Out-Null

    # Posicionar titulos
    $title.Measure([System.Windows.Size]::new($script:screenW, $script:screenH))
    $sub.Measure([System.Windows.Size]::new($script:screenW, $script:screenH))

    $titleY = $script:screenH * 0.80
    [System.Windows.Controls.Canvas]::SetLeft($title, ($script:screenW - $title.DesiredSize.Width) / 2)
    [System.Windows.Controls.Canvas]::SetTop($title, $titleY)
    [System.Windows.Controls.Canvas]::SetLeft($sub, ($script:screenW - $sub.DesiredSize.Width) / 2)
    [System.Windows.Controls.Canvas]::SetTop($sub, $titleY + 62)

    [System.Windows.Controls.Canvas]::SetLeft($atkText, 0)
    [System.Windows.Controls.Canvas]::SetTop($atkText, $script:screenH * 0.45)

    $script:title = $title
    $script:sub = $sub
    $script:atkText = $atkText

    $script:timer.Start()
})

# --- ANIMACION ---
$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds($frameMs)

$script:timer.Add_Tick({
    $script:frameCount++
    $elapsed = ([DateTime]::Now - $script:startTime).TotalSeconds

    # === ESTRELLAS: PARPADEO ===
    foreach ($s in $script:stars) {
        $twinkle = [Math]::Sin($elapsed * $s.Speed + $s.Phase)
        $s.El.Opacity = $s.BaseOp * (0.5 + $twinkle * 0.5)
    }

    # === CONSTELACION: FADE IN (2s-5s) luego pulso ===
    if ($elapsed -gt 2 -and $elapsed -lt 5) {
        $cf = [Math]::Min(1, ($elapsed - 2) / 3.0)
        foreach ($cl in $script:conLines) { $cl.El.Opacity = $cf * 0.6 }
        foreach ($cs in $script:conStarElements) { $cs.El.Opacity = $cf }
    }
    elseif ($elapsed -ge 5) {
        $cp = 0.55 + [Math]::Sin($elapsed * 0.8) * 0.15
        foreach ($cl in $script:conLines) { $cl.El.Opacity = $cp }
        foreach ($cs in $script:conStarElements) {
            $cs.El.Opacity = $cp + 0.25
            $cs.El.Effect.BlurRadius = 12 + [Math]::Sin($elapsed * 1.2) * 6
        }
    }

    # === SIMBOLO GEMINIS: FADE IN SUTIL ===
    if ($elapsed -gt 3 -and $elapsed -lt 6) {
        $script:gemSym.Opacity = [Math]::Min(1, ($elapsed - 3) / 3.0)
    } elseif ($elapsed -ge 6) {
        $script:gemSym.Opacity = 0.85 + [Math]::Sin($elapsed * 0.4) * 0.15
    }

    # === TITULO: FADE IN (4s-7s) luego pulso dorado ===
    if ($elapsed -gt 4 -and $elapsed -lt 7) {
        $tf = [Math]::Min(1, ($elapsed - 4) / 3.0)
        $script:title.Opacity = $tf
        $script:sub.Opacity = $tf * 0.5
    }
    elseif ($elapsed -ge 7) {
        $tp = [Math]::Sin($elapsed * 1.2) * 0.12 + 0.88
        $script:title.Opacity = $tp
        $script:sub.Opacity = $tp * 0.5
        $script:title.Effect.BlurRadius = 30 + [Math]::Sin($elapsed * 1.5) * 12
    }

    # === GALAXIAN EXPLOSION / ANOTHER DIMENSION (cada ~10s) ===
    if ($elapsed -gt 8 -and (-not $script:burstActive) -and ($elapsed - $script:lastBurstTime) -gt 10) {
        $script:burstActive = $true
        $script:burstTime = 0
        $script:lastBurstTime = $elapsed

        # Elegir ataque
        $phraseIdx = Get-Random -Minimum 0 -Maximum $script:attackPhrases.Count
        $phrase = $script:attackPhrases[$phraseIdx]
        $script:burstIsPurple = ($phrase -eq "ANOTHER DIMENSION")

        $script:atkText.Text = [char]0xA1 + $phrase + "!"
        $script:atkText.Measure([System.Windows.Size]::new($script:screenW * 2, $script:screenH))
        [System.Windows.Controls.Canvas]::SetLeft($script:atkText, ($script:screenW - $script:atkText.DesiredSize.Width) / 2)

        # Color segun ataque
        if ($script:burstIsPurple) {
            $script:atkText.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:purpleBright)
            $script:atkText.Effect.Color = $script:purple
        } else {
            $script:atkText.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:goldBright)
            $script:atkText.Effect.Color = $script:gold
        }

        # Reiniciar particulas
        $cx = $script:screenW / 2
        $cy = $script:screenH / 2
        foreach ($p in $script:particles) {
            $angle = (Get-Random -Minimum 0 -Maximum 628) / 100.0
            $spd = (Get-Random -Minimum 25 -Maximum 110) / 10.0
            $p.VX = [Math]::Cos($angle) * $spd
            $p.VY = [Math]::Sin($angle) * $spd
            $p.PX = $cx
            $p.PY = $cy
            $p.Life = 0
            $p.MaxLife = Get-Random -Minimum 35 -Maximum 80

            if ($script:burstIsPurple) {
                $p.El.Fill = New-Object System.Windows.Media.SolidColorBrush($script:purpleBright)
                $p.El.Effect.Color = $script:purple
            } else {
                $p.El.Fill = New-Object System.Windows.Media.SolidColorBrush($script:goldBright)
                $p.El.Effect.Color = $script:gold
            }
        }
    }

    # === ANIMAR EXPLOSION ===
    if ($script:burstActive) {
        $script:burstTime++

        # Texto de ataque: fade in/out
        if ($script:burstTime -lt 12) {
            $script:atkText.Opacity = $script:burstTime / 12.0
        } elseif ($script:burstTime -gt 55) {
            $script:atkText.Opacity = [Math]::Max(0, 1 - ($script:burstTime - 55) / 20.0)
        } else {
            $script:atkText.Opacity = 0.85 + [Math]::Sin($script:burstTime * 0.5) * 0.15
        }

        # Mover particulas
        $allDone = $true
        foreach ($p in $script:particles) {
            $p.Life++
            if ($p.Life -lt $p.MaxLife) {
                $allDone = $false
                $p.PX += $p.VX
                $p.PY += $p.VY
                $p.VX *= 0.97
                $p.VY *= 0.97
                [System.Windows.Controls.Canvas]::SetLeft($p.El, $p.PX)
                [System.Windows.Controls.Canvas]::SetTop($p.El, $p.PY)
                $ratio = 1 - ($p.Life / $p.MaxLife)
                $p.El.Opacity = $ratio * $ratio
            } else {
                $p.El.Opacity = 0
            }
        }

        # Efecto en constelacion durante Another Dimension
        if ($script:burstIsPurple -and $script:burstTime -lt 50) {
            $purpleIntensity = [Math]::Sin($script:burstTime * 0.15) * 0.5 + 0.5
            foreach ($cl in $script:conLines) {
                $cl.El.Stroke = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.Color]::FromRgb(
                        [byte](180 * (1-$purpleIntensity) + 160 * $purpleIntensity),
                        [byte](150 * (1-$purpleIntensity) + 50 * $purpleIntensity),
                        [byte](0 * (1-$purpleIntensity) + 255 * $purpleIntensity)
                    ))
            }
        }

        if ($script:burstTime -gt 85 -or ($script:burstTime -gt 50 -and $allDone)) {
            $script:burstActive = $false
            $script:atkText.Opacity = 0
            foreach ($p in $script:particles) { $p.El.Opacity = 0 }
            # Restaurar color dorado de constelacion
            foreach ($cl in $script:conLines) {
                $cl.El.Stroke = New-Object System.Windows.Media.SolidColorBrush($script:goldDim)
            }
        }
    }
})

# --- SALIR ---
$window.Add_KeyDown({
    $elapsed = ([DateTime]::Now - $script:startTime).TotalSeconds
    if ($elapsed -lt 3) { return }
    $script:timer.Stop()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
})

$window.Add_MouseMove({
    param($s, $e)
    $elapsed = ([DateTime]::Now - $script:startTime).TotalSeconds
    if ($elapsed -lt 5) { return }
    $pos = $e.GetPosition($window)
    if ($script:mouseStartX -eq -1) {
        $script:mouseStartX = $pos.X
        $script:mouseStartY = $pos.Y
        return
    }
    $dx = [Math]::Abs($pos.X - $script:mouseStartX)
    $dy = [Math]::Abs($pos.Y - $script:mouseStartY)
    if ($dx -gt 50 -or $dy -gt 50) {
        $script:timer.Stop()
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
    }
})

$window.Add_MouseDown({
    $elapsed = ([DateTime]::Now - $script:startTime).TotalSeconds
    if ($elapsed -lt 5) { return }
    $script:timer.Stop()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
})

# --- EJECUTAR ---
$window.Show()
[System.Windows.Threading.Dispatcher]::Run()
