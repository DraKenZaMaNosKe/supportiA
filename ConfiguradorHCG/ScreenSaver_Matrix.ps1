# =============================================================================
# HCG - PROTECTOR DE PANTALLA MATRIX
# =============================================================================
# "HOSPITAL CIVIL DE GUADALAJARA" estilo Matrix
# Presiona cualquier tecla o mueve el mouse para salir
# =============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Configuracion ---
$fontSize = 16
$charHeight = 20
$numRainCols = 90
$frameMs = 45

# --- Caracteres Matrix (katakana generados por Unicode + numeros + simbolos) ---
$matrixChars = ""
for ($i = 0xFF66; $i -le 0xFF9D; $i++) { $matrixChars += [char]$i }
$matrixChars += "0123456789ABCDEF<>{}|=+-*:.;~"

# --- Mensajes ocultos medicos (aparecen sutilmente entre la lluvia) ---
$medicalTerms = @(
    "PENICILINA", "PARACETAMOL", "ASPIRINA", "MORFINA", "INSULINA",
    "METFORMINA", "AMOXICILINA", "IBUPROFENO", "DICLOFENACO", "OMEPRAZOL",
    "C8H9NO2", "C9H8O4", "NaCl", "H2O2", "KCl",
    "DIAGNOSTICO", "ESTETOSCOPIO", "HEMOGLOBINA", "LEUCOCITOS", "ERITROCITOS",
    "PLAQUETAS", "CORAZON", "CEREBRO", "PULMON", "HIGADO",
    "HIPERTENSION", "DIABETES", "ANEMIA", "TAQUICARDIA", "BRADICARDIA",
    "CIRUGIA", "BIOPSIA", "SUTURA", "DIALISIS", "TOMOGRAFIA",
    "URGENCIAS", "PEDIATRIA", "ONCOLOGIA", "CARDIOLOGIA", "NEUROLOGIA",
    "FARMACIA", "LABORATORIO", "RADIOLOGIA", "PATOLOGIA", "ANATOMIA",
    "ADN", "ARN", "GENOMA", "CELULA", "TEJIDO",
    "ANTICUERPO", "ANTIGENO", "VACUNA", "VITAMINA", "PROTEINA",
    "OXIGENO", "BISTURI", "JERINGA", "CATETER", "SUERO",
    "HOSPITAL", "CIVIL", "GUADALAJARA", "OPD", "SALUD",
    "MEDICO", "ENFERMERA", "PACIENTE", "RECETA", "DOSIS",
    "QUIROFANO", "ANESTESIA", "ENDOSCOPIA", "RESONANCIA", "ECOGRAFIA"
)

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
$script:columns = @()
$script:colTextBlocks = @()
$script:titleVisible = $false

# --- Colores Matrix ---
$script:greenHead = [System.Windows.Media.Color]::FromRgb(220, 255, 220)
$script:greenBright = [System.Windows.Media.Color]::FromRgb(0, 255, 65)
$script:greenMid = [System.Windows.Media.Color]::FromRgb(0, 180, 40)
$script:greenDim = [System.Windows.Media.Color]::FromRgb(0, 80, 20)
$script:greenVDim = [System.Windows.Media.Color]::FromRgb(0, 30, 8)
$script:transparent = [System.Windows.Media.Color]::FromArgb(0, 0, 255, 0)

function Get-RandomChar {
    $matrixChars[(Get-Random -Minimum 0 -Maximum $matrixChars.Length)]
}

function Get-ColumnText {
    param([int]$numRows)
    $text = ""
    for ($r = 0; $r -lt $numRows; $r++) {
        if ($r -gt 0) { $text += "`n" }
        $text += (Get-RandomChar)
    }
    return $text
}

function New-RainGradient {
    param([double]$headNorm, [double]$tailNorm)

    $brush = New-Object System.Windows.Media.LinearGradientBrush
    $brush.StartPoint = New-Object System.Windows.Point(0.5, 0)
    $brush.EndPoint = New-Object System.Windows.Point(0.5, 1)

    # Clamp values
    if ($tailNorm -lt 0) { $tailNorm = 0 }
    if ($headNorm -gt 1) { $headNorm = 1 }
    if ($tailNorm -ge $headNorm) { $tailNorm = $headNorm - 0.01 }

    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop($script:transparent, [Math]::Max(0, $tailNorm - 0.02))))
    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop($script:greenVDim, $tailNorm)))
    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop($script:greenDim, ($tailNorm + $headNorm) / 2)))
    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop($script:greenBright, [Math]::Max(0, $headNorm - 0.03))))
    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop($script:greenHead, $headNorm)))
    $brush.GradientStops.Add((New-Object System.Windows.Media.GradientStop($script:transparent, [Math]::Min(1, $headNorm + 0.01))))

    return $brush
}

# --- TITULO ---
function Create-Title {
    # Linea 1
    $title1 = New-Object System.Windows.Controls.TextBlock
    $title1.Text = "HOSPITAL CIVIL"
    $title1.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $title1.FontSize = 52
    $title1.FontWeight = [System.Windows.FontWeights]::Bold
    $title1.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:greenBright)
    $title1.Opacity = 0
    $title1.TextAlignment = "Center"

    $glow1 = New-Object System.Windows.Media.Effects.DropShadowEffect
    $glow1.Color = $script:greenBright
    $glow1.BlurRadius = 30
    $glow1.ShadowDepth = 0
    $glow1.Opacity = 0.9
    $title1.Effect = $glow1

    $canvas.Children.Add($title1) | Out-Null

    # Linea 2
    $title2 = New-Object System.Windows.Controls.TextBlock
    $title2.Text = "DE GUADALAJARA"
    $title2.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $title2.FontSize = 52
    $title2.FontWeight = [System.Windows.FontWeights]::Bold
    $title2.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:greenBright)
    $title2.Opacity = 0
    $title2.TextAlignment = "Center"

    $glow2 = New-Object System.Windows.Media.Effects.DropShadowEffect
    $glow2.Color = $script:greenBright
    $glow2.BlurRadius = 30
    $glow2.ShadowDepth = 0
    $glow2.Opacity = 0.9
    $title2.Effect = $glow2

    $canvas.Children.Add($title2) | Out-Null

    # Subtitulo
    $sub = New-Object System.Windows.Controls.TextBlock
    $sub.Text = "- OPD -"
    $sub.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
    $sub.FontSize = 20
    $sub.Foreground = New-Object System.Windows.Media.SolidColorBrush($script:greenMid)
    $sub.Opacity = 0
    $sub.TextAlignment = "Center"
    $canvas.Children.Add($sub) | Out-Null

    $script:title1 = $title1
    $script:title2 = $title2
    $script:titleSub = $sub
}

# --- INICIALIZACION ---
$window.Add_Loaded({
    $script:screenW = $canvas.ActualWidth
    $script:screenH = $canvas.ActualHeight
    $numRows = [int]($script:screenH / $charHeight) + 2

    $colWidth = $script:screenW / $numRainCols

    # Crear columnas de lluvia
    for ($c = 0; $c -lt $numRainCols; $c++) {
        $speed = (Get-Random -Minimum 15 -Maximum 55) / 10.0
        $trailLen = Get-Random -Minimum 8 -Maximum 28

        $col = @{
            HeadY       = -(Get-Random -Minimum 0 -Maximum ([int]$script:screenH))
            Speed       = $speed
            TrailLength = $trailLen
            TrailPx     = $trailLen * $charHeight
            CharChangeTimer = (Get-Random -Minimum 3 -Maximum 15)
            HasMessage    = $false
            MessageText   = ""
            MessageStartRow = -1
        }

        # 15% de columnas tienen un mensaje oculto medico
        if ((Get-Random -Minimum 0 -Maximum 100) -lt 15) {
            $msg = $medicalTerms[(Get-Random -Minimum 0 -Maximum $medicalTerms.Count)]
            $maxStart = $numRows - $msg.Length - 2
            if ($maxStart -gt 2) {
                $col.HasMessage = $true
                $col.MessageText = $msg
                $col.MessageStartRow = (Get-Random -Minimum 2 -Maximum $maxStart)
            }
        }

        $script:columns += $col

        # TextBlock para esta columna
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = (Get-ColumnText -numRows $numRows)

        # Inyectar mensaje oculto si esta columna tiene uno
        if ($col.HasMessage) {
            $lines = $tb.Text.Split("`n")
            for ($mi = 0; $mi -lt $col.MessageText.Length; $mi++) {
                $rowIdx = $col.MessageStartRow + $mi
                if ($rowIdx -ge 0 -and $rowIdx -lt $lines.Count) {
                    $lines[$rowIdx] = [string]$col.MessageText[$mi]
                }
            }
            $tb.Text = $lines -join "`n"
        }

        $tb.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")
        $tb.FontSize = $fontSize
        $tb.LineHeight = $charHeight
        $tb.LineStackingStrategy = "BlockLineHeight"
        $tb.Foreground = [System.Windows.Media.Brushes]::Transparent

        [System.Windows.Controls.Canvas]::SetLeft($tb, $c * $colWidth)
        [System.Windows.Controls.Canvas]::SetTop($tb, 0)
        $canvas.Children.Add($tb) | Out-Null
        $script:colTextBlocks += $tb
    }

    # Crear titulo (se agrega al final para que quede encima)
    Create-Title

    # Posicionar titulo centrado
    $script:title1.Measure([System.Windows.Size]::new($script:screenW, $script:screenH))
    $script:title2.Measure([System.Windows.Size]::new($script:screenW, $script:screenH))

    $t1W = $script:title1.DesiredSize.Width
    $t2W = $script:title2.DesiredSize.Width
    $subW = $script:titleSub.DesiredSize.Width

    $centerY = $script:screenH / 2 - 60

    [System.Windows.Controls.Canvas]::SetLeft($script:title1, ($script:screenW - $t1W) / 2)
    [System.Windows.Controls.Canvas]::SetTop($script:title1, $centerY)

    [System.Windows.Controls.Canvas]::SetLeft($script:title2, ($script:screenW - $t2W) / 2)
    [System.Windows.Controls.Canvas]::SetTop($script:title2, $centerY + 65)

    $script:titleSub.Measure([System.Windows.Size]::new($script:screenW, $script:screenH))
    [System.Windows.Controls.Canvas]::SetLeft($script:titleSub, ($script:screenW - $script:titleSub.DesiredSize.Width) / 2)
    [System.Windows.Controls.Canvas]::SetTop($script:titleSub, $centerY + 135)

    # Iniciar animacion
    $script:timer.Start()
})

# --- ANIMACION ---
$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds($frameMs)

$script:timer.Add_Tick({
    $script:frameCount++
    $elapsed = ([DateTime]::Now - $script:startTime).TotalSeconds

    for ($c = 0; $c -lt $numRainCols; $c++) {
        $col = $script:columns[$c]

        # Avanzar cabeza
        $col.HeadY += $col.Speed

        # Si la cola paso la pantalla, reiniciar
        if (($col.HeadY - $col.TrailPx) -gt $script:screenH) {
            $col.HeadY = -(Get-Random -Minimum 100 -Maximum ([int]($script:screenH * 0.7)))
            $col.Speed = (Get-Random -Minimum 15 -Maximum 55) / 10.0
            $col.TrailLength = Get-Random -Minimum 8 -Maximum 28
            $col.TrailPx = $col.TrailLength * $charHeight

            # Reasignar mensaje oculto (15% probabilidad)
            $nRows = [int]($script:screenH / $charHeight) + 2
            $col.HasMessage = $false
            $col.MessageText = ""
            $col.MessageStartRow = -1
            if ((Get-Random -Minimum 0 -Maximum 100) -lt 15) {
                $msg = $medicalTerms[(Get-Random -Minimum 0 -Maximum $medicalTerms.Count)]
                $maxStart = $nRows - $msg.Length - 2
                if ($maxStart -gt 2) {
                    $col.HasMessage = $true
                    $col.MessageText = $msg
                    $col.MessageStartRow = (Get-Random -Minimum 2 -Maximum $maxStart)
                    # Inyectar en el texto
                    $lines = $script:colTextBlocks[$c].Text.Split("`n")
                    for ($mi = 0; $mi -lt $msg.Length; $mi++) {
                        $rowIdx = $col.MessageStartRow + $mi
                        if ($rowIdx -ge 0 -and $rowIdx -lt $lines.Count) {
                            $lines[$rowIdx] = [string]$msg[$mi]
                        }
                    }
                    $script:colTextBlocks[$c].Text = $lines -join "`n"
                }
            }
        }

        # Cambiar caracteres aleatoriamente
        $col.CharChangeTimer--
        if ($col.CharChangeTimer -le 0) {
            $col.CharChangeTimer = Get-Random -Minimum 3 -Maximum 12
            $numRows = [int]($script:screenH / $charHeight) + 2
            # Cambiar algunos caracteres
            $text = $script:colTextBlocks[$c].Text
            $lines = $text.Split("`n")
            $changeCount = Get-Random -Minimum 1 -Maximum 4
            for ($ch = 0; $ch -lt $changeCount; $ch++) {
                $ri = Get-Random -Minimum 0 -Maximum $lines.Count
                # No sobreescribir letras de mensajes ocultos
                if ($col.HasMessage -and $ri -ge $col.MessageStartRow -and $ri -lt ($col.MessageStartRow + $col.MessageText.Length)) { continue }
                $lines[$ri] = [string](Get-RandomChar)
            }
            $script:colTextBlocks[$c].Text = $lines -join "`n"
        }

        # Actualizar gradiente
        $headNorm = $col.HeadY / $script:screenH
        $tailNorm = ($col.HeadY - $col.TrailPx) / $script:screenH

        if ($headNorm -gt 0 -and $tailNorm -lt 1) {
            $script:colTextBlocks[$c].Foreground = (New-RainGradient -headNorm $headNorm -tailNorm $tailNorm)
        } else {
            $script:colTextBlocks[$c].Foreground = [System.Windows.Media.Brushes]::Transparent
        }
    }

    # --- Titulo: fade in despues de 3 segundos ---
    if ($elapsed -gt 3 -and $elapsed -lt 6) {
        $fadeProgress = ($elapsed - 3) / 3.0
        if ($fadeProgress -gt 1) { $fadeProgress = 1 }
        $script:title1.Opacity = $fadeProgress
        $script:title2.Opacity = $fadeProgress
        $script:titleSub.Opacity = $fadeProgress * 0.6
    }
    elseif ($elapsed -ge 6) {
        # Pulse del glow del titulo
        $pulse = [Math]::Sin($elapsed * 1.5) * 0.15 + 0.85
        $script:title1.Opacity = $pulse
        $script:title2.Opacity = $pulse
        $script:titleSub.Opacity = $pulse * 0.6

        # Variar el glow
        $glowSize = 25 + [Math]::Sin($elapsed * 2) * 10
        $script:title1.Effect.BlurRadius = $glowSize
        $script:title2.Effect.BlurRadius = $glowSize
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
