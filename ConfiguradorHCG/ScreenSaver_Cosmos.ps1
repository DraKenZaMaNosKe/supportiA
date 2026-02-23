# =============================================================================
# HCG - PROTECTOR DE PANTALLA COSMOS
# =============================================================================
# Starfield + particulas cosmicas
# Presiona cualquier tecla o mueve el mouse para salir
# =============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Configuracion ---
$NumStars = 150
$Speed = 4
$MaxDepth = 800

# --- Ventana fullscreen sin bordes ---
$window = New-Object System.Windows.Window
$window.WindowStyle = 'None'
$window.WindowState = 'Maximized'
$window.Background = [System.Windows.Media.Brushes]::Black
$window.Topmost = $true
$window.Cursor = [System.Windows.Input.Cursors]::None
$window.AllowsTransparency = $false

$canvas = New-Object System.Windows.Controls.Canvas
$canvas.Background = [System.Windows.Media.Brushes]::Black
$window.Content = $canvas

# --- Texto sutil en la esquina ---
$label = New-Object System.Windows.Controls.TextBlock
$label.Text = "Hospital Civil de Guadalajara"
$label.Foreground = New-Object System.Windows.Media.SolidColorBrush(
    [System.Windows.Media.Color]::FromArgb(40, 255, 255, 255)
)
$label.FontSize = 14
$label.FontFamily = New-Object System.Windows.Media.FontFamily("Segoe UI Light")
$canvas.Children.Add($label) | Out-Null

# --- Variables globales ---
$script:screenW = 0
$script:screenH = 0
$script:centerX = 0
$script:centerY = 0
$script:mouseStartX = -1
$script:mouseStartY = -1
$script:startTime = [DateTime]::Now
$script:stars = @()
$script:ellipses = @()

# --- Colores de estrellas (blanco, azul claro, azul, calido) ---
$script:starColors = @(
    [System.Windows.Media.Color]::FromRgb(255, 255, 255),
    [System.Windows.Media.Color]::FromRgb(200, 220, 255),
    [System.Windows.Media.Color]::FromRgb(180, 200, 255),
    [System.Windows.Media.Color]::FromRgb(255, 240, 220),
    [System.Windows.Media.Color]::FromRgb(160, 180, 255)
)

function New-Star {
    return @{
        X = (Get-Random -Minimum -600 -Maximum 600) * 1.0
        Y = (Get-Random -Minimum -400 -Maximum 400) * 1.0
        Z = (Get-Random -Minimum 10 -Maximum $MaxDepth) * 1.0
        ColorIndex = (Get-Random -Minimum 0 -Maximum $script:starColors.Count)
        BaseSize = (Get-Random -Minimum 15 -Maximum 40) / 10.0
    }
}

# --- Inicializar estrellas ---
$window.Add_Loaded({
    $script:screenW = $canvas.ActualWidth
    $script:screenH = $canvas.ActualHeight
    $script:centerX = $script:screenW / 2
    $script:centerY = $script:screenH / 2

    # Posicionar texto
    [System.Windows.Controls.Canvas]::SetLeft($label, 20)
    [System.Windows.Controls.Canvas]::SetTop($label, $script:screenH - 40)

    # Crear estrellas
    for ($i = 0; $i -lt $NumStars; $i++) {
        $script:stars += (New-Star)

        $ellipse = New-Object System.Windows.Shapes.Ellipse
        $ellipse.Width = 2
        $ellipse.Height = 2
        $ellipse.Fill = [System.Windows.Media.Brushes]::White
        $canvas.Children.Add($ellipse) | Out-Null
        $script:ellipses += $ellipse
    }

    # Iniciar animacion
    $script:timer.Start()
})

# --- Timer de animacion (~30fps) ---
$script:timer = New-Object System.Windows.Threading.DispatcherTimer
$script:timer.Interval = [TimeSpan]::FromMilliseconds(33)

$script:timer.Add_Tick({
    for ($i = 0; $i -lt $NumStars; $i++) {
        # Mover estrella hacia el espectador
        $script:stars[$i].Z -= $Speed

        # Si la estrella pasa, reiniciar
        if ($script:stars[$i].Z -le 1) {
            $script:stars[$i].X = (Get-Random -Minimum -600 -Maximum 600) * 1.0
            $script:stars[$i].Y = (Get-Random -Minimum -400 -Maximum 400) * 1.0
            $script:stars[$i].Z = $MaxDepth * 1.0
        }

        $z = $script:stars[$i].Z

        # Proyeccion 3D a 2D
        $factor = 300.0 / $z
        $sx = $script:stars[$i].X * $factor + $script:centerX
        $sy = $script:stars[$i].Y * $factor + $script:centerY

        # Tamano segun profundidad
        $depthRatio = 1.0 - ($z / $MaxDepth)
        $size = $script:stars[$i].BaseSize * $depthRatio * 4 + 0.5

        # Opacidad segun profundidad
        $opacity = $depthRatio * $depthRatio
        if ($opacity -gt 1) { $opacity = 1 }
        if ($opacity -lt 0.05) { $opacity = 0.05 }

        # Actualizar visual
        $script:ellipses[$i].Width = $size
        $script:ellipses[$i].Height = $size
        $script:ellipses[$i].Opacity = $opacity

        # Color
        $c = $script:starColors[$script:stars[$i].ColorIndex]
        $script:ellipses[$i].Fill = New-Object System.Windows.Media.SolidColorBrush($c)

        [System.Windows.Controls.Canvas]::SetLeft($script:ellipses[$i], $sx - $size / 2)
        [System.Windows.Controls.Canvas]::SetTop($script:ellipses[$i], $sy - $size / 2)
    }
})

# --- Salir con tecla ---
$window.Add_KeyDown({
    $script:timer.Stop()
    $window.Close()
})

# --- Salir con mouse (con gracia de 1 segundo y 10px de movimiento) ---
$window.Add_MouseMove({
    param($s, $e)
    $pos = $e.GetPosition($window)

    if ($script:mouseStartX -eq -1) {
        $script:mouseStartX = $pos.X
        $script:mouseStartY = $pos.Y
        return
    }

    $elapsed = ([DateTime]::Now - $script:startTime).TotalSeconds
    if ($elapsed -lt 1.5) { return }

    $dx = [Math]::Abs($pos.X - $script:mouseStartX)
    $dy = [Math]::Abs($pos.Y - $script:mouseStartY)

    if ($dx -gt 10 -or $dy -gt 10) {
        $script:timer.Stop()
        $window.Close()
    }
})

$window.Add_MouseDown({
    $script:timer.Stop()
    $window.Close()
})

# --- Mostrar ---
$app = [System.Windows.Application]::new()
$app.Run($window) | Out-Null
