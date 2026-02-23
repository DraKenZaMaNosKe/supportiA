using System.ComponentModel;
using iAUptowin.Services;

namespace iAUptowin;

public partial class MainForm : Form
{
    private readonly AudioService _audioService;
    private readonly ConfigService _configService;
    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isRunning = false;

    // Controles de la interfaz
    private Panel panelHeader = null!;
    private Label lblTitle = null!;
    private Label lblSubtitle = null!;
    private TextBox txtInventario = null!;
    private CheckBox chkOPD = null!;
    private Button btnIniciar = null!;
    private Button btnCancelar = null!;
    private ProgressBar progressBar = null!;
    private Label lblProgreso = null!;
    private ListBox lstPasos = null!;
    private RichTextBox txtLog = null!;
    private Label lblEstado = null!;

    public MainForm()
    {
        InitializeComponent();
        _audioService = new AudioService();
        _configService = new ConfigService();
        SetupEventHandlers();
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();

        // Configuración del formulario
        this.Text = "iA Uptowin - Configurador Cósmico HCG";
        this.Size = new Size(900, 700);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.FromArgb(20, 20, 35);
        this.ForeColor = Color.White;
        this.Font = new Font("Segoe UI", 10F);
        this.FormBorderStyle = FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;

        // Panel Header con gradiente
        panelHeader = new Panel
        {
            Dock = DockStyle.Top,
            Height = 100,
            BackColor = Color.FromArgb(30, 30, 50)
        };
        panelHeader.Paint += PanelHeader_Paint;

        // Título
        lblTitle = new Label
        {
            Text = "★ CONFIGURADOR CÓSMICO ★",
            Font = new Font("Segoe UI", 24F, FontStyle.Bold),
            ForeColor = Color.FromArgb(255, 215, 0), // Dorado
            AutoSize = true,
            Location = new Point(20, 15),
            BackColor = Color.Transparent
        };

        // Subtítulo
        lblSubtitle = new Label
        {
            Text = "Hospital Civil de Guadalajara - Los Caballeros de Informática",
            Font = new Font("Segoe UI", 11F, FontStyle.Italic),
            ForeColor = Color.FromArgb(100, 200, 255), // Cyan
            AutoSize = true,
            Location = new Point(22, 55),
            BackColor = Color.Transparent
        };

        panelHeader.Controls.AddRange(new Control[] { lblTitle, lblSubtitle });

        // Panel de configuración
        var panelConfig = new Panel
        {
            Location = new Point(20, 115),
            Size = new Size(400, 100),
            BackColor = Color.FromArgb(35, 35, 55)
        };

        var lblInventario = new Label
        {
            Text = "Número de Inventario:",
            Location = new Point(15, 15),
            AutoSize = true,
            ForeColor = Color.White
        };

        txtInventario = new TextBox
        {
            Location = new Point(15, 40),
            Size = new Size(200, 30),
            Font = new Font("Consolas", 14F, FontStyle.Bold),
            BackColor = Color.FromArgb(50, 50, 70),
            ForeColor = Color.FromArgb(0, 255, 200),
            BorderStyle = BorderStyle.FixedSingle
        };

        chkOPD = new CheckBox
        {
            Text = "Es equipo OPD",
            Location = new Point(230, 42),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 200, 100)
        };

        panelConfig.Controls.AddRange(new Control[] { lblInventario, txtInventario, chkOPD });

        // Botones
        btnIniciar = new Button
        {
            Text = "★ INICIAR CONFIGURACIÓN",
            Location = new Point(440, 130),
            Size = new Size(220, 45),
            BackColor = Color.FromArgb(0, 120, 80),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 11F, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnIniciar.FlatAppearance.BorderColor = Color.FromArgb(0, 200, 120);

        btnCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(680, 130),
            Size = new Size(100, 45),
            BackColor = Color.FromArgb(120, 40, 40),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Enabled = false,
            Cursor = Cursors.Hand
        };

        // Barra de progreso
        progressBar = new ProgressBar
        {
            Location = new Point(20, 230),
            Size = new Size(840, 30),
            Style = ProgressBarStyle.Continuous,
            Maximum = 100
        };

        lblProgreso = new Label
        {
            Text = "Esperando inicio...",
            Location = new Point(20, 265),
            Size = new Size(840, 25),
            ForeColor = Color.FromArgb(150, 150, 150)
        };

        // Lista de pasos
        var lblPasos = new Label
        {
            Text = "Pasos de Configuración:",
            Location = new Point(20, 295),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 215, 0)
        };

        lstPasos = new ListBox
        {
            Location = new Point(20, 320),
            Size = new Size(350, 320),
            BackColor = Color.FromArgb(25, 25, 40),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Consolas", 9F)
        };

        // Inicializar lista de pasos
        string[] pasos = {
            "○ Conectar al servidor",
            "○ Eliminar Office previo",
            "○ Crear usuario soporte",
            "○ Crear usuario equipo",
            "○ Configurar auto-login",
            "○ Instalar WinRAR",
            "○ Instalar .NET 3.5",
            "○ Instalar Acrobat Reader",
            "○ Instalar navegadores",
            "○ Instalar Office 2007",
            "○ Instalar Dedalus",
            "○ Configurar sincronizador",
            "○ Instalar antivirus ESET",
            "○ Copiar accesos directos",
            "○ Limpiar iconos duplicados",
            "○ Habilitar Windows Update",
            "○ Quitar admin a usuario actual",
            "○ Instalar reportes automáticos",
            "○ Renombrar equipo",
            "○ Enviar datos a Google Sheets"
        };
        lstPasos.Items.AddRange(pasos);

        // Log
        var lblLog = new Label
        {
            Text = "Log de Actividad:",
            Location = new Point(390, 295),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 215, 0)
        };

        txtLog = new RichTextBox
        {
            Location = new Point(390, 320),
            Size = new Size(470, 280),
            BackColor = Color.FromArgb(15, 15, 25),
            ForeColor = Color.FromArgb(0, 255, 150),
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Consolas", 9F),
            ReadOnly = true
        };

        // Estado
        lblEstado = new Label
        {
            Text = "★ Los Caballeros de Informática protegen este equipo ★",
            Location = new Point(20, 650),
            Size = new Size(840, 25),
            ForeColor = Color.FromArgb(200, 100, 255),
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 9F, FontStyle.Italic)
        };

        // Agregar controles al formulario
        this.Controls.AddRange(new Control[] {
            panelHeader,
            panelConfig,
            btnIniciar,
            btnCancelar,
            progressBar,
            lblProgreso,
            lblPasos,
            lstPasos,
            lblLog,
            txtLog,
            lblEstado
        });

        this.ResumeLayout(false);
    }

    private void PanelHeader_Paint(object? sender, PaintEventArgs e)
    {
        // Dibujar estrellas decorativas
        var starPoints = new[] {
            new Point(750, 20), new Point(800, 35), new Point(830, 60),
            new Point(780, 75), new Point(720, 50)
        };

        using var starBrush = new SolidBrush(Color.FromArgb(80, 255, 215, 0));
        foreach (var point in starPoints)
        {
            e.Graphics.FillEllipse(starBrush, point.X, point.Y, 8, 8);
        }
    }

    private void SetupEventHandlers()
    {
        btnIniciar.Click += BtnIniciar_Click;
        btnCancelar.Click += BtnCancelar_Click;
        this.FormClosing += MainForm_FormClosing;
    }

    private async void BtnIniciar_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtInventario.Text))
        {
            MessageBox.Show("Por favor ingresa el número de inventario", "Atención",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            txtInventario.Focus();
            return;
        }

        _isRunning = true;
        _cancellationTokenSource = new CancellationTokenSource();

        // Actualizar UI
        btnIniciar.Enabled = false;
        btnCancelar.Enabled = true;
        txtInventario.Enabled = false;
        chkOPD.Enabled = false;

        Log("═══════════════════════════════════════════════════", Color.Gold);
        Log("★ INICIANDO CONFIGURACIÓN CÓSMICA ★", Color.Gold);
        Log($"  Inventario: {txtInventario.Text}", Color.Cyan);
        Log($"  Es OPD: {(chkOPD.Checked ? "Sí" : "No")}", Color.Cyan);
        Log("═══════════════════════════════════════════════════", Color.Gold);

        try
        {
            await _configService.EjecutarConfiguracion(
                txtInventario.Text,
                chkOPD.Checked,
                new Progress<ConfigProgress>(UpdateProgress),
                _cancellationTokenSource.Token
            );

            // Victoria!
            _audioService.PlayVictorySound();
            Log("", Color.White);
            Log("★★★ CONFIGURACIÓN COMPLETADA CON ÉXITO ★★★", Color.Lime);
            MessageBox.Show("¡Configuración cósmica completada!\n\nEl equipo está listo.",
                "Victoria", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (OperationCanceledException)
        {
            Log("⚠ Configuración cancelada por el usuario", Color.Orange);
        }
        catch (Exception ex)
        {
            Log($"✖ Error: {ex.Message}", Color.Red);
            MessageBox.Show($"Error durante la configuración:\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _isRunning = false;
            btnIniciar.Enabled = true;
            btnCancelar.Enabled = false;
            txtInventario.Enabled = true;
            chkOPD.Enabled = true;
            _audioService.StopBackgroundMelody();
        }
    }

    private void BtnCancelar_Click(object? sender, EventArgs e)
    {
        if (_cancellationTokenSource != null)
        {
            _cancellationTokenSource.Cancel();
            Log("Cancelando operación...", Color.Orange);
        }
    }

    private void UpdateProgress(ConfigProgress progress)
    {
        if (InvokeRequired)
        {
            Invoke(() => UpdateProgress(progress));
            return;
        }

        progressBar.Value = progress.Percentage;
        lblProgreso.Text = $"[{progress.Percentage}%] {progress.CurrentStep}";

        if (progress.StepIndex >= 0 && progress.StepIndex < lstPasos.Items.Count)
        {
            // Actualizar el paso en la lista
            var status = progress.StepCompleted ? "✓" : "●";
            var color = progress.StepCompleted ? "Verde" : "Amarillo";
            lstPasos.Items[progress.StepIndex] = $"{status} {progress.StepName}";
            lstPasos.SelectedIndex = progress.StepIndex;
        }

        if (!string.IsNullOrEmpty(progress.LogMessage))
        {
            Log(progress.LogMessage, progress.LogColor);
        }

        // Control de melodía de fondo
        if (progress.StartMelody)
        {
            _audioService.StartBackgroundMelody();
        }
        else if (progress.StopMelody)
        {
            _audioService.StopBackgroundMelody();
        }
    }

    private void Log(string message, Color color)
    {
        if (InvokeRequired)
        {
            Invoke(() => Log(message, color));
            return;
        }

        txtLog.SelectionStart = txtLog.TextLength;
        txtLog.SelectionColor = color;
        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
        txtLog.ScrollToCaret();
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_isRunning)
        {
            var result = MessageBox.Show(
                "Hay una configuración en progreso. ¿Deseas cancelar y salir?",
                "Confirmar salida",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }

            _cancellationTokenSource?.Cancel();
        }

        _audioService.Dispose();
    }
}

// Clase para reportar progreso
public class ConfigProgress
{
    public int Percentage { get; set; }
    public string CurrentStep { get; set; } = "";
    public int StepIndex { get; set; } = -1;
    public string StepName { get; set; } = "";
    public bool StepCompleted { get; set; }
    public string LogMessage { get; set; } = "";
    public Color LogColor { get; set; } = Color.White;
    public bool StartMelody { get; set; }
    public bool StopMelody { get; set; }
}
