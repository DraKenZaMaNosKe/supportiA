using System.Media;

namespace iAUptowin.Services;

/// <summary>
/// Servicio de audio con melodías de Saint Seiya
/// Usa hilos separados para que el audio no se interrumpa durante procesos largos
/// </summary>
public class AudioService : IDisposable
{
    private Thread? _melodyThread;
    private CancellationTokenSource? _melodyCts;
    private bool _isPlaying = false;
    private readonly object _lock = new object();

    // Notas musicales (frecuencias en Hz)
    private static class Notes
    {
        public const int C4 = 262;  // Do4
        public const int D4 = 294;  // Re4
        public const int E4 = 330;  // Mi4
        public const int F4 = 349;  // Fa4
        public const int G4 = 392;  // Sol4
        public const int A4 = 440;  // La4
        public const int B4 = 494;  // Si4
        public const int C5 = 523;  // Do5
        public const int D5 = 587;  // Re5
        public const int E5 = 659;  // Mi5
        public const int G5 = 784;  // Sol5
        public const int A5 = 880;  // La5
    }

    /// <summary>
    /// Inicia la melodía ambiental del Santuario en un hilo separado
    /// </summary>
    public void StartBackgroundMelody()
    {
        lock (_lock)
        {
            if (_isPlaying) return;
            _isPlaying = true;
        }

        _melodyCts = new CancellationTokenSource();
        var token = _melodyCts.Token;

        _melodyThread = new Thread(() => PlaySanctuaryMelody(token))
        {
            IsBackground = true,
            Priority = ThreadPriority.BelowNormal
        };
        _melodyThread.Start();
    }

    /// <summary>
    /// Detiene la melodía de fondo
    /// </summary>
    public void StopBackgroundMelody()
    {
        lock (_lock)
        {
            if (!_isPlaying) return;
            _isPlaying = false;
        }

        _melodyCts?.Cancel();
        _melodyThread?.Join(1000);
        _melodyCts?.Dispose();
        _melodyCts = null;
        _melodyThread = null;
    }

    /// <summary>
    /// Melodía del Santuario - estilo contemplativo de Saint Seiya
    /// </summary>
    private void PlaySanctuaryMelody(CancellationToken token)
    {
        // Secuencia de notas: Do -> Mi -> Sol -> La -> Sol -> Fa -> Mi -> Re -> Do
        var melody = new (int frequency, int duration)[]
        {
            (Notes.C4, 1500),
            (Notes.E4, 1500),
            (Notes.G4, 1500),
            (Notes.A4, 2000),
            (Notes.G4, 1500),
            (Notes.F4, 1500),
            (Notes.E4, 1500),
            (Notes.D4, 1500),
            (Notes.C4, 2000),
            (Notes.D4, 1200),
            (Notes.E4, 1200),
        };

        while (!token.IsCancellationRequested)
        {
            foreach (var (frequency, duration) in melody)
            {
                if (token.IsCancellationRequested) break;

                try
                {
                    Console.Beep(frequency, duration);
                }
                catch
                {
                    // Silencioso si falla el beep
                }

                // Pequeña pausa entre notas
                if (!token.IsCancellationRequested)
                {
                    Thread.Sleep(100);
                }
            }
        }
    }

    /// <summary>
    /// Toca un sonido corto de éxito (nota de Pegasus Fantasy)
    /// </summary>
    public void PlaySuccessSound()
    {
        Task.Run(() =>
        {
            try
            {
                Console.Beep(Notes.E5, 150);
                Console.Beep(Notes.G5, 200);
            }
            catch { }
        });
    }

    /// <summary>
    /// Toca un sonido corto de error
    /// </summary>
    public void PlayErrorSound()
    {
        Task.Run(() =>
        {
            try
            {
                Console.Beep(Notes.E4, 150);
                Console.Beep(Notes.C4, 200);
            }
            catch { }
        });
    }

    /// <summary>
    /// Melodía de victoria épica - Pegasus Fantasy
    /// </summary>
    public void PlayVictorySound()
    {
        Task.Run(() =>
        {
            try
            {
                // Pegasus Fantasy intro
                var victory = new (int freq, int dur)[]
                {
                    (Notes.E5, 150), (Notes.E5, 150), (Notes.D5, 150),
                    (Notes.E5, 250), (Notes.G5, 150), (Notes.G5, 150),
                    (Notes.E5, 250), (Notes.D5, 150), (Notes.C5, 150),
                    (Notes.B4, 250), (Notes.A4, 150), (Notes.B4, 150),
                    (Notes.C5, 150), (Notes.D5, 150), (Notes.E5, 400),
                    (Notes.G5, 200), (Notes.A5, 400)
                };

                foreach (var (freq, dur) in victory)
                {
                    Console.Beep(freq, dur);
                    Thread.Sleep(30);
                }
            }
            catch { }
        });
    }

    /// <summary>
    /// Melodía triste (para actualización a Windows 11)
    /// </summary>
    public void PlaySadMelody()
    {
        Task.Run(() =>
        {
            try
            {
                var sad = new (int freq, int dur)[]
                {
                    (Notes.A4, 600), (Notes.G4, 600), (Notes.F4, 800),
                    (Notes.E4, 400), (Notes.D4, 600), (Notes.C4, 1000),
                    (Notes.E4, 500), (Notes.F4, 500), (Notes.G4, 700),
                    (Notes.A4, 600), (Notes.G4, 400), (Notes.F4, 800),
                    (Notes.C5, 800), (Notes.B4, 500), (Notes.A4, 600),
                    (Notes.G4, 700), (Notes.F4, 500), (Notes.E4, 900),
                    (Notes.D4, 600), (Notes.C4, 1200)
                };

                foreach (var (freq, dur) in sad)
                {
                    Console.Beep(freq, dur);
                    Thread.Sleep(50);
                }
            }
            catch { }
        });
    }

    public void Dispose()
    {
        StopBackgroundMelody();
    }
}
