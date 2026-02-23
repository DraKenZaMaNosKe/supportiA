using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

class MatrixScreenSaver
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
                        Arguments = "-ExecutionPolicy Bypass -NoProfile -STA -File \"C:\\Windows\\System32\\HCG_Matrix.ps1\"",
                        UseShellExecute = false
                    };
                    var proc = Process.Start(psi);
                    proc.WaitForExit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "HCG Matrix",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                break;

            case "/c":
                MessageBox.Show(
                    "Hospital Civil de Guadalajara\n" +
                    "Protector de Pantalla Matrix\n\n" +
                    "No hay opciones configurables.",
                    "HCG Matrix", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;

            case "/p":
                break;
        }
    }
}
