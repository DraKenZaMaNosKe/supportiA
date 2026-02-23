using System.Diagnostics;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;

namespace iAUptowin.Services;

/// <summary>
/// Servicio principal de configuración de equipos HCG
/// Migración completa de ConfigurarEquipoHCG.ps1
/// </summary>
public class ConfigService
{
    // Configuración del servidor
    private const string Servidor = "10.2.1.13";
    private const string Usuario = "2010201";
    private const string Password = "7v3l73v37nG06";

    // Rutas del servidor
    private const string RutaBase = @"\\10.2.1.13\soportefaa\pack_installer_iA";
    private static readonly string RutaAccesos = $@"{RutaBase}\accesos_directos";
    private static readonly string RutaAcrobat = $@"{RutaBase}\acrobat_reader";
    private static readonly string RutaAntivirus = $@"{RutaBase}\antivirus";
    private static readonly string RutaChrome = $@"{RutaBase}\chrome";
    private static readonly string RutaOffice = $@"{RutaBase}\office";
    private static readonly string RutaWallpaper = $@"{RutaBase}\wallpaper";
    private static readonly string RutaDotNet = $@"{RutaBase}\netframework3_5\sxs";
    private static readonly string RutaWinRAR = $@"{RutaBase}\winrar_licence";
    private static readonly string RutaDedalus = @"\\10.2.1.17\distribucion\dedalus";

    // Configuración de usuarios
    private const string UsuarioSoporte = "Soporte";
    private const string PasswordSoporte = "*TIsoporte";
    private const string RutaLogs = @"C:\HCG_Logs";

    // Google Sheets
    private const string GoogleSheetURL = "https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec";

    // Estado
    private List<string> _softwareInstalado = new();
    private string _usuarioOriginal = Environment.UserName;
    private bool _esOPD = false;
    private int _filaRegistro = 0;

    private readonly HttpClient _httpClient;

    public ConfigService()
    {
        _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
    }

    /// <summary>
    /// Ejecuta la configuración completa del equipo
    /// </summary>
    public async Task EjecutarConfiguracion(
        string numInventario,
        bool esOPD,
        IProgress<ConfigProgress> progress,
        CancellationToken cancellationToken)
    {
        _esOPD = esOPD;
        _softwareInstalado.Clear();

        // Obtener datos del equipo
        var datosEquipo = GetDatosEquipo();

        var pasos = new List<(string nombre, Func<string, IProgress<ConfigProgress>, Task<bool>> accion, bool usaMelodia)>
        {
            ("Registrar en Google Sheets", (inv, p) => SendDatosInicio(inv, datosEquipo, p), false),
            ("Conectar al servidor", (inv, p) => ConectarServidor(p), false),
            ("Eliminar Office previo", (inv, p) => RemoverOfficePrevio(p), true),
            ("Crear usuario soporte", (inv, p) => CrearUsuarioSoporte(p), false),
            ("Crear usuario equipo", (inv, p) => CrearUsuarioEquipo(inv, p), false),
            ("Configurar red privada", (inv, p) => ConfigurarRedPrivada(p), false),
            ("Configurar hora automática", (inv, p) => ConfigurarHoraAutomatica(p), false),
            ("Instalar WinRAR", (inv, p) => InstalarWinRAR(p), false),
            ("Instalar .NET 3.5", (inv, p) => InstalarDotNet35(p), true),
            ("Instalar Acrobat Reader", (inv, p) => InstalarAcrobat(p), false),
            ("Instalar Chrome", (inv, p) => InstalarChrome(p), false),
            ("Copiar fondo de pantalla", (inv, p) => CopiarFondoPantalla(p), false),
            ("Instalar Office 2007", (inv, p) => InstalarOffice(p), true),
            ("Instalar Dedalus", (inv, p) => InstalarDedalus(p), true),
            ("Configurar sincronizador Dedalus", (inv, p) => ConfigurarSincronizadorDedalus(p), false),
            ("Instalar antivirus ESET", (inv, p) => InstalarAntivirus(p), true),
            ("Copiar accesos directos", (inv, p) => CopiarAccesosDirectos(p), false),
            ("Limpiar iconos duplicados", (inv, p) => LimpiarIconosDuplicados(p), false),
            ("Habilitar Windows Update", (inv, p) => HabilitarWindowsUpdate(p), false),
            ("Quitar admin a usuario actual", (inv, p) => QuitarAdminUsuarioActual(p), false),
            ("Instalar reporte de IP", (inv, p) => InstalarReporteIP(p), false),
            ("Instalar reporte de sistema", (inv, p) => InstalarReporteSistema(p), false),
            ("Instalar reporte diagnóstico", (inv, p) => InstalarReporteDiagnostico(p), false),
            ("Renombrar equipo", (inv, p) => RenombrarEquipo(inv, p), false),
            ("Actualizar Google Sheets", (inv, p) => SendDatosFin(inv, p), false),
            ("Registrar software", (inv, p) => SendSoftwareInfo(inv, p), false),
        };

        for (int i = 0; i < pasos.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var (nombre, accion, usaMelodia) = pasos[i];
            int porcentaje = (int)((i + 1) * 100.0 / pasos.Count);

            progress.Report(new ConfigProgress
            {
                Percentage = porcentaje,
                CurrentStep = nombre,
                StepIndex = i,
                StepName = nombre,
                StepCompleted = false,
                LogMessage = $"→ Iniciando: {nombre}...",
                LogColor = Color.Cyan,
                StartMelody = usaMelodia
            });

            try
            {
                bool exito = await accion(numInventario, progress);

                progress.Report(new ConfigProgress
                {
                    Percentage = porcentaje,
                    CurrentStep = nombre,
                    StepIndex = i,
                    StepName = nombre,
                    StepCompleted = true,
                    LogMessage = exito ? $"★ {nombre} completado" : $"⚠ {nombre} con advertencias",
                    LogColor = exito ? Color.Lime : Color.Orange,
                    StopMelody = usaMelodia
                });
            }
            catch (Exception ex)
            {
                progress.Report(new ConfigProgress
                {
                    Percentage = porcentaje,
                    CurrentStep = nombre,
                    StepIndex = i,
                    StepName = nombre,
                    StepCompleted = false,
                    LogMessage = $"✖ Error en {nombre}: {ex.Message}",
                    LogColor = Color.Red,
                    StopMelody = true
                });
            }

            await Task.Delay(300, cancellationToken);
        }
    }

    #region Datos del Equipo

    private Dictionary<string, string> GetDatosEquipo()
    {
        var datos = new Dictionary<string, string>
        {
            ["Serie"] = "",
            ["MACEthernet"] = "",
            ["MACWiFi"] = "",
            ["ProductKey"] = ""
        };

        try
        {
            // Número de serie
            var biosQuery = RunWmiQuery("SELECT SerialNumber FROM Win32_BIOS");
            if (biosQuery.Count > 0)
                datos["Serie"] = biosQuery[0].GetValueOrDefault("SerialNumber", "");

            // MAC Ethernet
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet &&
                    !nic.Description.Contains("Virtual") &&
                    nic.GetPhysicalAddress().ToString().Length == 12)
                {
                    datos["MACEthernet"] = nic.GetPhysicalAddress().ToString();
                    break;
                }
            }

            // MAC WiFi
            foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 &&
                    nic.GetPhysicalAddress().ToString().Length == 12)
                {
                    datos["MACWiFi"] = nic.GetPhysicalAddress().ToString();
                    break;
                }
            }

            // Product Key
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform");
                datos["ProductKey"] = key?.GetValue("BackupProductKeyDefault")?.ToString() ?? "";
            }
            catch { }
        }
        catch { }

        return datos;
    }

    #endregion

    #region Google Sheets

    private async Task<bool> SendDatosInicio(string invST, Dictionary<string, string> datos, IProgress<ConfigProgress> progress)
    {
        try
        {
            Log(progress, "Enviando datos a Google Sheets...", Color.Cyan);

            var cpuQuery = RunWmiQuery("SELECT Name, NumberOfCores FROM Win32_Processor");
            var cpuName = cpuQuery.Count > 0 ? cpuQuery[0].GetValueOrDefault("Name", "").Replace("(R)", "").Replace("(TM)", "").Trim() : "";
            var nucleos = cpuQuery.Count > 0 ? cpuQuery[0].GetValueOrDefault("NumberOfCores", "4") : "4";

            var csQuery = RunWmiQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
            var ramBytes = csQuery.Count > 0 ? long.Parse(csQuery[0].GetValueOrDefault("TotalPhysicalMemory", "0")) : 0;
            var ramGB = Math.Round(ramBytes / 1024.0 / 1024.0 / 1024.0, 0);

            var diskQuery = RunWmiQuery("SELECT Size, Model FROM Win32_DiskDrive");
            var discoGB = 0L;
            var discoTipo = "HDD";
            if (diskQuery.Count > 0)
            {
                discoGB = (long)Math.Round(long.Parse(diskQuery[0].GetValueOrDefault("Size", "0")) / 1024.0 / 1024.0 / 1024.0, 0);
                var model = diskQuery[0].GetValueOrDefault("Model", "");
                if (model.Contains("SSD") || model.Contains("NVMe"))
                    discoTipo = "SSD";
            }

            var body = new
            {
                Accion = "crear",
                Fecha = DateTime.Now.ToString("dd/MM/yyyy"),
                InvST = invST,
                Serie = datos["Serie"],
                Marca = "Lenovo",
                Modelo = "ThinkCentre M70s Gen 5",
                Procesador = cpuName,
                Nucleos = nucleos,
                RAM = ramGB,
                Disco = discoGB,
                DiscoTipo = discoTipo,
                MACEthernet = datos["MACEthernet"],
                MACWiFi = datos["MACWiFi"],
                ProductKey = datos["ProductKey"],
                FechaFab = DateTime.Now.ToString("dd/MM/yyyy"),
                Garantia = DateTime.Now.AddYears(3).ToString("dd/MM/yyyy")
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(GoogleSheetURL, content);
            var responseText = await response.Content.ReadAsStringAsync();

            if (responseText.Contains("\"status\":\"OK\""))
            {
                // Parsear respuesta JSON para extraer fila y FAA
                try
                {
                    using var doc = JsonDocument.Parse(responseText);
                    var root = doc.RootElement;

                    // Extraer número de fila
                    if (root.TryGetProperty("row", out var rowElement))
                    {
                        _filaRegistro = rowElement.GetInt32();
                    }

                    // Extraer FAA
                    var faaStatus = "";
                    if (root.TryGetProperty("faa", out var faaElement))
                    {
                        faaStatus = faaElement.GetString() ?? "";
                    }

                    Log(progress, "═══════════════════════════════════════════", Color.Green);
                    Log(progress, "    EQUIPO REGISTRADO EN SHEETS", Color.Green);
                    Log(progress, "═══════════════════════════════════════════", Color.Green);
                    Log(progress, $"  Inventario:  {invST}", Color.White);
                    Log(progress, $"  No. Serie:   {datos["Serie"]}", Color.White);
                    Log(progress, $"  FILA:        #{_filaRegistro}", Color.Yellow);

                    if (!string.IsNullOrEmpty(faaStatus))
                    {
                        if (faaStatus == "NO ENCONTRADO")
                        {
                            Log(progress, $"  FAA:         {faaStatus}", Color.Red);
                            Log(progress, "", Color.White);
                            Log(progress, "⚠⚠⚠ ALERTA: EQUIPO NO EN LISTA FAA ⚠⚠⚠", Color.Red);
                            Log(progress, "Configurando como equipo OPD...", Color.Orange);
                            _esOPD = true;
                        }
                        else
                        {
                            Log(progress, $"  FAA:         {faaStatus}", Color.Lime);
                        }
                    }
                    Log(progress, "═══════════════════════════════════════════", Color.Green);
                }
                catch
                {
                    Log(progress, $"Equipo registrado (no se pudo leer fila)", Color.Yellow);
                }

                return true;
            }

            return false;
        }
        catch (Exception ex)
        {
            Log(progress, $"Error al registrar: {ex.Message}", Color.Orange);
            return true; // Continuar aunque falle
        }
    }

    private async Task<bool> SendDatosFin(string invST, IProgress<ConfigProgress> progress)
    {
        try
        {
            var softwareList = _softwareInstalado.Count > 0
                ? string.Join(", ", _softwareInstalado)
                : "Configuración completa";

            var body = new
            {
                Accion = "actualizar",
                InvST = invST,
                SoftwareInstalado = softwareList
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            await _httpClient.PostAsync(GoogleSheetURL, content);

            Log(progress, "Equipo marcado como ACTIVO en Google Sheets", Color.Lime);
            return true;
        }
        catch (Exception ex)
        {
            Log(progress, $"Error: {ex.Message}", Color.Orange);
            return true;
        }
    }

    private async Task<bool> SendSoftwareInfo(string invST, IProgress<ConfigProgress> progress)
    {
        try
        {
            var osQuery = RunWmiQuery("SELECT Caption, BuildNumber FROM Win32_OperatingSystem");
            var winVersion = osQuery.Count > 0 ? osQuery[0].GetValueOrDefault("Caption", "").Replace("Microsoft ", "") : "";
            var winBuild = osQuery.Count > 0 ? osQuery[0].GetValueOrDefault("BuildNumber", "") : "";

            var body = new
            {
                Accion = "software",
                InvST = invST,
                NombreEquipo = $"PC-{invST}",
                WindowsVersion = winVersion,
                WindowsBuild = winBuild,
                Office = File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE") ? "2007" : "No",
                Chrome = File.Exists(@"C:\Program Files\Google\Chrome\Application\chrome.exe") ? "Si" : "No",
                Acrobat = File.Exists(@"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe") ? "Si" : "No",
                Dedalus = Directory.Exists(@"C:\Dedalus") ? "Si" : "No",
                ESET = ServiceExists("ekrn") ? "Si" : "No",
                WinRAR = File.Exists(@"C:\Program Files\WinRAR\WinRAR.exe") ? "Si" : "No",
                FechaConfig = DateTime.Now.ToString("dd/MM/yyyy HH:mm")
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            await _httpClient.PostAsync(GoogleSheetURL, content);

            Log(progress, "Inventario de software registrado", Color.Lime);
            return true;
        }
        catch (Exception ex)
        {
            Log(progress, $"Error: {ex.Message}", Color.Orange);
            return true;
        }
    }

    #endregion

    #region Conexión al Servidor

    private async Task<bool> ConectarServidor(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                // Desconectar conexiones previas
                RunCommand("net", $@"use \\{Servidor} /delete /y", waitForExit: true, timeout: 10000);

                // Conectar con credenciales
                var result = RunCommand("net", $@"use \\{Servidor}\soportefaa /user:{Usuario} {Password} /persistent:no", waitForExit: true);

                if (Directory.Exists(RutaBase))
                {
                    Log(progress, $@"Conectado a \\{Servidor}\soportefaa", Color.Lime);
                    return true;
                }

                Log(progress, "Error al conectar al servidor", Color.Red);
                return false;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error de conexión: {ex.Message}", Color.Red);
                return false;
            }
        });
    }

    #endregion

    #region Usuarios

    private async Task<bool> CrearUsuarioSoporte(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var checkResult = RunCommand("net", $"user {UsuarioSoporte}");
                if (!checkResult.Contains("no se ha encontrado") && !checkResult.Contains("not be found"))
                {
                    Log(progress, $"Usuario {UsuarioSoporte} ya existe", Color.Gray);
                }
                else
                {
                    RunCommand("net", $"user {UsuarioSoporte} \"{PasswordSoporte}\" /add /passwordchg:no");
                    Log(progress, $"Usuario {UsuarioSoporte} creado", Color.Lime);
                }

                // Agregar a administradores
                RunCommand("net", $"localgroup Administradores {UsuarioSoporte} /add");
                _softwareInstalado.Add("Usuario Soporte");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> CrearUsuarioEquipo(string numInventario, IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var nombreUsuario = _esOPD ? "OPD" : numInventario;

                var checkResult = RunCommand("net", $"user {nombreUsuario}");
                if (!checkResult.Contains("no se ha encontrado") && !checkResult.Contains("not be found"))
                {
                    Log(progress, $"Usuario {nombreUsuario} ya existe", Color.Gray);
                }
                else
                {
                    RunCommand("net", $"user {nombreUsuario} \"\" /add /passwordchg:no");
                    Log(progress, $"Usuario {nombreUsuario} creado", Color.Lime);
                }

                // Agregar a usuarios normales
                RunCommand("net", $"localgroup Usuarios {nombreUsuario} /add");

                // Quitar de administradores
                RunCommand("net", $"localgroup Administradores {nombreUsuario} /delete");

                // Configurar auto-login
                using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", true);
                if (key != null)
                {
                    key.SetValue("AutoAdminLogon", "1");
                    key.SetValue("DefaultUserName", nombreUsuario);
                    key.SetValue("DefaultPassword", "");
                    key.SetValue("DefaultDomainName", Environment.MachineName);
                }

                Log(progress, $"Auto-login configurado para {nombreUsuario}", Color.Lime);
                _softwareInstalado.Add($"Usuario {nombreUsuario}");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> QuitarAdminUsuarioActual(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (_usuarioOriginal == UsuarioSoporte || _usuarioOriginal.ToLower() == "administrator" || _usuarioOriginal.ToLower() == "administrador")
                {
                    Log(progress, $"Usuario {_usuarioOriginal} mantiene privilegios", Color.Gray);
                    return true;
                }

                RunCommand("net", $"localgroup Administradores {_usuarioOriginal} /delete");
                Log(progress, $"Privilegios admin removidos de {_usuarioOriginal}", Color.Lime);
                _softwareInstalado.Add("Admin removido");
                return true;
            }
            catch
            {
                return true;
            }
        });
    }

    #endregion

    #region Configuración de Red y Sistema

    private async Task<bool> ConfigurarRedPrivada(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                RunCommand("powershell", "-Command \"Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private\"");
                RunCommand("netsh", "advfirewall firewall set rule group=\"Network Discovery\" new enable=Yes");
                RunCommand("netsh", "advfirewall firewall set rule group=\"File and Printer Sharing\" new enable=Yes");
                Log(progress, "Red configurada como privada", Color.Lime);
                _softwareInstalado.Add("Red privada");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> ConfigurarHoraAutomatica(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                RunCommand("tzutil", "/s \"Central Standard Time (Mexico)\"");
                RunCommand("sc", "config w32time start= auto");
                RunCommand("net", "start w32time");
                RunCommand("w32tm", "/config /manualpeerlist:\"time.windows.com\" /syncfromflags:manual /reliable:yes /update");
                RunCommand("w32tm", "/resync /force");
                Log(progress, "Hora automática configurada", Color.Lime);
                _softwareInstalado.Add("Hora auto");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> HabilitarWindowsUpdate(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                // Habilitar servicio Windows Update
                RunCommand("sc", "config wuauserv start= auto");
                RunCommand("net", "start wuauserv");

                // Habilitar BITS
                RunCommand("sc", "config BITS start= auto");
                RunCommand("net", "start BITS");

                // Configurar políticas de Windows Update via registro
                var wuPolicyPath = @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU";
                using var key = Registry.LocalMachine.CreateSubKey(wuPolicyPath);
                if (key != null)
                {
                    key.SetValue("NoAutoUpdate", 0, RegistryValueKind.DWord);
                    key.SetValue("AUOptions", 4, RegistryValueKind.DWord);
                    key.SetValue("ScheduledInstallDay", 0, RegistryValueKind.DWord);
                    key.SetValue("ScheduledInstallTime", 3, RegistryValueKind.DWord);
                }

                // Habilitar Feature Updates
                var wuPath = @"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";
                using var wuKey = Registry.LocalMachine.CreateSubKey(wuPath);
                // Eliminar restricciones
                try
                {
                    wuKey?.DeleteValue("TargetReleaseVersion", false);
                    wuKey?.DeleteValue("TargetReleaseVersionInfo", false);
                    wuKey?.DeleteValue("ProductVersion", false);
                }
                catch { }

                // Habilitar actualizaciones continuas
                var settingsPath = @"SOFTWARE\Microsoft\WindowsUpdate\UX\Settings";
                using var settingsKey = Registry.LocalMachine.CreateSubKey(settingsPath);
                if (settingsKey != null)
                {
                    settingsKey.SetValue("IsContinuousInnovationOptedIn", 1, RegistryValueKind.DWord);
                    settingsKey.SetValue("AllowMUUpdateService", 1, RegistryValueKind.DWord);
                }

                // Forzar detección
                RunCommand("UsoClient.exe", "StartScan", waitForExit: false);

                Log(progress, "Windows Update habilitado con Feature Updates", Color.Lime);
                _softwareInstalado.Add("Windows Update Auto");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    #endregion

    #region Instalaciones de Software

    private async Task<bool> RemoverOfficePrevio(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var ctrPath = @"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe";

                if (!File.Exists(ctrPath))
                {
                    Log(progress, "No hay versiones previas de Office", Color.Gray);
                    return true;
                }

                Log(progress, "Desinstalando Office Click-to-Run...", Color.Yellow);

                var psi = new ProcessStartInfo
                {
                    FileName = ctrPath,
                    Arguments = "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=AllProducts DisplayLevel=False",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var proc = Process.Start(psi);
                proc?.WaitForExit(300000);

                // Esperar a que termine
                for (int i = 0; i < 30; i++)
                {
                    var running = Process.GetProcessesByName("OfficeClickToRun");
                    if (running.Length == 0) break;
                    Thread.Sleep(2000);
                }

                Log(progress, "Office previo desinstalado", Color.Lime);
                _softwareInstalado.Add("Office previo removido");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> InstalarWinRAR(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (File.Exists(@"C:\Program Files\WinRAR\WinRAR.exe"))
                {
                    Log(progress, "WinRAR ya está instalado", Color.Gray);
                    _softwareInstalado.Add("WinRAR");
                    return true;
                }

                var instaladores = Directory.GetFiles(RutaWinRAR, "winrar*.exe", SearchOption.AllDirectories);
                var instalador = instaladores.FirstOrDefault(f => f.Contains("x64"));
                if (instalador == null) instalador = instaladores.FirstOrDefault();

                if (instalador != null && File.Exists(instalador))
                {
                    Log(progress, "Instalando WinRAR...", Color.Yellow);
                    var proc = Process.Start(instalador, "/S");
                    proc?.WaitForExit(120000);
                    Log(progress, "WinRAR instalado", Color.Lime);

                    // Aplicar licencia
                    var licencia = Directory.GetFiles(RutaWinRAR, "rarreg.key", SearchOption.AllDirectories).FirstOrDefault();
                    if (licencia != null)
                    {
                        File.Copy(licencia, @"C:\Program Files\WinRAR\rarreg.key", true);
                        Log(progress, "Licencia de WinRAR aplicada", Color.Lime);
                    }

                    _softwareInstalado.Add("WinRAR");
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> InstalarDotNet35(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                // Verificar si ya está instalado
                var checkResult = RunCommand("dism", "/Online /Get-FeatureInfo /FeatureName:NetFx3", timeout: 30000);
                if (checkResult.Contains("State : Enabled") || checkResult.Contains("Estado : Habilitado"))
                {
                    Log(progress, ".NET 3.5 ya está instalado", Color.Gray);
                    _softwareInstalado.Add(".NET 3.5");
                    return true;
                }

                Log(progress, "Instalando .NET 3.5 (puede tardar)...", Color.Yellow);

                // Intentar instalación offline
                if (Directory.Exists(RutaDotNet))
                {
                    RunCommand("dism", $"/Online /Enable-Feature /FeatureName:NetFx3 /All /Source:\"{RutaDotNet}\" /LimitAccess /NoRestart", timeout: 600000);
                }
                else
                {
                    // Instalación online
                    RunCommand("dism", "/Online /Enable-Feature /FeatureName:NetFx3 /All /NoRestart", timeout: 600000);
                }

                Log(progress, ".NET 3.5 instalado", Color.Lime);
                _softwareInstalado.Add(".NET 3.5");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> InstalarAcrobat(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (File.Exists(@"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"))
                {
                    Log(progress, "Acrobat Reader ya está instalado", Color.Gray);
                    _softwareInstalado.Add("Acrobat");
                    return true;
                }

                var instalador = Directory.GetFiles(RutaAcrobat, "*.exe", SearchOption.AllDirectories).FirstOrDefault();
                if (instalador != null)
                {
                    Log(progress, "Instalando Acrobat Reader...", Color.Yellow);
                    var proc = Process.Start(instalador, "/sAll /rs /msi EULA_ACCEPT=YES");
                    proc?.WaitForExit(300000);
                    Log(progress, "Acrobat Reader instalado", Color.Lime);
                    _softwareInstalado.Add("Acrobat");
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> InstalarChrome(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (File.Exists(@"C:\Program Files\Google\Chrome\Application\chrome.exe"))
                {
                    Log(progress, "Chrome ya está instalado", Color.Gray);
                    _softwareInstalado.Add("Chrome");
                    return true;
                }

                var instaladores = Directory.GetFiles(RutaChrome, "*chrome*.exe", SearchOption.AllDirectories);
                var instalador = instaladores.FirstOrDefault();

                if (instalador != null)
                {
                    Log(progress, "Instalando Chrome...", Color.Yellow);
                    var proc = Process.Start(instalador, "/silent /install");
                    proc?.WaitForExit(180000);
                    Log(progress, "Chrome instalado", Color.Lime);
                    _softwareInstalado.Add("Chrome");
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> CopiarFondoPantalla(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var imagenServidor = Path.Combine(RutaWallpaper, "wallpaper_hcg.jpg");
                var carpetaDestino = @"C:\Users\Public\Pictures";
                var fondoLocal = Path.Combine(carpetaDestino, "wallpaper_hcg.jpg");

                if (!Directory.Exists(carpetaDestino))
                    Directory.CreateDirectory(carpetaDestino);

                if (File.Exists(imagenServidor))
                {
                    File.Copy(imagenServidor, fondoLocal, true);
                    Log(progress, "Fondo de pantalla copiado", Color.Lime);

                    // Aplicar fondo
                    using var key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
                    if (key != null)
                    {
                        key.SetValue("Wallpaper", fondoLocal);
                        key.SetValue("WallpaperStyle", "10");
                    }

                    _softwareInstalado.Add("Fondo HCG");
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> InstalarOffice(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                // Verificar si ya está instalado
                if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.EXE"))
                {
                    Log(progress, "Office 2007 ya está instalado", Color.Gray);
                    _softwareInstalado.Add("Office 2007");
                    return true;
                }

                var setupPath = Path.Combine(RutaOffice, "Ofice2007", "setup.exe");
                if (!File.Exists(setupPath))
                {
                    Log(progress, "No se encontró instalador de Office", Color.Orange);
                    return false;
                }

                // Leer serial
                var serialFile = Path.Combine(RutaOffice, "Ofice2007", "SERIAL.txt");
                if (!File.Exists(serialFile))
                {
                    Log(progress, "No se encontró serial de Office", Color.Orange);
                    return false;
                }

                var serial = File.ReadAllText(serialFile).Split('\n')[0].Trim().Replace("-", "").Split(' ')[0];

                // Crear config XML
                var configXml = $@"<Configuration Product=""Enterprise"">
    <Display Level=""basic"" CompletionNotice=""no"" SuppressModal=""yes"" AcceptEula=""yes"" />
    <PIDKEY Value=""{serial}"" />
    <OptionState Id=""WORDFiles"" State=""local"" Children=""force"" />
    <OptionState Id=""EXCELFiles"" State=""local"" Children=""force"" />
    <OptionState Id=""PPTFiles"" State=""local"" Children=""force"" />
    <OptionState Id=""OUTLOOKFiles"" State=""absent"" Children=""force"" />
    <OptionState Id=""ACCESSFiles"" State=""absent"" Children=""force"" />
    <OptionState Id=""PUBFiles"" State=""absent"" Children=""force"" />
    <OptionState Id=""ONOTEFiles"" State=""absent"" Children=""force"" />
    <OptionState Id=""GROOVEFiles"" State=""absent"" Children=""force"" />
    <OptionState Id=""INFOPATHFiles"" State=""absent"" Children=""force"" />
</Configuration>";

                var configPath = Path.Combine(Path.GetTempPath(), "Office2007Config.xml");
                File.WriteAllText(configPath, configXml);

                Log(progress, "Instalando Office 2007 (Word, Excel, PowerPoint)...", Color.Yellow);
                var proc = Process.Start(setupPath, $"/config \"{configPath}\"");
                proc?.WaitForExit(600000);

                Log(progress, "Office 2007 instalado", Color.Lime);
                _softwareInstalado.Add("Office 2007");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> InstalarDedalus(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (Directory.Exists(@"C:\Dedalus\xHIS") && Directory.GetFiles(@"C:\Dedalus\xHIS").Length > 0)
                {
                    Log(progress, "Dedalus ya está instalado", Color.Gray);
                    _softwareInstalado.Add("Dedalus");
                    return true;
                }

                var netlogon = Path.Combine(RutaDedalus, "netlogon6.bat");
                if (!File.Exists(netlogon))
                {
                    Log(progress, "No se encontró netlogon6.bat", Color.Orange);
                    return false;
                }

                if (!Directory.Exists(@"C:\Dedalus"))
                    Directory.CreateDirectory(@"C:\Dedalus");

                Log(progress, "Ejecutando netlogon6.bat...", Color.Yellow);

                var psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"{netlogon}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var proc = Process.Start(psi);
                proc?.WaitForExit(600000);

                Log(progress, "Dedalus instalado", Color.Lime);
                _softwareInstalado.Add("Dedalus");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> ConfigurarSincronizadorDedalus(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                // =========================================================================
                // 1. GUARDAR CREDENCIALES DEL SERVIDOR EN WINDOWS CREDENTIAL MANAGER
                // =========================================================================
                var serverIP = "10.2.1.17";
                var credUser = "distribucion";
                var credPass = "distribucion";

                try
                {
                    RunCommand("cmdkey", $"/delete:{serverIP}");
                    RunCommand("cmdkey", $"/add:{serverIP} /user:{credUser} /pass:{credPass}");
                    Log(progress, $"Credenciales del servidor {serverIP} guardadas", Color.Lime);
                }
                catch
                {
                    Log(progress, "No se pudo guardar credenciales (no crítico)", Color.Orange);
                }

                // =========================================================================
                // 2. AGREGAR SERVIDOR A ZONA DE INTRANET (evita bloqueos de seguridad)
                // =========================================================================
                try
                {
                    var intranetZone = $@"HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\{serverIP}";
                    RunCommand("reg", $"add \"{intranetZone}\" /v file /t REG_DWORD /d 1 /f");
                    Log(progress, "Servidor agregado a zona de Intranet", Color.Lime);
                }
                catch
                {
                    Log(progress, "No se pudo configurar zona Intranet (no crítico)", Color.Orange);
                }

                // =========================================================================
                // 3. COPIAR SINCRONIZADOR AL STARTUP (sync_xhis6.bat)
                // =========================================================================
                var syncSource = @"\\10.2.1.17\distribucion\dedalus\sincronizador\sync_xhis6.bat";
                var startupFolder = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup";
                var syncDestino = Path.Combine(startupFolder, "sync_xhis6.bat");

                Log(progress, "Copiando sincronizador al inicio de Windows...", Color.Cyan);

                if (File.Exists(syncSource))
                {
                    File.Copy(syncSource, syncDestino, true);

                    // Desbloquear el archivo copiado
                    RunCommand("powershell", $"-Command \"Unblock-File -Path '{syncDestino}' -ErrorAction SilentlyContinue\"");

                    Log(progress, $"Sincronizador copiado a Startup", Color.Lime);
                    _softwareInstalado.Add("Sync Dedalus");
                }
                else
                {
                    Log(progress, $"No se encontró sincronizador en: {syncSource}", Color.Orange);
                }

                // =========================================================================
                // 4. DESBLOQUEAR ARCHIVOS DE DEDALUS
                // =========================================================================
                if (Directory.Exists(@"C:\Dedalus"))
                {
                    RunCommand("powershell", "-Command \"Get-ChildItem 'C:\\Dedalus' -Recurse | Unblock-File -ErrorAction SilentlyContinue\"");
                    Log(progress, "Archivos Dedalus desbloqueados", Color.Lime);
                }

                Log(progress, "Configuración de sincronizador completada", Color.Lime);
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true; // Continuar aunque falle
            }
        });
    }

    private async Task<bool> InstalarAntivirus(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                if (ServiceExists("ekrn"))
                {
                    Log(progress, "ESET ya está instalado", Color.Gray);
                    _softwareInstalado.Add("ESET");
                    return true;
                }

                var instalador = Directory.GetFiles(RutaAntivirus, "*.exe", SearchOption.AllDirectories).FirstOrDefault();
                if (instalador != null)
                {
                    Log(progress, "Instalando ESET PROTECT...", Color.Yellow);
                    var proc = Process.Start(instalador, "--silent --accepteula");
                    proc?.WaitForExit(600000);

                    Thread.Sleep(5000);
                    if (ServiceExists("ekrn"))
                    {
                        Log(progress, "ESET instalado correctamente", Color.Lime);
                        _softwareInstalado.Add("ESET");
                    }
                    else
                    {
                        Log(progress, "ESET instalado (verificar manualmente)", Color.Orange);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    private async Task<bool> CopiarAccesosDirectos(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var desktop = @"C:\Users\Public\Desktop";

                if (Directory.Exists(RutaAccesos))
                {
                    var accesos = Directory.GetFiles(RutaAccesos, "*.lnk", SearchOption.AllDirectories);
                    foreach (var acceso in accesos)
                    {
                        var destino = Path.Combine(desktop, Path.GetFileName(acceso));
                        File.Copy(acceso, destino, true);
                    }

                    var urls = Directory.GetFiles(RutaAccesos, "*.url", SearchOption.AllDirectories);
                    foreach (var url in urls)
                    {
                        var destino = Path.Combine(desktop, Path.GetFileName(url));
                        File.Copy(url, destino, true);
                    }

                    Log(progress, $"{accesos.Length + urls.Length} accesos directos copiados", Color.Lime);
                    _softwareInstalado.Add("Accesos");
                }

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> LimpiarIconosDuplicados(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var palabrasClave = new[] { "xfarma", "xhis", "escritorio clinico", "hpresc", "dedalus" };
                var desktopPaths = new List<string> { @"C:\Users\Public\Desktop" };

                // Agregar escritorios de usuarios
                if (Directory.Exists(@"C:\Users"))
                {
                    foreach (var userDir in Directory.GetDirectories(@"C:\Users"))
                    {
                        var userDesktop = Path.Combine(userDir, "Desktop");
                        if (Directory.Exists(userDesktop))
                            desktopPaths.Add(userDesktop);
                    }
                }

                int eliminados = 0;

                foreach (var desktopPath in desktopPaths)
                {
                    if (!Directory.Exists(desktopPath)) continue;

                    var shortcuts = Directory.GetFiles(desktopPath, "*.lnk");

                    foreach (var palabra in palabrasClave)
                    {
                        var matching = shortcuts.Where(s =>
                            Path.GetFileNameWithoutExtension(s).ToLower().Contains(palabra)).ToList();

                        if (matching.Count > 1)
                        {
                            // Mantener el de nombre más corto
                            var sorted = matching.OrderBy(s => Path.GetFileNameWithoutExtension(s).Length).ToList();
                            foreach (var dup in sorted.Skip(1))
                            {
                                try
                                {
                                    File.Delete(dup);
                                    eliminados++;
                                }
                                catch { }
                            }
                        }
                    }
                }

                if (eliminados > 0)
                    Log(progress, $"Se eliminaron {eliminados} iconos duplicados", Color.Lime);
                else
                    Log(progress, "No se encontraron iconos duplicados", Color.Gray);

                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    #endregion

    #region Reportes Automáticos

    private async Task<bool> InstalarReporteIP(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                EnsureLogsFolder();

                var scriptPath = Path.Combine(RutaLogs, "report_ip.ps1");
                var scriptContent = GetReporteIPScript();
                File.WriteAllText(scriptPath, scriptContent, Encoding.UTF8);

                // Crear VBS launcher
                var vbsPath = scriptPath.Replace(".ps1", "_launcher.vbs");
                var vbsContent = $@"Set objShell = CreateObject(""WScript.Shell"")
objShell.Run ""powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -NonInteractive -File """"{scriptPath}"""""", 0, False
Set objShell = Nothing";
                File.WriteAllText(vbsPath, vbsContent);

                // Crear tarea programada
                CreateScheduledTask("HCG_ReporteIP", vbsPath, "PT60S", "PT3H");

                Log(progress, "Reporte de IP configurado (cada 3 horas)", Color.Lime);
                _softwareInstalado.Add("Reporte IP");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> InstalarReporteSistema(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                EnsureLogsFolder();

                var scriptPath = Path.Combine(RutaLogs, "report_system.ps1");
                var scriptContent = GetReporteSistemaScript();
                File.WriteAllText(scriptPath, scriptContent, Encoding.UTF8);

                var vbsPath = scriptPath.Replace(".ps1", "_launcher.vbs");
                var vbsContent = $@"Set objShell = CreateObject(""WScript.Shell"")
objShell.Run ""powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -NonInteractive -File """"{scriptPath}"""""", 0, False
Set objShell = Nothing";
                File.WriteAllText(vbsPath, vbsContent);

                CreateScheduledTask("HCG_ReporteSistema", vbsPath, "PT120S", null);

                Log(progress, "Reporte de sistema configurado (al iniciar sesión)", Color.Lime);
                _softwareInstalado.Add("Reporte Sistema");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private async Task<bool> InstalarReporteDiagnostico(IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                EnsureLogsFolder();

                var scriptPath = Path.Combine(RutaLogs, "report_diagnostico.ps1");
                var scriptContent = GetReporteDiagnosticoScript();
                File.WriteAllText(scriptPath, scriptContent, Encoding.UTF8);

                var vbsPath = scriptPath.Replace(".ps1", "_launcher.vbs");
                var vbsContent = $@"Set objShell = CreateObject(""WScript.Shell"")
objShell.Run ""powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -NonInteractive -File """"{scriptPath}"""""", 0, False
Set objShell = Nothing";
                File.WriteAllText(vbsPath, vbsContent);

                CreateScheduledTask("HCG_ReporteDiagnostico", vbsPath, "PT180S", "PT4H");

                Log(progress, "Reporte diagnóstico configurado (cada 4 horas)", Color.Lime);
                _softwareInstalado.Add("Reporte Diagnóstico");
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return true;
            }
        });
    }

    private void CreateScheduledTask(string taskName, string vbsPath, string logonDelay, string? repeatInterval)
    {
        // Eliminar tarea existente
        RunCommand("schtasks", $"/Delete /TN \"{taskName}\" /F");

        // Crear tarea que se ejecuta al iniciar sesión
        var args = $"/Create /TN \"{taskName}\" /TR \"wscript.exe \\\"{vbsPath}\\\"\" /SC ONLOGON /DELAY {logonDelay} /RL LIMITED /F";
        RunCommand("schtasks", args);

        // Si hay intervalo de repetición, crear otra tarea
        if (!string.IsNullOrEmpty(repeatInterval))
        {
            var repeatTaskName = $"{taskName}_Repeat";
            RunCommand("schtasks", $"/Delete /TN \"{repeatTaskName}\" /F");

            var intervalMinutes = repeatInterval == "PT3H" ? 180 : repeatInterval == "PT4H" ? 240 : 60;
            var repeatArgs = $"/Create /TN \"{repeatTaskName}\" /TR \"wscript.exe \\\"{vbsPath}\\\"\" /SC MINUTE /MO {intervalMinutes} /RL LIMITED /F";
            RunCommand("schtasks", repeatArgs);
        }
    }

    private string GetReporteIPScript() => @"
# HCG - Reporte automatico de IP (cada 3 horas + al iniciar sesion)
$ErrorActionPreference = ""SilentlyContinue""
try {
    $GoogleSheetURL = ""https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec""
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)

    # Detectar adaptadores
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ""*Ethernet*"" } | Select-Object -First 1
    $WiFiAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ""*Wi-Fi*"" -or $_.Name -like ""*WiFi*"" -or $_.Name -like ""*Wireless*"" } | Select-Object -First 1

    $MACEthernet = """"
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace ""-"", """").ToUpper() }
    if (-not $MACEthernet) { exit }

    # IP Ethernet
    $IPEthernet = """"
    if ($EthAdapter -and $EthAdapter.Status -eq ""Up"") {
        $IPEthernet = (Get-NetIPAddress -InterfaceIndex $EthAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike ""169.254.*"" } | Select-Object -First 1).IPAddress
    }

    # IP WiFi
    $IPWiFi = """"
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq ""Up"") {
        $IPWiFi = (Get-NetIPAddress -InterfaceIndex $WiFiAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike ""169.254.*"" } | Select-Object -First 1).IPAddress
    }

    # SSID WiFi
    $SSIDWiFi = """"
    if ($WiFiAdapter -and $WiFiAdapter.Status -eq ""Up"") {
        $NetshOutput = netsh wlan show interfaces 2>$null
        if ($NetshOutput) {
            $SSIDLine = $NetshOutput | Select-String ""^\s+SSID\s+:"" | Select-Object -First 1
            if ($SSIDLine) { $SSIDWiFi = ($SSIDLine.ToString() -replace ""^\s+SSID\s+:\s+"", """").Trim() }
        }
    }

    if (-not $IPEthernet -and -not $IPWiFi) { exit }

    # Verificar conectividad
    $TestOK = Test-Connection -ComputerName ""8.8.8.8"" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName ""dns.google"" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri ""https://www.google.com"" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }

    $Body = @{
        Accion = ""ip""; MACEthernet = $MACEthernet; IPEthernet = $IPEthernet
        IPWiFi = $IPWiFi; SSIDWiFi = $SSIDWiFi; NombreEquipo = $env:COMPUTERNAME
        FechaReporte = (Get-Date -Format ""dd/MM/yyyy HH:mm"")
    } | ConvertTo-Json

    # Reintentos
    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType ""application/json; charset=utf-8"" -TimeoutSec 30 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) } }
    }
} catch { exit }";

    private string GetReporteSistemaScript() => @"
# HCG - Reporte de sistema y limpieza automatica (cada inicio de sesion)
$ErrorActionPreference = ""SilentlyContinue""
try {
    $GoogleSheetURL = ""https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec""
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 180)

    # --- Identificar equipo por MAC Ethernet ---
    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ""*Ethernet*"" } | Select-Object -First 1
    $MACEthernet = """"
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace ""-"", """").ToUpper() }
    if (-not $MACEthernet) { exit }

    # === 1. LIMPIEZA DE ARCHIVOS TEMPORALES (solo archivos > 1 dia) ===
    $BytesLimpiados = 0
    $TempFolders = @(""$env:TEMP"", ""C:\Windows\Temp"", ""$env:LOCALAPPDATA\Temp"")
    foreach ($folder in $TempFolders) {
        if (Test-Path $folder) {
            Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
                ForEach-Object {
                    $sz = $_.Length; $fp = $_.FullName
                    Remove-Item $fp -Force -ErrorAction SilentlyContinue
                    if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
                }
        }
    }
    # Prefetch (>7 dias)
    if (Test-Path ""C:\Windows\Prefetch"") {
        Get-ChildItem ""C:\Windows\Prefetch"" -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
            ForEach-Object {
                $sz = $_.Length; $fp = $_.FullName
                Remove-Item $fp -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
            }
    }
    # Windows Update Download (>30 dias)
    $WUFolder = ""C:\Windows\SoftwareDistribution\Download""
    if (Test-Path $WUFolder) {
        Get-ChildItem $WUFolder -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
            ForEach-Object {
                $sz = $_.Length; $fp = $_.FullName
                Remove-Item $fp -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fp)) { $BytesLimpiados += $sz }
            }
    }
    $MBLimpiados = [math]::Round($BytesLimpiados / 1MB, 1)

    # === 2. IMPRESORAS INSTALADAS ===
    $PrinterList = @()
    Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
        $Name = $_.Name; $PortName = $_.PortName; $Type = ""Local""; $IP = """"
        if ($PortName -like ""*USB*"") { $Type = ""USB"" }
        elseif ($PortName -match ""\d+\.\d+\.\d+\.\d+"") {
            $Type = ""Red""; $IP = [regex]::Match($PortName, ""\d+\.\d+\.\d+\.\d+"").Value
        } else {
            $Port = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
            if ($Port -and $Port.PrinterHostAddress) { $Type = ""Red""; $IP = $Port.PrinterHostAddress }
        }
        if ($IP) { $PrinterList += ""$Name [$Type - $IP]"" } else { $PrinterList += ""$Name [$Type]"" }
    }

    # === 3. USUARIOS DEL SISTEMA ===
    $UserList = @()
    $AdminMembers = @()
    try {
        $AdminGroupName = (Get-LocalGroup -ErrorAction SilentlyContinue | Where-Object { $_.SID.Value -eq 'S-1-5-32-544' }).Name
        if ($AdminGroupName) {
            $AdminMembers = Get-LocalGroupMember -Group $AdminGroupName -ErrorAction SilentlyContinue | ForEach-Object { ($_.Name -split '\\')[-1] }
        }
    } catch {}
    Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.Enabled } | ForEach-Object {
        $UName = $_.Name; $IsAdmin = $AdminMembers -contains $UName
        if ($IsAdmin) { $UserList += ""$UName [Admin]"" } else { $UserList += $UName }
    }

    # === 4. APLICACIONES INSTALADAS ===
    $Apps = @()
    $RegPaths = @(""HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"", ""HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"")
    foreach ($rp in $RegPaths) {
        Get-ItemProperty $rp -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne """" -and $_.DisplayName -notlike ""Update for*"" -and $_.DisplayName -notlike ""Security Update*"" } |
            ForEach-Object { $Apps += $_.DisplayName.Trim() }
    }
    $Apps = $Apps | Select-Object -Unique | Sort-Object

    # === 5. ACCESOS DIRECTOS DEL ESCRITORIO ===
    $Shortcuts = @()
    $DesktopPaths = @(""C:\Users\Public\Desktop"")
    if ($env:USERPROFILE -and (Test-Path ""$env:USERPROFILE\Desktop"")) { $DesktopPaths += ""$env:USERPROFILE\Desktop"" }
    foreach ($dp in $DesktopPaths) {
        if (Test-Path $dp) {
            Get-ChildItem $dp -Filter ""*.lnk"" -ErrorAction SilentlyContinue | ForEach-Object { $Shortcuts += $_.BaseName }
            Get-ChildItem $dp -Filter ""*.url"" -ErrorAction SilentlyContinue | ForEach-Object { $Shortcuts += $_.BaseName + "" (web)"" }
        }
    }
    $Shortcuts = $Shortcuts | Select-Object -Unique | Sort-Object

    # === 6. ESPACIO LIBRE EN DISCO ===
    $Disco = Get-CimInstance Win32_LogicalDisk -Filter ""DeviceID='C:'"" -ErrorAction SilentlyContinue
    $EspacioLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $EspacioTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # === 7. VERSION DE WINDOWS ===
    $WinVer = """"
    try {
        $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        $Build = (Get-ItemProperty ""HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"" -ErrorAction SilentlyContinue)
        $DisplayVersion = $Build.DisplayVersion
        $WinVer = ""$($OS.Caption) $DisplayVersion (Build $($Build.CurrentBuild))""
    } catch { $WinVer = ""Desconocida"" }

    # === 8. VERIFICAR INTERNET Y ENVIAR (con reintentos) ===
    $TestOK = Test-Connection -ComputerName ""8.8.8.8"" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName ""dns.google"" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri ""https://www.google.com"" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }

    # Agregar version de Windows al inicio de las apps instaladas
    $AppsConWindows = @($WinVer) + $Apps

    $Body = @{
        Accion            = ""sistema""
        MACEthernet       = $MACEthernet
        NombreEquipo      = $env:COMPUTERNAME
        Impresoras        = ($PrinterList -join "" | "")
        Usuarios          = ($UserList -join "" | "")
        AppsInstaladas    = ($AppsConWindows -join "" | "")
        AccesosEscritorio = ($Shortcuts -join "" | "")
        EspacioLibreGB    = ""$EspacioLibreGB / $EspacioTotalGB GB""
        MBLimpiados       = $MBLimpiados
        FechaReporte      = (Get-Date -Format ""dd/MM/yyyy HH:mm"")
    } | ConvertTo-Json -Depth 3

    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType ""application/json; charset=utf-8"" -TimeoutSec 60 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 15 -Maximum 45) } }
    }
} catch { exit }";

    private string GetReporteDiagnosticoScript() => @"
# HCG - Reporte de diagnostico de salud (cada 4 horas + al iniciar sesion)
$ErrorActionPreference = ""SilentlyContinue""
try {
    $GoogleSheetURL = ""https://script.google.com/macros/s/AKfycbw74FizN4Uql3ZIp4sWT9KdYO8fAargqD-urOfrTreceUJGTeaO79jZXMnL6bqfUc01/exec""
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Start-Sleep -Seconds (Get-Random -Minimum 0 -Maximum 120)

    $EthAdapter = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Name -like ""*Ethernet*"" } | Select-Object -First 1
    $MACEthernet = """"
    if ($EthAdapter) { $MACEthernet = ($EthAdapter.MacAddress -replace ""-"", """").ToUpper() }
    if (-not $MACEthernet) { exit }

    # Verificar conectividad
    $TestOK = Test-Connection -ComputerName ""8.8.8.8"" -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $TestOK) { $TestOK = Test-Connection -ComputerName ""dns.google"" -Count 1 -Quiet -ErrorAction SilentlyContinue }
    if (-not $TestOK) { try { $null = Invoke-WebRequest -Uri ""https://www.google.com"" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop; $TestOK = $true } catch {} }
    if (-not $TestOK) { exit }

    # --- RAM ---
    $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $RAMTotalGB = if ($OS.TotalVisibleMemorySize) { [math]::Round($OS.TotalVisibleMemorySize / 1MB, 1) } else { 0 }
    $RAMLibreGB = if ($OS.FreePhysicalMemory) { [math]::Round($OS.FreePhysicalMemory / 1MB, 1) } else { 0 }
    $RAMUsadaGB = [math]::Round($RAMTotalGB - $RAMLibreGB, 1)
    $RAMPct = if ($RAMTotalGB -gt 0) { [math]::Round(($RAMUsadaGB / $RAMTotalGB) * 100, 0) } else { 0 }

    # --- Top 5 procesos ---
    $Top5 = Get-Process -ErrorAction SilentlyContinue |
        Sort-Object WorkingSet64 -Descending | Select-Object -First 5 |
        ForEach-Object { ""$($_.ProcessName) ($([math]::Round($_.WorkingSet64 / 1MB, 0)) MB)"" }
    $Top5Str = $Top5 -join "" | ""

    # --- Chrome ---
    $ChromeProcs = Get-Process -Name ""chrome"" -ErrorAction SilentlyContinue
    $ChromeMB = 0; $ChromeCount = 0
    if ($ChromeProcs) {
        $ChromeCount = @($ChromeProcs).Count
        $ChromeMB = [math]::Round(($ChromeProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    # --- Dedalus ---
    $DedalusProcs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like ""*dedalus*"" }
    $DedalusMB = 0; $DedalusCount = 0
    if ($DedalusProcs) {
        $DedalusCount = @($DedalusProcs).Count
        $DedalusMB = [math]::Round(($DedalusProcs | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    }

    $TotalProcs = @(Get-Process -ErrorAction SilentlyContinue).Count
    $CPU = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $CPUPct = if ($CPU) { $CPU.LoadPercentage } else { 0 }
    if (-not $CPUPct) { $CPUPct = 0 }

    $PageFile = Get-CimInstance Win32_PageFileUsage -ErrorAction SilentlyContinue | Select-Object -First 1
    $PageFileUsadoMB = if ($PageFile) { $PageFile.CurrentUsage } else { 0 }
    $PageFileTotalMB = if ($PageFile) { $PageFile.AllocatedBaseSize } else { 0 }

    $LastBoot = $OS.LastBootUpTime
    $UptimeDias = if ($LastBoot) { [math]::Round(((Get-Date) - $LastBoot).TotalDays, 1) } else { 0 }

    $Disco = Get-CimInstance Win32_LogicalDisk -Filter ""DeviceID='C:'"" -ErrorAction SilentlyContinue
    $DiscoLibreGB = if ($Disco) { [math]::Round($Disco.FreeSpace / 1GB, 1) } else { 0 }
    $DiscoTotalGB = if ($Disco) { [math]::Round($Disco.Size / 1GB, 0) } else { 0 }

    # --- Estado y Recomendaciones ---
    $Estado = ""OK""; $Recomendaciones = @()
    if ($RAMPct -gt 85) {
        $Estado = ""Critico""; $Recomendaciones += ""RAM critica ($RAMPct%). Se recomienda ampliar memoria""
    } elseif ($RAMPct -gt 70) {
        $Estado = ""Atencion""; $Recomendaciones += ""RAM elevada ($RAMPct%). Monitorear uso. Considerar ampliacion""
    } else {
        $Recomendaciones += ""Equipo operando con recursos suficientes""
    }
    if ($ChromeMB -gt 1500) { $Recomendaciones += ""Chrome consumiendo $ChromeMB MB. Reducir pestanas"" }
    if ($UptimeDias -gt 15) { $Recomendaciones += ""Sin reinicio hace $UptimeDias dias. Reiniciar pronto"" }
    if ($DiscoLibreGB -lt 20) { $Recomendaciones += ""Disco bajo: $DiscoLibreGB GB libres. Liberar espacio"" }
    $RecomendacionStr = $Recomendaciones -join "" | ""

    $Body = @{
        Accion = ""diagnostico""; MACEthernet = $MACEthernet; NombreEquipo = $env:COMPUTERNAME
        RAMTotalGB = $RAMTotalGB; RAMUsadaGB = $RAMUsadaGB; RAMLibreGB = $RAMLibreGB; RAMPct = $RAMPct
        Top5Procesos = $Top5Str; ChromeMB = $ChromeMB; ChromeProcs = $ChromeCount
        DedalusMB = $DedalusMB; DedalusProcs = $DedalusCount; TotalProcs = $TotalProcs
        CPUPct = $CPUPct; PageFileUsado = $PageFileUsadoMB; PageFileTotal = $PageFileTotalMB
        UptimeDias = $UptimeDias; DiscoLibreGB = ""$DiscoLibreGB / $DiscoTotalGB GB""
        Estado = $Estado; Recomendacion = $RecomendacionStr
        FechaReporte = (Get-Date -Format ""dd/MM/yyyy HH:mm"")
    } | ConvertTo-Json

    for ($i = 1; $i -le 3; $i++) {
        try {
            Invoke-RestMethod -Uri $GoogleSheetURL -Method Post -Body $Body -ContentType ""application/json; charset=utf-8"" -TimeoutSec 30 | Out-Null
            break
        } catch { if ($i -lt 3) { Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 30) } }
    }
} catch { exit }";

    #endregion

    #region Renombrar Equipo

    private async Task<bool> RenombrarEquipo(string numInventario, IProgress<ConfigProgress> progress)
    {
        return await Task.Run(() =>
        {
            try
            {
                var nuevoNombre = $"PC-{numInventario}";

                RunCommand("wmic", $"computersystem where name=\"%computername%\" call rename name=\"{nuevoNombre}\"");

                Log(progress, $"Equipo renombrado a {nuevoNombre}", Color.Lime);
                return true;
            }
            catch (Exception ex)
            {
                Log(progress, $"Error: {ex.Message}", Color.Orange);
                return false;
            }
        });
    }

    #endregion

    #region Utilidades

    private void EnsureLogsFolder()
    {
        if (!Directory.Exists(RutaLogs))
            Directory.CreateDirectory(RutaLogs);
    }

    private void Log(IProgress<ConfigProgress> progress, string message, Color color)
    {
        progress.Report(new ConfigProgress
        {
            LogMessage = message,
            LogColor = color
        });
    }

    private string RunCommand(string fileName, string arguments, bool waitForExit = true, int timeout = 120000)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null) return "";

            if (waitForExit)
            {
                var output = process.StandardOutput.ReadToEnd();
                process.WaitForExit(timeout);
                return output;
            }

            return "";
        }
        catch
        {
            return "";
        }
    }

    private List<Dictionary<string, string>> RunWmiQuery(string query)
    {
        var results = new List<Dictionary<string, string>>();
        try
        {
            var output = RunCommand("wmic", query.Replace("SELECT", "").Replace("FROM", "/format:list"), timeout: 30000);
            // Parsear output básico
            var dict = new Dictionary<string, string>();
            foreach (var line in output.Split('\n'))
            {
                if (line.Contains("="))
                {
                    var parts = line.Split('=');
                    if (parts.Length >= 2)
                        dict[parts[0].Trim()] = parts[1].Trim();
                }
            }
            if (dict.Count > 0)
                results.Add(dict);
        }
        catch { }
        return results;
    }

    private bool ServiceExists(string serviceName)
    {
        try
        {
            var services = System.ServiceProcess.ServiceController.GetServices();
            return services.Any(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return false;
        }
    }

    #endregion
}
