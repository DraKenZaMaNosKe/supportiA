================================================================================
     CONFIGURADOR HCG - HOSPITAL CIVIL DE GUADALAJARA
     Coordinacion General de Informatica
================================================================================

CONTENIDO DE ESTA CARPETA:
--------------------------

  ConfigurarEquipoHCG.ps1    - Script principal de configuracion
  Config.ps1                  - Archivo de configuracion (rutas, servidores)
  GuardarCredenciales.ps1     - Guardar credenciales de red (ejecutar 1 vez)
  EJECUTAR_ConfiguradorHCG.bat - Doble clic para iniciar

  Reportes/                   - Scripts de reportes automaticos
    report_system.ps1         - Reporte de sistema (se instala en equipos)
    report_ip.ps1             - Reporte de IP
    report_diagnostico.ps1    - Diagnostico de salud
    InstalarReportes.ps1      - Instala tareas programadas

  Utilidades/                 - Herramientas adicionales
    ProcesarReportes.ps1      - Procesa JSONs y actualiza Excel
    EnviarPendientes.ps1      - Envia reportes pendientes
    Diagnostico_Red.ps1       - Diagnostico de conectividad

  Documentacion/              - Documentacion del proyecto
    LEEME.txt                 - Manual de usuario
    DIAGRAMAS_UML.md          - Diagramas del sistema (Mermaid)


INSTRUCCIONES DE USO:
---------------------

PRIMERA VEZ (en tu equipo de soporte):
  1. Ejecutar GuardarCredenciales.ps1 para guardar credenciales de red
  2. Esto solo se hace UNA vez por equipo de soporte

PARA CONFIGURAR UN EQUIPO NUEVO:
  1. Conectar el equipo a la red del hospital
  2. Copiar esta carpeta al equipo (o ejecutar desde USB)
  3. Doble clic en: EJECUTAR_ConfiguradorHCG.bat
  4. Seguir las instrucciones en pantalla
  5. Ingresar numero de inventario cuando se solicite


QUE HACE EL CONFIGURADOR:
-------------------------

  [x] Renombra el equipo (PC-XXXXX)
  [x] Crea usuario "Soporte" para acceso remoto
  [x] Instala software:
      - Office 2007
      - .NET Framework 3.5
      - Google Chrome
      - Adobe Reader DC
      - WinRAR (con licencia)
      - ESET Antivirus
  [x] Instala Dedalus / xHIS v6
  [x] Configura sincronizador al inicio de Windows
  [x] Genera wallpaper con datos del equipo
  [x] Registra equipo en Google Sheets
  [x] Instala reportes automaticos (cada 3 horas)


REQUISITOS:
-----------

  - Windows 10/11 Pro
  - Conexion a la red del hospital
  - Acceso a servidores: 10.2.1.13, 10.2.1.17
  - Ejecutar como Administrador


SOPORTE:
--------

  Coordinacion General de Informatica
  Hospital Civil de Guadalajara


================================================================================
  Version: 2.0 | Fecha: Febrero 2026
================================================================================
