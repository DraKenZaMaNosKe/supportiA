# Diagramas UML - ConfiguradorHCG
## Hospital Civil de Guadalajara - Coordinación General de Informática

---

## 1. Diagrama de Casos de Uso

```mermaid
flowchart TB
    subgraph Actores
        TEC[("👨‍💻 Técnico de Soporte")]
        USER[("👤 Usuario Final")]
        SYS[("🖥️ Sistema Windows")]
        SERVER[("🗄️ Servidor Dedalus\n10.2.1.17")]
    end

    subgraph "Sistema ConfiguradorHCG"
        UC1[Configurar Equipo Nuevo]
        UC2[Instalar Software Base]
        UC3[Configurar Usuario Soporte]
        UC4[Generar Wallpaper Personalizado]
        UC5[Registrar en Google Sheets]
        UC6[Instalar Sincronizador Dedalus]
        UC7[Configurar Reportes Automáticos]
    end

    subgraph "Sistema Sincronizador Dedalus"
        UC8[Ejecutar Sincronización al Inicio]
        UC9[Sincronizar Aplicaciones xHIS]
        UC10[Actualizar Archivos Host]
        UC11[Crear Accesos Directos]
    end

    TEC --> UC1
    TEC --> UC2
    TEC --> UC3
    TEC --> UC7

    UC1 --> UC4
    UC1 --> UC5
    UC1 --> UC6

    SYS --> UC8
    UC8 --> SERVER
    SERVER --> UC9
    SERVER --> UC10
    SERVER --> UC11

    USER --> UC11
```

---

## 2. Diagrama de Flujo - Configuración de Equipo Nuevo

```mermaid
flowchart TD
    START([🚀 Inicio: Ejecutar ConfigurarEquipoHCG.ps1]) --> A1

    subgraph PREP["📋 PREPARACIÓN"]
        A1[Verificar permisos de Administrador] --> A2{¿Es Admin?}
        A2 -->|No| A3[Auto-elevar permisos]
        A3 --> A1
        A2 -->|Sí| A4[Cargar credenciales de red]
        A4 --> A5[Conectar a servidores]
    end

    A5 --> B1

    subgraph CONFIG["⚙️ CONFIGURACIÓN BÁSICA"]
        B1[Solicitar número de inventario] --> B2[Obtener datos del hardware]
        B2 --> B3[Renombrar equipo: PC-XXXXX]
        B3 --> B4[Crear usuario Soporte]
        B4 --> B5[Configurar red como privada]
        B5 --> B6[Sincronizar hora NTP]
    end

    B6 --> C1

    subgraph SOFT["📦 INSTALACIÓN DE SOFTWARE"]
        C1[Instalar .NET Framework 3.5] --> C2[Instalar Office 2007]
        C2 --> C3[Instalar Google Chrome]
        C3 --> C4[Instalar Adobe Reader]
        C4 --> C5[Instalar WinRAR]
        C5 --> C6[Instalar ESET Antivirus]
        C6 --> C7[Desinstalar Office 365]
    end

    C7 --> D1

    subgraph DEDALUS["🏥 DEDALUS / xHIS"]
        D1[Copiar carpeta Dedalus] --> D2[Instalar xHIS v6]
        D2 --> D3[Configurar sincronizador en Startup]
        D3 --> D4[Crear accesos directos]
    end

    D4 --> E1

    subgraph FINAL["✅ FINALIZACIÓN"]
        E1[Generar wallpaper personalizado] --> E2[Generar reporte JSON]
        E2 --> E3[Enviar a Google Sheets]
        E3 --> E4[Instalar tareas programadas]
        E4 --> E5[Mostrar resumen]
    end

    E5 --> END([🎉 Fin: Equipo Configurado])

    style START fill:#4CAF50,color:#fff
    style END fill:#4CAF50,color:#fff
    style PREP fill:#E3F2FD
    style CONFIG fill:#FFF3E0
    style SOFT fill:#F3E5F5
    style DEDALUS fill:#E8F5E9
    style FINAL fill:#FFFDE7
```

---

## 3. Diagrama de Secuencia - Sincronización Dedalus al Inicio

```mermaid
sequenceDiagram
    autonumber
    participant WIN as 🖥️ Windows
    participant LNK as 📁 Startup Folder
    participant NET as 📄 netlogon6.bat
    participant SYNC as 📄 sync_xhis6.bat
    participant SRV as 🗄️ Servidor 10.2.1.17
    participant CS as ⚙️ Create Synchronicity.exe

    WIN->>LNK: Inicia sesión de usuario
    LNK->>NET: Ejecuta "Dedalus Sync.lnk"

    Note over NET: netlogon6.bat (servidor)

    NET->>SRV: net use /user:distribucion
    SRV-->>NET: Conexión establecida

    NET->>NET: Copiar archivos host
    NET->>NET: Crear C:\Dedalus
    NET->>SRV: Copiar sync_xhis6.bat
    SRV-->>NET: sync_xhis6.bat copiado a C:\Dedalus

    NET->>NET: Eliminar accesos viejos de Startup
    NET->>SYNC: Ejecutar C:\Dedalus\sync_xhis6.bat

    Note over SYNC: sync_xhis6.bat (local)

    SYNC->>SRV: Conectar con credenciales
    SYNC->>SYNC: Sincronizar hora (NTP)
    SYNC->>SRV: Copiar Create Synchronicity.exe
    SRV-->>SYNC: Copiado a LogMaquinas\%COMPUTERNAME%

    loop Para cada perfil de sincronización
        SYNC->>CS: Ejecutar perfil (xHIS_sync)
        CS->>SRV: Sincronizar archivos
        SRV-->>CS: Archivos actualizados
        CS-->>SYNC: Perfil completado
    end

    Note over CS: Perfiles: xHIS, eHC, Hpresc,<br/>xFARMA, xGPC, reportes,<br/>medemp, ec, ecjdk

    SYNC->>SYNC: Crear accesos en escritorio
    SYNC-->>WIN: Sincronización completada
```

---

## 4. Diagrama de Componentes

```mermaid
flowchart TB
    subgraph LOCAL["💻 EQUIPO LOCAL"]
        subgraph Scripts["Scripts PowerShell"]
            CFG[ConfigurarEquipoHCG.ps1]
            CRED[GuardarCredenciales.ps1]
            REP[ProcesarReportes.ps1]
            INST[InstalarReportes.ps1]
        end

        subgraph Dedalus["C:\Dedalus"]
            XHIS[xHIS v6]
            EHC[EHC]
            HPRESC[hPRESC]
            EC[Escritorio Clínico]
            SYNCBAT[sync_xhis6.bat]
        end

        subgraph Reportes["C:\HCG_Reportes"]
            LOGS[Logs\]
            JSON[Pendientes\*.json]
        end

        subgraph Startup["Startup Folder"]
            DLNK[Dedalus Sync.lnk]
        end
    end

    subgraph SERVIDOR["🗄️ SERVIDOR 10.2.1.17"]
        subgraph Distribucion["\\distribucion\dedalus"]
            SINCRO[sincronizador\]
            NETLOG[netlogon6.bat]
            SYNCSRV[sync_xhis6.bat]
            CSEXE[Create Synchronicity.exe]
        end
    end

    subgraph CLOUD["☁️ GOOGLE CLOUD"]
        SHEETS[Google Sheets\nInventario HCG]
        APPS[Apps Script v4]
    end

    CFG -->|Crea| DLNK
    CFG -->|Instala| Dedalus
    CFG -->|Genera| JSON

    DLNK -->|Ejecuta| NETLOG
    NETLOG -->|Copia y ejecuta| SYNCBAT
    SYNCBAT -->|Usa| CSEXE

    REP -->|Lee| JSON
    REP -->|Actualiza| SHEETS
    APPS -->|Formatea| SHEETS

    style LOCAL fill:#E3F2FD
    style SERVIDOR fill:#FFF3E0
    style CLOUD fill:#E8F5E9
```

---

## 5. Diagrama de Estados - Ciclo de Vida del Equipo

```mermaid
stateDiagram-v2
    [*] --> Nuevo: Equipo recibido

    Nuevo --> Configurando: Ejecutar ConfigurarEquipoHCG.ps1

    state Configurando {
        [*] --> Preparacion
        Preparacion --> InstalacionSoftware
        InstalacionSoftware --> ConfiguracionDedalus
        ConfiguracionDedalus --> Finalizacion
        Finalizacion --> [*]
    }

    Configurando --> Activo: Configuración exitosa
    Configurando --> Error: Fallo en configuración

    Error --> Configurando: Reintentar

    state Activo {
        [*] --> Operativo
        Operativo --> Sincronizando: Inicio de Windows
        Sincronizando --> Operativo: Sync completado
        Operativo --> Reportando: Cada 3 horas
        Reportando --> Operativo: Reporte enviado
    }

    Activo --> Mantenimiento: Problema detectado
    Mantenimiento --> Activo: Problema resuelto

    Activo --> Baja: Equipo dado de baja
    Baja --> [*]
```

---

## 6. Diagrama de Flujo - Reportes Automáticos

```mermaid
flowchart TD
    START([⏰ Tarea Programada: Cada 3 horas]) --> A1

    subgraph REPORTE["📊 GENERACIÓN DE REPORTE"]
        A1[report_system.ps1] --> A2[Recopilar datos del sistema]
        A2 --> A3[CPU, RAM, Disco, Red]
        A3 --> A4[Software instalado]
        A4 --> A5[Usuarios y impresoras]
        A5 --> A6[Generar JSON]
    end

    A6 --> B1

    subgraph ENVIO["📤 ENVÍO"]
        B1{¿Hay conexión a internet?}
        B1 -->|Sí| B2[Enviar a Google Sheets]
        B1 -->|No| B3[Guardar en Pendientes]
        B2 --> B4{¿Éxito?}
        B4 -->|Sí| B5[Eliminar JSON local]
        B4 -->|No| B3
        B3 --> B6[Reintentar después]
    end

    B5 --> END([✅ Reporte enviado])
    B6 --> END2([🔄 Pendiente para después])

    style START fill:#2196F3,color:#fff
    style END fill:#4CAF50,color:#fff
    style END2 fill:#FF9800,color:#fff
```

---

## 7. Diagrama de Clases - Estructura de Datos

```mermaid
classDiagram
    class EquipoHCG {
        +String NumInventario
        +String NombrePC
        +String NumeroSerie
        +String DireccionMAC
        +String ProductKey
        +String DireccionIP
        +DateTime FechaConfiguracion
        +String Estado
        +configurar()
        +generarReporte()
        +sincronizar()
    }

    class ReporteJSON {
        +String Hostname
        +String SerialNumber
        +String MACAddress
        +String IPAddress
        +Object SystemInfo
        +Object DiskInfo
        +Object NetworkInfo
        +DateTime Timestamp
        +generar()
        +enviar()
    }

    class Sincronizador {
        +String RutaServidor
        +String Usuario
        +String Password
        +Array Perfiles
        +conectar()
        +sincronizarPerfil()
        +desconectar()
    }

    class GoogleSheets {
        +String SheetID
        +String WebAppURL
        +registrarEquipo()
        +actualizarEstado()
        +obtenerInventario()
    }

    EquipoHCG "1" --> "many" ReporteJSON : genera
    EquipoHCG "1" --> "1" Sincronizador : usa
    ReporteJSON "many" --> "1" GoogleSheets : envía a
    EquipoHCG "1" --> "1" GoogleSheets : registrado en
```

---

## 8. Diagrama de Despliegue

```mermaid
flowchart TB
    subgraph HOSPITAL["🏥 RED HOSPITAL CIVIL DE GUADALAJARA"]
        subgraph EQUIPOS["Equipos de Trabajo"]
            PC1[💻 PC-00001\nLenovo ThinkCentre]
            PC2[💻 PC-00002\nLenovo ThinkCentre]
            PCN[💻 PC-XXXXX\n...]
        end

        subgraph SERVIDORES["Servidores Internos"]
            SRV1[🗄️ 10.2.1.17\nDistribución Dedalus]
            SRV2[🗄️ 10.2.1.13\nSoporte FAA]
        end

        subgraph SOPORTE["Equipo de Soporte"]
            TECH[👨‍💻 Técnico\ncon ConfiguradorHCG]
            USB[💾 USB con\ninstaladores]
        end
    end

    subgraph INTERNET["☁️ INTERNET"]
        GOOGLE[📊 Google Sheets\nInventario HCG]
        NTP[🕐 time.windows.com]
    end

    TECH -->|Configura| PC1
    TECH -->|Configura| PC2
    USB -->|Instaladores| TECH

    PC1 <-->|Sincroniza| SRV1
    PC2 <-->|Sincroniza| SRV1
    PCN <-->|Sincroniza| SRV1

    PC1 -->|Reporta| GOOGLE
    PC2 -->|Reporta| GOOGLE
    PCN -->|Reporta| GOOGLE

    PC1 -->|Sync hora| NTP

    SRV2 -->|Instaladores\nSoftware| TECH

    style HOSPITAL fill:#E3F2FD
    style INTERNET fill:#E8F5E9
```

---

## Notas

- Estos diagramas están en formato **Mermaid**
- Se pueden visualizar en:
  - GitHub (automático en archivos .md)
  - GitLab (automático)
  - VS Code (extensión Mermaid)
  - [Mermaid Live Editor](https://mermaid.live/)
  - Notion, Obsidian, etc.

---

*Generado para: Hospital Civil de Guadalajara - Coordinación General de Informática*
*Sistema: ConfiguradorHCG*
*Fecha: 2026-02-11*
