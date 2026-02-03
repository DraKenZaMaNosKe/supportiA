// =============================================================================
// GOOGLE APPS SCRIPT - REGISTRO DE EQUIPOS HCG v3.0
// =============================================================================
// Copiar este codigo en: Extensiones > Apps Script
// Luego: Implementar > Nueva implementacion > Aplicacion web
// =============================================================================

// ID del archivo seriesFAA convertido a Google Sheet
var ID_SERIES_FAA = "1goG4i0Q9Lqo3xVAuV0IcLI3_dl5QXn1z1hYRquVZarw";

// Nombre de la hoja de inventario de software
var HOJA_SOFTWARE = "Inventario_Software";

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Esperar hasta 30 segundos si otro equipo esta escribiendo
    if (!lock.tryLock(30000)) {
      return jsonResponse({ status: "BUSY", mensaje: "Servidor ocupado, reintentar" });
    }

    var sheetRegistro = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    var accion = data.Accion || "crear";

    var resultado;

    if (accion == "crear") {
      resultado = crearRegistro(sheetRegistro, data);
    } else if (accion == "actualizar") {
      resultado = actualizarRegistro(sheetRegistro, data);
    } else if (accion == "formatear") {
      resultado = formatearHojaCompleta(sheetRegistro);
    } else if (accion == "limpiar") {
      resultado = limpiarYFormatear(sheetRegistro);
    } else if (accion == "software") {
      resultado = registrarSoftware(data);
    } else if (accion == "ip") {
      resultado = actualizarIP(sheetRegistro, data);
    } else if (accion == "sistema") {
      resultado = reporteSistema(data);
    } else {
      resultado = jsonResponse({ status: "ERROR", mensaje: "Accion no reconocida" });
    }

    // Pausa minima para que Google Sheets guarde los cambios
    SpreadsheetApp.flush();
    lock.releaseLock();
    return resultado;

  } catch (error) {
    try { lock.releaseLock(); } catch (e2) {}
    return jsonResponse({ status: "ERROR", mensaje: error.toString() });
  }
}

function doGet(e) {
  return jsonResponse({
    status: "OK",
    mensaje: "API Registro Equipos HCG v3.0",
    fecha: new Date().toLocaleString("es-MX")
  });
}

function crearRegistro(sheet, data) {
  // Verificar si ya existe el equipo
  var filaExistente = buscarEquipoPorInvST(sheet, data.InvST);
  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 23).setValue("En proceso");
    formatearFila(sheet, filaExistente);
    return jsonResponse({
      status: "OK",
      accion: "actualizado",
      row: filaExistente,
      mensaje: "Equipo ya existia, actualizado a En proceso"
    });
  }

  // Primero limpiar todo lo que no sea datos (leyenda, contador)
  limpiarExtras(sheet);

  // Ahora calcular la fila correcta (justo despues del ultimo equipo)
  var ultimaFilaDatos = obtenerUltimaFilaDatos(sheet);
  var newRow = ultimaFilaDatos + 1;
  var numConsecutivo = newRow - 4;

  // Buscar FAA: primero del PowerShell, si no buscar en archivo
  var resultadoFAA = data.FAA || "";
  if (!resultadoFAA || resultadoFAA == "") {
    try {
      resultadoFAA = buscarEnSeriesFAA(data.Serie);
    } catch (e) {
      resultadoFAA = "";
    }
  }

  // Insertar datos
  var rowData = [
    numConsecutivo,                           // A: No.
    data.Fecha || Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy"), // B: Fecha
    data.InvST,                               // C: Inv. ST
    data.Marca || "Lenovo",                   // D: Marca
    data.Modelo || "ThinkCentre M70s Gen 5",  // E: Modelo
    data.Serie,                               // F: No. Serie
    data.Procesador || "Intel Core i5-14500 vPro", // G: Procesador
    data.Nucleos || 14,                       // H: Nucleos
    formatearRAM(data.RAM),                   // I: RAM
    formatearDisco(data.Disco, data.DiscoTipo), // J: Disco
    data.Graficos || "Intel UHD 770",         // K: Graficos
    data.WiFi || "Wi-Fi 6",                   // L: WiFi
    data.BT || "5.1",                         // M: BT
    data.SO || "Win 11 Pro",                  // N: S.O.
    formatearMAC(data.MACEthernet),           // O: MAC Ethernet
    formatearMAC(data.MACWiFi),               // P: MAC WiFi
    data.ProductKey || "",                    // Q: Product Key
    data.FechaFab || data.Fecha,              // R: Fab.
    data.Garantia || calcularGarantia(),      // S: Garantia
    data.Ubicacion || "",                     // T: Ubicacion
    data.Departamento || "",                  // U: Departamento
    data.Usuario || "",                       // V: Usuario
    "En proceso",                             // W: Estado
    resultadoFAA                              // X: FAA
  ];

  // Insertar fila
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

  // Aplicar formato
  formatearFila(sheet, newRow);

  // Ordenar por inventario (mayor a menor)
  ordenarPorInventario(sheet);

  // Buscar la fila final despues de ordenar
  var filaFinal = buscarEquipoPorInvST(sheet, data.InvST);

  // Recolocar contador y leyenda
  colocarContadorYLeyenda(sheet);

  return jsonResponse({
    status: "OK",
    accion: "crear",
    row: filaFinal > 0 ? filaFinal : newRow,
    faa: resultadoFAA,
    serie: data.Serie,
    inventario: data.InvST
  });
}

function actualizarRegistro(sheet, data) {
  var fila = buscarEquipoPorInvST(sheet, data.InvST);

  if (fila <= 0) {
    return jsonResponse({ status: "ERROR", mensaje: "Equipo no encontrado: " + data.InvST });
  }

  // Actualizar estado a Activo
  sheet.getRange(fila, 23).setValue("Activo");

  // Actualizar otros campos si vienen
  if (data.Ubicacion) sheet.getRange(fila, 20).setValue(data.Ubicacion);
  if (data.Departamento) sheet.getRange(fila, 21).setValue(data.Departamento);
  if (data.Usuario) sheet.getRange(fila, 22).setValue(data.Usuario);

  // Aplicar formato de activo
  formatearFila(sheet, fila);

  // Recolocar contador y leyenda
  colocarContadorYLeyenda(sheet);

  return jsonResponse({
    status: "OK",
    accion: "actualizar",
    row: fila,
    estado: "Activo"
  });
}

// =============================================================================
// REPORTE AUTOMATICO DE IP
// =============================================================================

function actualizarIP(sheet, data) {
  // Asegurar que existan las columnas de IP
  asegurarColumnasIP(sheet);

  var lastRow = obtenerUltimaFilaDatos(sheet);
  var macBuscada = String(data.MACEthernet || "").toUpperCase().replace(/[^A-F0-9]/g, "");

  if (!macBuscada || macBuscada.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC Ethernet invalida" });
  }

  for (var i = 5; i <= lastRow; i++) {
    var macCelda = String(sheet.getRange(i, 15).getValue()).toUpperCase().replace(/[^A-F0-9]/g, "");
    if (macCelda === macBuscada) {
      // Actualizar columnas de IP
      sheet.getRange(i, 25).setValue(data.IPEthernet || "");      // Y: IP Ethernet
      sheet.getRange(i, 26).setValue(data.IPWiFi || "");           // Z: IP WiFi
      sheet.getRange(i, 27).setValue(data.SSIDWiFi || "");         // AA: Red WiFi
      sheet.getRange(i, 28).setValue(data.FechaReporte || "");     // AB: Ult. Conexion

      // Formatear celdas de IP
      var rangoIP = sheet.getRange(i, 25, 1, 4);
      rangoIP.setHorizontalAlignment("center");
      rangoIP.setVerticalAlignment("middle");
      rangoIP.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

      return jsonResponse({
        status: "OK",
        accion: "ip_actualizada",
        row: i,
        ip: data.IPEthernet || data.IPWiFi,
        mensaje: "IP actualizada"
      });
    }
  }

  return jsonResponse({ status: "ERROR", mensaje: "MAC no encontrada: " + macBuscada });
}

function asegurarColumnasIP(sheet) {
  var headerY = sheet.getRange(4, 25).getValue();
  if (!headerY || String(headerY).trim() === "") {
    sheet.getRange(4, 25).setValue("IP Ethernet");
    sheet.getRange(4, 26).setValue("IP WiFi");
    sheet.getRange(4, 27).setValue("Red WiFi");
    sheet.getRange(4, 28).setValue("Ult. Conexion");

    var rangoHeaders = sheet.getRange(4, 25, 1, 4);
    rangoHeaders.setFontWeight("bold");
    rangoHeaders.setHorizontalAlignment("center");
    rangoHeaders.setVerticalAlignment("middle");
    rangoHeaders.setBackground("#0D47A1");
    rangoHeaders.setFontColor("#FFFFFF");
    rangoHeaders.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

    sheet.setColumnWidth(25, 110);
    sheet.setColumnWidth(26, 110);
    sheet.setColumnWidth(27, 100);
    sheet.setColumnWidth(28, 130);
  }
}

// =============================================================================
// REPORTE DE SISTEMA (impresoras, usuarios, apps, accesos)
// =============================================================================

function reporteSistema(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reporte_Sistema");
  if (!sheet) { sheet = crearHojaReporteSistema(ss); }

  var macBuscada = String(data.MACEthernet || "").toUpperCase().replace(/[^A-F0-9]/g, "");
  if (!macBuscada || macBuscada.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC invalida" });
  }

  // Buscar InvST en la hoja principal por MAC Ethernet
  var mainSheet = ss.getSheets()[0];
  var invST = "";
  var mainData = mainSheet.getDataRange().getValues();
  for (var i = 4; i < mainData.length; i++) {
    var macMain = String(mainData[i][14]).toUpperCase().replace(/[^A-F0-9]/g, "");
    if (macMain === macBuscada) {
      invST = String(mainData[i][2]);
      break;
    }
  }
  if (!invST) invST = data.NombreEquipo || macBuscada;

  // Buscar fila existente en Reporte_Sistema por MAC
  var filaExistente = -1;
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var macCelda = String(sheet.getRange(i, 3).getValue()).toUpperCase().replace(/[^A-F0-9]/g, "");
    if (macCelda === macBuscada) { filaExistente = i; break; }
  }

  var rowData = [
    invST,                                    // A: Inv. ST
    data.NombreEquipo || "",                  // B: Nombre Equipo
    formatearMAC(data.MACEthernet),           // C: MAC Ethernet
    data.Impresoras || "Ninguna",             // D: Impresoras
    data.Usuarios || "",                      // E: Usuarios
    data.AppsInstaladas || "",                // F: Apps Instaladas
    data.AccesosEscritorio || "",             // G: Accesos Escritorio
    data.EspacioLibreGB || "",                // H: Espacio Libre
    data.MBLimpiados ? data.MBLimpiados + " MB limpiados" : "", // I: Limpieza
    data.FechaReporte || ""                   // J: Ult. Reporte
  ];

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaReporte(sheet, filaExistente);
  } else {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaReporte(sheet, newRow);
  }

  return jsonResponse({
    status: "OK",
    accion: "sistema_actualizado",
    invST: invST,
    mensaje: "Reporte de sistema actualizado"
  });
}

function crearHojaReporteSistema(ss) {
  var sheet = ss.insertSheet("Reporte_Sistema");

  var headers = [
    "Inv. ST",            // A
    "Nombre Equipo",      // B
    "MAC Ethernet",       // C
    "Impresoras",         // D
    "Usuarios",           // E
    "Apps Instaladas",    // F
    "Accesos Escritorio", // G
    "Espacio Libre",      // H
    "Limpieza",           // I
    "Ult. Reporte"        // J
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formato encabezados
  var rangoHeaders = sheet.getRange(1, 1, 1, headers.length);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setBackground("#1565C0");
  rangoHeaders.setFontColor("#FFFFFF");
  rangoHeaders.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 30);

  // Anchos de columna
  sheet.setColumnWidth(1, 70);    // Inv. ST
  sheet.setColumnWidth(2, 110);   // Nombre Equipo
  sheet.setColumnWidth(3, 130);   // MAC Ethernet
  sheet.setColumnWidth(4, 300);   // Impresoras
  sheet.setColumnWidth(5, 200);   // Usuarios
  sheet.setColumnWidth(6, 400);   // Apps Instaladas
  sheet.setColumnWidth(7, 300);   // Accesos Escritorio
  sheet.setColumnWidth(8, 100);   // Espacio Libre
  sheet.setColumnWidth(9, 120);   // Limpieza
  sheet.setColumnWidth(10, 130);  // Ult. Reporte

  // Congelar encabezado
  sheet.setFrozenRows(1);

  return sheet;
}

function formatearFilaReporte(sheet, fila) {
  var rango = sheet.getRange(fila, 1, 1, 10);
  rango.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  rango.setVerticalAlignment("middle");
  rango.setWrap(true);

  // Centrar las primeras columnas
  sheet.getRange(fila, 1, 1, 3).setHorizontalAlignment("center");
  sheet.getRange(fila, 8, 1, 3).setHorizontalAlignment("center");

  // Alinear listas a la izquierda
  sheet.getRange(fila, 4, 1, 4).setHorizontalAlignment("left").setFontSize(9);

  // Color alterno
  var colorFondo = (fila % 2 == 0) ? "#F5F5F5" : "#FFFFFF";
  rango.setBackground(colorFondo);

  // Espacio libre: verde si > 50GB, amarillo si > 20GB, rojo si < 20GB
  var espacio = sheet.getRange(fila, 8).getValue();
  var gbLibre = parseFloat(String(espacio).replace(/[^0-9.]/g, ""));
  if (gbLibre > 50) {
    sheet.getRange(fila, 8).setBackground("#C8E6C9").setFontColor("#2E7D32");
  } else if (gbLibre > 20) {
    sheet.getRange(fila, 8).setBackground("#FFF9C4").setFontColor("#F57F17");
  } else if (gbLibre > 0) {
    sheet.getRange(fila, 8).setBackground("#FFCDD2").setFontColor("#C62828");
  }
}

function buscarEquipoPorInvST(sheet, invST) {
  var data = sheet.getDataRange().getValues();
  for (var i = 4; i < data.length; i++) {
    if (String(data[i][2]) == String(invST)) {
      return i + 1;
    }
  }
  return -1;
}

// Obtiene la ultima fila que tiene datos reales de equipo (columna D = Marca)
function obtenerUltimaFilaDatos(sheet) {
  var data = sheet.getDataRange().getValues();
  var ultimaFila = 4; // Fila de encabezados
  for (var i = 4; i < data.length; i++) {
    var marca = data[i][3]; // Columna D: Marca
    if (marca && String(marca).trim() != "") {
      ultimaFila = i + 1;
    }
  }
  return ultimaFila;
}

// Limpia todo lo que NO sea encabezado o datos de equipo (contador, leyenda, basura)
function limpiarExtras(sheet) {
  var lastRow = sheet.getLastRow();
  var ultimaFilaDatos = obtenerUltimaFilaDatos(sheet);

  // Borrar todo lo que este despues de la ultima fila de datos reales
  if (lastRow > ultimaFilaDatos) {
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearContent();
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearFormat();

    // Eliminar filas vacias de abajo hacia arriba
    for (var fila = lastRow; fila > ultimaFilaDatos; fila--) {
      try {
        sheet.deleteRow(fila);
      } catch (e) {
        // Ignorar si no se puede eliminar
      }
    }
  }
}

// Coloca el contador y la leyenda en su lugar correcto
function colocarContadorYLeyenda(sheet) {
  var ultimaFilaDatos = obtenerUltimaFilaDatos(sheet);
  var numEquipos = ultimaFilaDatos - 4;

  // Limpiar filas despues de datos
  var lastRow = sheet.getLastRow();
  if (lastRow > ultimaFilaDatos) {
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearContent();
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearFormat();
  }

  // Fila vacia de separacion
  var filaContador = ultimaFilaDatos + 2;

  // Contador
  var progreso = Math.round((numEquipos / 150) * 100 * 10) / 10;
  var fecha = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy");
  var texto = "Total: " + numEquipos + " de 150 equipos | Progreso: " + progreso + "% | Actualizado: " + fecha;

  sheet.getRange(filaContador, 1).setValue(texto);
  sheet.getRange(filaContador, 1).setFontWeight("bold").setFontColor("#1565C0").setFontSize(9);

  // Leyenda
  var filaLeyenda = filaContador + 2;
  sheet.getRange(filaLeyenda, 1).setValue("LEYENDA:").setFontWeight("bold");

  sheet.getRange(filaLeyenda + 1, 1).setValue("Activo").setBackground("#E8F5E9");
  sheet.getRange(filaLeyenda + 1, 2).setValue("Equipo operativo");

  sheet.getRange(filaLeyenda + 2, 1).setValue("En proceso").setBackground("#FFF3E0");
  sheet.getRange(filaLeyenda + 2, 2).setValue("Configuracion en progreso");

  sheet.getRange(filaLeyenda + 3, 1).setValue("Baja").setBackground("#FFEBEE");
  sheet.getRange(filaLeyenda + 3, 2).setValue("Equipo dado de baja");
}

// =============================================================================
// FUNCIONES DE FORMATO
// =============================================================================

function formatearMAC(mac) {
  if (!mac) return "";
  var limpia = String(mac).replace(/[^A-Fa-f0-9]/g, "").toUpperCase();
  if (limpia.length == 12) {
    return limpia.match(/.{2}/g).join(":");
  }
  return mac;
}

function formatearRAM(ram) {
  if (!ram) return "8 GB DDR5";
  var num = String(ram).replace(/\D/g, "");
  return num ? num + " GB DDR5" : "8 GB DDR5";
}

function formatearDisco(disco, tipo) {
  if (!disco) return "512 GB SSD";
  var num = String(disco).replace(/\D/g, "");
  var tipoStr = tipo || "SSD";
  return num ? num + " GB " + tipoStr : "512 GB SSD";
}

function calcularGarantia() {
  var fecha = new Date();
  fecha.setFullYear(fecha.getFullYear() + 3);
  return Utilities.formatDate(fecha, "America/Mexico_City", "dd/MM/yyyy");
}

function formatearFila(sheet, fila) {
  var rango = sheet.getRange(fila, 1, 1, 28);

  // Bordes
  rango.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  // Alineacion centrada
  rango.setHorizontalAlignment("center");
  rango.setVerticalAlignment("middle");

  // Procesador alineado a la izquierda
  sheet.getRange(fila, 7).setHorizontalAlignment("left");

  // Color segun estado
  var estado = sheet.getRange(fila, 23).getValue();
  if (estado == "Activo") {
    rango.setBackground("#E8F5E9"); // Verde claro
  } else if (estado == "En proceso") {
    rango.setBackground("#FFF3E0"); // Naranja claro
  } else if (estado == "Baja") {
    rango.setBackground("#FFEBEE"); // Rojo claro
  } else {
    rango.setBackground("#FFFFFF"); // Blanco
  }
}

function ordenarPorInventario(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 5) return; // Solo ordenar si hay datos

  var rango = sheet.getRange(5, 1, lastRow - 4, 28);
  rango.sort({ column: 3, ascending: false }); // Columna C (Inv. ST) de mayor a menor

  // Renumerar columna A
  var numFilas = lastRow - 4;
  for (var i = 0; i < numFilas; i++) {
    sheet.getRange(5 + i, 1).setValue(i + 1);
  }
}


function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// BUSCAR EN ARCHIVO SERIES FAA
// =============================================================================

function buscarEnSeriesFAA(numeroDeSerie) {
  try {
    var ss = SpreadsheetApp.openById(ID_SERIES_FAA);
    var sheet = ss.getSheets()[0]; // Primera hoja
    var data = sheet.getDataRange().getValues();

    // Buscar en todas las columnas (principalmente columna B que tiene los CPU)
    for (var row = 1; row < data.length; row++) { // Empezar en 1 para saltar encabezado
      for (var col = 0; col < data[row].length; col++) {
        var cellValue = String(data[row][col]).trim().toUpperCase();
        var searchValue = String(numeroDeSerie).trim().toUpperCase();

        if (cellValue == searchValue) {
          // Encontrado! Obtener el numero de la columna A de esa fila
          var numSI = data[row][0]; // Columna A tiene el numero
          return "SI #" + numSI;
        }
      }
    }

    return "NO ENCONTRADO";
  } catch (error) {
    return "Error FAA: " + error.message;
  }
}

// Funcion para probar la busqueda manualmente
function testBuscarFAA() {
  var resultado = buscarEnSeriesFAA("MZ02W5K4");
  Logger.log("Resultado: " + resultado);
  SpreadsheetApp.getUi().alert("Resultado: " + resultado);
}

// =============================================================================
// FUNCIONES MANUALES - Ejecutar desde el editor
// =============================================================================

function formatearTodasLasFilas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  for (var fila = 5; fila <= lastRow; fila++) {
    formatearFila(sheet, fila);
  }

  SpreadsheetApp.getUi().alert("Todas las filas han sido formateadas");
}

function ordenarManual() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ordenarPorInventario(sheet);
  SpreadsheetApp.getUi().alert("Ordenado por inventario (mayor a menor)");
}

// =============================================================================
// INVENTARIO DE SOFTWARE
// =============================================================================

function registrarSoftware(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_SOFTWARE);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = crearHojaSoftware(ss);
  }

  // Buscar si ya existe el equipo
  var filaExistente = buscarEquipoEnSoftware(sheet, data.InvST);

  // Preparar datos
  var fecha = data.FechaConfig || Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy HH:mm");

  var rowData = [
    data.InvST,                    // A: Inv. ST
    data.NombreEquipo || "",       // B: Nombre Equipo
    data.WindowsVersion || "",     // C: Windows Version
    data.WindowsBuild || "",       // D: Build
    data.WindowsActivado || "",    // E: Activado
    data.ProductKey || "",         // F: Product Key
    data.Office || "",             // G: Office
    data.Chrome || "",             // H: Chrome
    data.Acrobat || "",            // I: Acrobat
    data.DotNet35 || "",           // J: .NET 3.5
    data.Dedalus || "",            // K: Dedalus
    data.ESET || "",               // L: ESET
    data.WinRAR || "",             // M: WinRAR
    data.OtroSoftware || "",       // N: Otro Software
    data.UsuarioWindows || "",     // O: Usuario Windows
    fecha,                         // P: Fecha Config
    data.Notas || ""               // Q: Notas
  ];

  if (filaExistente > 0) {
    // Actualizar fila existente
    sheet.getRange(filaExistente, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaSoftware(sheet, filaExistente);
    return jsonResponse({
      status: "OK",
      accion: "software_actualizado",
      row: filaExistente,
      inventario: data.InvST
    });
  } else {
    // Nueva fila
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaSoftware(sheet, newRow);

    // Ordenar por Inv. ST descendente
    if (newRow > 2) {
      var rango = sheet.getRange(2, 1, newRow - 1, 16);
      rango.sort({ column: 1, ascending: false });
    }

    return jsonResponse({
      status: "OK",
      accion: "software_creado",
      row: newRow,
      inventario: data.InvST
    });
  }
}

function crearHojaSoftware(ss) {
  var sheet = ss.insertSheet(HOJA_SOFTWARE);

  // Encabezados
  var headers = [
    "Inv. ST",           // A
    "Nombre Equipo",     // B
    "Windows",           // C
    "Build",             // D
    "Activado",          // E
    "Product Key",       // F
    "Office",            // G
    "Chrome",            // H
    "Acrobat",           // I
    ".NET 3.5",          // J
    "Dedalus",           // K
    "ESET",              // L
    "Otro Software",     // M
    "Usuario Windows",   // N
    "Fecha Config",      // O
    "Notas"              // P
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formato encabezados
  var rangoHeaders = sheet.getRange(1, 1, 1, headers.length);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setBackground("#1565C0");
  rangoHeaders.setFontColor("#FFFFFF");
  rangoHeaders.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Anchos de columna
  sheet.setColumnWidth(1, 70);   // Inv. ST
  sheet.setColumnWidth(2, 100);  // Nombre Equipo
  sheet.setColumnWidth(3, 120);  // Windows
  sheet.setColumnWidth(4, 70);   // Build
  sheet.setColumnWidth(5, 70);   // Activado
  sheet.setColumnWidth(6, 180);  // Product Key
  sheet.setColumnWidth(7, 80);   // Office
  sheet.setColumnWidth(8, 80);   // Chrome
  sheet.setColumnWidth(9, 80);   // Acrobat
  sheet.setColumnWidth(10, 70);  // .NET 3.5
  sheet.setColumnWidth(11, 70);  // Dedalus
  sheet.setColumnWidth(12, 70);  // ESET
  sheet.setColumnWidth(13, 150); // Otro Software
  sheet.setColumnWidth(14, 120); // Usuario Windows
  sheet.setColumnWidth(15, 130); // Fecha Config
  sheet.setColumnWidth(16, 150); // Notas

  // Congelar encabezado
  sheet.setFrozenRows(1);

  return sheet;
}

function buscarEquipoEnSoftware(sheet, invST) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(invST)) {
      return i + 1;
    }
  }
  return -1;
}

function formatearFilaSoftware(sheet, fila) {
  var rango = sheet.getRange(fila, 1, 1, 16);

  rango.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  rango.setHorizontalAlignment("center");
  rango.setVerticalAlignment("middle");

  // Product Key mas pequeño
  sheet.getRange(fila, 6).setFontSize(9);

  // Color alterno
  var colorFondo = (fila % 2 == 0) ? "#F5F5F5" : "#FFFFFF";
  rango.setBackground(colorFondo);

  // Activado en verde o rojo
  var activado = sheet.getRange(fila, 5).getValue();
  if (activado == "Sí" || activado == "Si" || activado == "Yes") {
    sheet.getRange(fila, 5).setBackground("#C8E6C9").setFontColor("#2E7D32");
  } else if (activado == "No") {
    sheet.getRange(fila, 5).setBackground("#FFCDD2").setFontColor("#C62828");
  }

  // ESET en verde o rojo (columna 12)
  var eset = sheet.getRange(fila, 12).getValue();
  if (eset == "Sí" || eset == "Si" || eset == "Yes") {
    sheet.getRange(fila, 12).setBackground("#C8E6C9").setFontColor("#2E7D32");
  } else if (eset == "No") {
    sheet.getRange(fila, 12).setBackground("#FFCDD2").setFontColor("#C62828");
  }
}

// =============================================================================
// LIMPIAR Y FORMATEAR HOJA COMPLETA
// =============================================================================

function limpiarYFormatear(sheet) {
  // 1. Limpiar extras (contador, leyenda, filas basura)
  limpiarExtras(sheet);

  // 2. Recalcular ultima fila de datos
  var ultimaFila = obtenerUltimaFilaDatos(sheet);
  var numEquipos = ultimaFila - 4;

  // 3. Aplicar formato a TODAS las filas de datos
  for (var fila = 5; fila <= ultimaFila; fila++) {
    formatearFila(sheet, fila);
  }

  // 4. Renumerar columna A
  for (var i = 0; i < numEquipos; i++) {
    sheet.getRange(5 + i, 1).setValue(i + 1);
  }

  // 5. Ordenar por inventario (mayor a menor)
  ordenarPorInventario(sheet);

  // 6. Colocar contador y leyenda
  colocarContadorYLeyenda(sheet);

  return jsonResponse({
    status: "OK",
    mensaje: "Hoja limpiada y formateada",
    equipos: numEquipos
  });
}

function formatearHojaCompleta(sheet) {
  return limpiarYFormatear(sheet);
}
