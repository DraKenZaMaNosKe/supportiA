// =============================================================================
// SCRIPT DE LIMPIEZA Y FORMATO - EJECUTAR UNA VEZ
// =============================================================================
// Copia este codigo en Apps Script y ejecuta la funcion "limpiarYFormatearHoja"
// =============================================================================

function limpiarYFormatearHoja() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  ui.alert("Iniciando limpieza y formato...");

  // 1. Limpiar toda la hoja primero
  sheet.clear();

  // 2. Configurar encabezado principal (Filas 1-3)
  configurarEncabezado(sheet);

  // 3. Configurar fila de titulos de columnas (Fila 4)
  configurarTitulosColumnas(sheet);

  // 4. Configurar anchos de columna
  configurarAnchosColumnas(sheet);

  // 5. Congelar filas de encabezado
  sheet.setFrozenRows(4);

  ui.alert("Hoja limpiada y formateada.\n\nAhora necesitas volver a registrar los equipos o copiar los datos manualmente a partir de la fila 5.");
}

function configurarEncabezado(sheet) {
  // Fila 1: Titulo principal
  sheet.getRange("A1:X1").merge();
  sheet.getRange("A1").setValue("HOSPITAL CIVIL DE GUADALAJARA - FRAY ANTONIO ALCALDE");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A1").setBackground("#1565C0").setFontColor("#FFFFFF");
  sheet.setRowHeight(1, 35);

  // Fila 2: Subtitulo
  sheet.getRange("A2:X2").merge();
  sheet.getRange("A2").setValue("REGISTRO DE EQUIPOS DE COMPUTO - COORDINACION GENERAL DE INFORMATICA - EXT. 54425");
  sheet.getRange("A2").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2").setBackground("#1976D2").setFontColor("#FFFFFF");
  sheet.setRowHeight(2, 25);

  // Fila 3: Espacio
  sheet.setRowHeight(3, 10);
  sheet.getRange("A3:X3").setBackground("#FFFFFF");
}

function configurarTitulosColumnas(sheet) {
  var titulos = [
    "No.",           // A
    "Fecha",         // B
    "Inv. ST",       // C
    "Marca",         // D
    "Modelo",        // E
    "No. Serie",     // F
    "Procesador",    // G
    "Nucleos",       // H
    "RAM",           // I
    "Disco",         // J
    "Graficos",      // K
    "WiFi",          // L
    "BT",            // M
    "S.O.",          // N
    "MAC Ethernet",  // O
    "MAC WiFi",      // P
    "Product Key",   // Q
    "Fab.",          // R
    "Garantia",      // S
    "Ubicacion",     // T
    "Departamento",  // U
    "Usuario",       // V
    "Estado",        // W
    "FAA",           // X
    "IP Ethernet",   // Y
    "IP WiFi",       // Z
    "Red WiFi",      // AA
    "Ult. Conexion"  // AB
  ];

  sheet.getRange(4, 1, 1, titulos.length).setValues([titulos]);

  // Formato de encabezados
  var rangoTitulos = sheet.getRange(4, 1, 1, titulos.length);
  rangoTitulos.setFontWeight("bold");
  rangoTitulos.setHorizontalAlignment("center");
  rangoTitulos.setVerticalAlignment("middle");
  rangoTitulos.setBackground("#0D47A1");
  rangoTitulos.setFontColor("#FFFFFF");
  rangoTitulos.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(4, 30);
}

function configurarAnchosColumnas(sheet) {
  var anchos = {
    1: 40,    // A: No.
    2: 85,    // B: Fecha
    3: 55,    // C: Inv. ST
    4: 60,    // D: Marca
    5: 170,   // E: Modelo
    6: 95,    // F: No. Serie
    7: 175,   // G: Procesador
    8: 55,    // H: Nucleos
    9: 80,    // I: RAM
    10: 85,   // J: Disco
    11: 95,   // K: Graficos
    12: 55,   // L: WiFi
    13: 35,   // M: BT
    14: 75,   // N: S.O.
    15: 115,  // O: MAC Ethernet
    16: 115,  // P: MAC WiFi
    17: 130,  // Q: Product Key
    18: 80,   // R: Fab.
    19: 80,   // S: Garantia
    20: 90,   // T: Ubicacion
    21: 100,  // U: Departamento
    22: 80,   // V: Usuario
    23: 70,   // W: Estado
    24: 65,    // X: FAA
    25: 110,   // Y: IP Ethernet
    26: 110,   // Z: IP WiFi
    27: 100,   // AA: Red WiFi
    28: 130    // AB: Ult. Conexion
  };

  for (var col in anchos) {
    sheet.setColumnWidth(parseInt(col), anchos[col]);
  }
}

// =============================================================================
// FUNCION PARA AGREGAR LEYENDA AL FINAL
// =============================================================================

function agregarLeyenda(sheet) {
  var ultimaFila = sheet.getLastRow();
  var filaLeyenda = ultimaFila + 3;

  sheet.getRange(filaLeyenda, 1).setValue("LEYENDA DE ESTADOS:");
  sheet.getRange(filaLeyenda, 1).setFontWeight("bold");

  sheet.getRange(filaLeyenda + 1, 1).setValue("Activo");
  sheet.getRange(filaLeyenda + 1, 1).setBackground("#E8F5E9");
  sheet.getRange(filaLeyenda + 1, 2).setValue("Equipo operativo");

  sheet.getRange(filaLeyenda + 2, 1).setValue("En proceso");
  sheet.getRange(filaLeyenda + 2, 1).setBackground("#FFF3E0");
  sheet.getRange(filaLeyenda + 2, 2).setValue("Configuracion en progreso");

  sheet.getRange(filaLeyenda + 3, 1).setValue("Baja");
  sheet.getRange(filaLeyenda + 3, 1).setBackground("#FFEBEE");
  sheet.getRange(filaLeyenda + 3, 2).setValue("Dado de baja");
}

// =============================================================================
// FUNCION PARA FORMATEAR DATOS EXISTENTES
// =============================================================================

function formatearDatosExistentes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ultimaFila = sheet.getLastRow();

  if (ultimaFila < 5) {
    SpreadsheetApp.getUi().alert("No hay datos para formatear");
    return;
  }

  // Formatear cada fila de datos
  for (var fila = 5; fila <= ultimaFila; fila++) {
    var rango = sheet.getRange(fila, 1, 1, 28);

    // Bordes
    rango.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

    // Alineacion
    rango.setHorizontalAlignment("center");
    rango.setVerticalAlignment("middle");

    // Procesador a la izquierda
    sheet.getRange(fila, 7).setHorizontalAlignment("left");

    // Color segun estado
    var estado = sheet.getRange(fila, 23).getValue();
    if (estado == "Activo") {
      rango.setBackground("#E8F5E9");
    } else if (estado == "En proceso") {
      rango.setBackground("#FFF3E0");
    } else if (estado == "Baja") {
      rango.setBackground("#FFEBEE");
    } else {
      rango.setBackground("#FFFFFF");
    }
  }

  // Agregar leyenda
  agregarLeyenda(sheet);

  SpreadsheetApp.getUi().alert("Formato aplicado a " + (ultimaFila - 4) + " filas");
}
