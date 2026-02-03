// =============================================================================
// GOOGLE APPS SCRIPT - REGISTRO COSMICO DE EQUIPOS HCG v4.0
// =============================================================================
// TEMA: SAINT SEIYA - CABALLEROS DE INFORMATICA
// Copiar este codigo en: Extensiones > Apps Script
// Luego: Implementar > Nueva implementacion > Aplicacion web
// =============================================================================

// ID del archivo seriesFAA convertido a Google Sheet
var ID_SERIES_FAA = "1goG4i0Q9Lqo3xVAuV0IcLI3_dl5QXn1z1hYRquVZarw";

// Nombre de la hoja de inventario de software
var HOJA_SOFTWARE = "Inventario_Software";

// =============================================================================
// PALETA DE COLORES COSMICA - SAINT SEIYA
// =============================================================================

var COSMOS = {
  // Fondos principales
  GOLD_DARK:       "#8B6914",   // Dorado oscuro (armadura)
  GOLD:            "#DAA520",   // Dorado (cosmo)
  GOLD_LIGHT:      "#FFD700",   // Dorado brillante
  NIGHT_DARK:      "#0a0a1a",   // Noche cosmica profunda
  NIGHT:           "#12122e",   // Azul noche (fondo principal)
  NIGHT_LIGHT:     "#1a1a3e",   // Azul noche claro (filas alternas)
  COSMOS_BLUE:     "#1e3a5f",   // Azul cosmico
  PEGASUS_CYAN:    "#00CED1",   // Cyan Pegasus
  ATHENA_PURPLE:   "#7B2D8E",   // Purpura Athena

  // Texto
  TEXT_GOLD:       "#FFD700",   // Texto dorado
  TEXT_WHITE:      "#FFFFFF",   // Texto blanco
  TEXT_LIGHT:      "#E0E0E0",   // Texto gris claro
  TEXT_CYAN:       "#00CED1",   // Texto cyan

  // Estados
  ACTIVO_BG:       "#1a3a1a",   // Verde cosmico fondo
  ACTIVO_TEXT:     "#66FF66",   // Verde cosmico texto
  PROCESO_BG:      "#3a3a1a",   // Amarillo cosmico fondo
  PROCESO_TEXT:    "#FFD700",   // Amarillo cosmico texto
  BAJA_BG:         "#3a1a1a",   // Rojo cosmico fondo
  BAJA_TEXT:       "#FF6666",   // Rojo cosmico texto

  // Software Si/No
  SI_BG:           "#1a3a1a",
  SI_TEXT:         "#66FF66",
  NO_BG:           "#3a1a1a",
  NO_TEXT:         "#FF6666",

  // Espacio libre disco
  DISCO_OK_BG:     "#1a3a1a",
  DISCO_OK_TEXT:   "#66FF66",
  DISCO_WARN_BG:   "#3a3a1a",
  DISCO_WARN_TEXT: "#FFD700",
  DISCO_CRIT_BG:   "#3a1a1a",
  DISCO_CRIT_TEXT: "#FF6666",

  // Bordes
  BORDER_GOLD:     "#8B6914",
  BORDER_DARK:     "#333355"
};

// Simbolos cosmicos Unicode
var SYM = {
  STAR:      "\u2605",  // ★
  STAR_OPEN: "\u2606",  // ☆
  SPARK:     "\u2734",  // ✴
  CROSS:     "\u2716",  // ✖
  ARROW:     "\u2192",  // →
  SHIELD:    "\u2726",  // ✦
  DIAMOND:   "\u25C6",  // ◆
  CIRCLE:    "\u25CF",  // ●
  COSMOS:    "\u2733",  // ✳
  FLAME:     "\u2740",  // ❀
  CONSTEL:   "\u2721",  // ✡
  TRIANGLE:  "\u25B2"   // ▲
};

// =============================================================================
// MENU PERSONALIZADO - CABALLEROS DE INFORMATICA
// =============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(SYM.STAR + ' Caballeros de Informatica ' + SYM.STAR)
    .addItem(SYM.SPARK + ' Aplicar Tema Cosmico (todas las hojas)', 'aplicarTemaSaintSeiya')
    .addSeparator()
    .addItem(SYM.STAR + ' Formatear hoja Registro', 'temaRegistro')
    .addItem(SYM.STAR + ' Formatear hoja Reporte Sistema', 'temaReporteSistema')
    .addItem(SYM.STAR + ' Formatear hoja Inventario Software', 'temaInventarioSoftware')
    .addItem(SYM.STAR + ' Formatear hoja Diagnostico Salud', 'temaDiagnosticoSalud')
    .addSeparator()
    .addItem(SYM.SHIELD + ' Ordenar Registro por Inventario', 'ordenarManual')
    .addItem(SYM.SHIELD + ' Limpiar y Reformatear Registro', 'limpiarManual')
    .addSeparator()
    .addItem(SYM.COSMOS + ' Acerca de...', 'acercaDe')
    .addToUi();
}

function acercaDe() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    SYM.STAR + ' Configurador Cosmico HCG v4.0 ' + SYM.STAR,
    'Hospital Civil FAA\n' +
    'Coordinacion General de Informatica - Ext. 54425\n\n' +
    SYM.SPARK + ' Tema: Saint Seiya - Caballeros del Zodiaco\n' +
    SYM.SHIELD + ' Los Caballeros de Informatica protegen esta red\n\n' +
    '"Enciende tu cosmo!"',
    ui.ButtonSet.OK
  );
}

// =============================================================================
// TEMA SAINT SEIYA - FUNCION PRINCIPAL
// =============================================================================

function aplicarTemaSaintSeiya() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tema en hoja Registro
  var registro = ss.getSheets()[0];
  if (registro) aplicarTemaRegistro(registro);

  // Tema en hoja Reporte_Sistema
  var reporte = ss.getSheetByName("Reporte_Sistema");
  if (reporte) aplicarTemaReporteSistema(reporte);

  // Tema en hoja Inventario_Software
  var software = ss.getSheetByName(HOJA_SOFTWARE);
  if (software) aplicarTemaInventarioSoftware(software);

  // Tema en hoja Diagnostico_Salud
  var diagnostico = ss.getSheetByName("Diagnostico_Salud");
  if (diagnostico) aplicarTemaDiagnosticoSalud(diagnostico);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    SYM.STAR + ' Tema Cosmico Aplicado ' + SYM.STAR,
    'El cosmo ha sido encendido en todas las hojas.\n\n' +
    SYM.SPARK + ' Registro: tema dorado cosmico\n' +
    SYM.SPARK + ' Reporte Sistema: tema noche estrellada\n' +
    SYM.SPARK + ' Inventario Software: tema constelaciones\n' +
    SYM.SPARK + ' Diagnostico Salud: tema vitales cosmicos\n\n' +
    '"Los Caballeros de Informatica protegen esta red"',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Funciones de atajo para el menu
function temaRegistro() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  aplicarTemaRegistro(sheet);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(SYM.STAR + ' Tema cosmico aplicado a Registro');
}

function temaReporteSistema() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reporte_Sistema");
  if (sheet) {
    aplicarTemaReporteSistema(sheet);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(SYM.STAR + ' Tema cosmico aplicado a Reporte Sistema');
  } else {
    SpreadsheetApp.getUi().alert('No se encontro la hoja Reporte_Sistema');
  }
}

function temaInventarioSoftware() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_SOFTWARE);
  if (sheet) {
    aplicarTemaInventarioSoftware(sheet);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(SYM.STAR + ' Tema cosmico aplicado a Inventario Software');
  } else {
    SpreadsheetApp.getUi().alert('No se encontro la hoja Inventario_Software');
  }
}

function temaDiagnosticoSalud() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Diagnostico_Salud");
  if (sheet) {
    aplicarTemaDiagnosticoSalud(sheet);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert(SYM.STAR + ' Tema cosmico aplicado a Diagnostico Salud');
  } else {
    SpreadsheetApp.getUi().alert('No se encontro la hoja Diagnostico_Salud');
  }
}

// =============================================================================
// TEMA REGISTRO - HOJA PRINCIPAL
// =============================================================================

function aplicarTemaRegistro(sheet) {
  var lastRow = obtenerUltimaFilaDatos(sheet);
  var lastCol = 28; // Hasta la columna AB

  // --- Fila 1: Barra superior cosmica (vacia, decorativa) ---
  var fila1 = sheet.getRange(1, 1, 1, lastCol);
  fila1.setBackground(COSMOS.NIGHT_DARK);
  fila1.merge();
  fila1.setValue("");
  sheet.setRowHeight(1, 8);

  // --- Fila 2: Titulo epico ---
  var fila2 = sheet.getRange(2, 1, 1, lastCol);
  fila2.merge();
  fila2.setValue(SYM.STAR + "  " + SYM.CONSTEL + " REGISTRO COSMICO DE EQUIPOS " + SYM.CONSTEL + " - CABALLEROS DE INFORMATICA - EXT. 54425 " + SYM.STAR);
  fila2.setBackground(COSMOS.NIGHT_DARK);
  fila2.setFontColor(COSMOS.TEXT_GOLD);
  fila2.setFontWeight("bold");
  fila2.setFontSize(13);
  fila2.setHorizontalAlignment("center");
  fila2.setVerticalAlignment("middle");
  fila2.setFontFamily("Trebuchet MS");
  sheet.setRowHeight(2, 40);

  // --- Fila 3: Sub-barra decorativa ---
  var fila3 = sheet.getRange(3, 1, 1, lastCol);
  fila3.merge();
  var subBarra = "";
  for (var i = 0; i < 40; i++) {
    subBarra += (i % 4 == 0) ? SYM.STAR + " " : SYM.CIRCLE + " ";
  }
  fila3.setValue(subBarra);
  fila3.setBackground(COSMOS.GOLD_DARK);
  fila3.setFontColor(COSMOS.TEXT_GOLD);
  fila3.setFontSize(6);
  fila3.setHorizontalAlignment("center");
  sheet.setRowHeight(3, 12);

  // --- Fila 4: Encabezados cosmicos ---
  var headersCosmicos = [
    SYM.DIAMOND + " No.",
    SYM.STAR + " Fecha",
    SYM.SHIELD + " Inv. ST",
    SYM.CIRCLE + " Marca",
    SYM.CIRCLE + " Modelo",
    SYM.SPARK + " No. Serie",
    SYM.COSMOS + " Procesador",
    SYM.CIRCLE + " Nucleos",
    SYM.TRIANGLE + " RAM",
    SYM.TRIANGLE + " Disco",
    SYM.CIRCLE + " Graficos",
    SYM.COSMOS + " WiFi",
    SYM.CIRCLE + " BT",
    SYM.SHIELD + " S.O.",
    SYM.SPARK + " MAC Ethernet",
    SYM.SPARK + " MAC WiFi",
    SYM.CONSTEL + " Product Key",
    SYM.CIRCLE + " Fab.",
    SYM.CIRCLE + " Garantia",
    SYM.ARROW + " Ubicacion",
    SYM.ARROW + " Departamento",
    SYM.ARROW + " Usuario",
    SYM.STAR + " Estado",
    SYM.STAR + " FAA",
    SYM.COSMOS + " IP Ethernet",
    SYM.COSMOS + " IP WiFi",
    SYM.COSMOS + " Red WiFi",
    SYM.COSMOS + " Ult. Conexion"
  ];

  sheet.getRange(4, 1, 1, headersCosmicos.length).setValues([headersCosmicos]);
  var rangoHeaders = sheet.getRange(4, 1, 1, lastCol);
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setFontSize(10);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setVerticalAlignment("middle");
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(4, 32);

  // Congelar hasta fila 4
  sheet.setFrozenRows(4);

  // --- Filas de datos ---
  for (var fila = 5; fila <= lastRow; fila++) {
    formatearFilaCosmica(sheet, fila);
  }

  // --- Actualizar contador y leyenda cosmicos ---
  colocarContadorYLeyendaCosmica(sheet);

  // --- Color de pestana ---
  sheet.setTabColor(COSMOS.GOLD);
}

function formatearFilaCosmica(sheet, fila) {
  var rango = sheet.getRange(fila, 1, 1, 28);

  // Fondo alterno oscuro
  var esPar = (fila % 2 == 0);
  var bgColor = esPar ? COSMOS.NIGHT : COSMOS.NIGHT_LIGHT;

  // Color segun estado (columna W = 23)
  var estado = String(sheet.getRange(fila, 23).getValue());
  if (estado == "Activo") {
    bgColor = COSMOS.ACTIVO_BG;
  } else if (estado == "En proceso") {
    bgColor = COSMOS.PROCESO_BG;
  } else if (estado == "Baja") {
    bgColor = COSMOS.BAJA_BG;
  }

  rango.setBackground(bgColor);
  rango.setFontColor(COSMOS.TEXT_LIGHT);
  rango.setFontSize(10);
  rango.setFontFamily("Consolas");
  rango.setHorizontalAlignment("center");
  rango.setVerticalAlignment("middle");
  rango.setBorder(true, true, true, true, true, true, COSMOS.BORDER_DARK, SpreadsheetApp.BorderStyle.SOLID);

  // Procesador alineado a la izquierda
  sheet.getRange(fila, 7).setHorizontalAlignment("left");

  // Colorear estado con su color especifico
  var celdaEstado = sheet.getRange(fila, 23);
  if (estado == "Activo") {
    celdaEstado.setFontColor(COSMOS.ACTIVO_TEXT).setFontWeight("bold");
  } else if (estado == "En proceso") {
    celdaEstado.setFontColor(COSMOS.PROCESO_TEXT).setFontWeight("bold");
  } else if (estado == "Baja") {
    celdaEstado.setFontColor(COSMOS.BAJA_TEXT).setFontWeight("bold");
  }

  // Inventario ST en dorado
  sheet.getRange(fila, 3).setFontColor(COSMOS.TEXT_GOLD).setFontWeight("bold");

  // No. Serie en cyan
  sheet.getRange(fila, 6).setFontColor(COSMOS.TEXT_CYAN);

  // MACs en cyan
  sheet.getRange(fila, 15).setFontColor(COSMOS.PEGASUS_CYAN).setFontSize(9);
  sheet.getRange(fila, 16).setFontColor(COSMOS.PEGASUS_CYAN).setFontSize(9);

  // Product Key mas pequeno
  sheet.getRange(fila, 17).setFontSize(8).setFontColor("#999999");

  // IPs en cyan
  sheet.getRange(fila, 25).setFontColor(COSMOS.PEGASUS_CYAN);
  sheet.getRange(fila, 26).setFontColor(COSMOS.PEGASUS_CYAN);

  // FAA en dorado
  sheet.getRange(fila, 24).setFontColor(COSMOS.TEXT_GOLD);
}

function colocarContadorYLeyendaCosmica(sheet) {
  var ultimaFilaDatos = obtenerUltimaFilaDatos(sheet);
  var numEquipos = ultimaFilaDatos - 4;

  // Limpiar filas despues de datos
  var lastRow = sheet.getLastRow();
  if (lastRow > ultimaFilaDatos) {
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearContent();
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearFormat();
  }

  // --- Separador cosmico ---
  var filaSep = ultimaFilaDatos + 1;
  var rangoSep = sheet.getRange(filaSep, 1, 1, 28);
  rangoSep.merge();
  var sepLine = "";
  for (var i = 0; i < 50; i++) {
    sepLine += (i % 5 == 0) ? SYM.STAR + " " : SYM.CIRCLE + " ";
  }
  rangoSep.setValue(sepLine);
  rangoSep.setBackground(COSMOS.GOLD_DARK);
  rangoSep.setFontColor(COSMOS.TEXT_GOLD);
  rangoSep.setFontSize(6);
  rangoSep.setHorizontalAlignment("center");
  sheet.setRowHeight(filaSep, 12);

  // --- Contador con barra de progreso cosmica ---
  var filaContador = ultimaFilaDatos + 2;
  var progreso = Math.round((numEquipos / 150) * 100 * 10) / 10;
  var fecha = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy");

  // Barra visual de progreso
  var barraLlena = Math.round((numEquipos / 150) * 30);
  var barra = "";
  for (var i = 0; i < 30; i++) {
    barra += (i < barraLlena) ? "\u2593" : "\u2591";
  }

  var textoContador = SYM.SPARK + " Cosmo: " + numEquipos + " de 150 equipos  [" + barra + "] " + progreso + "%  " + SYM.STAR + " Actualizado: " + fecha;

  var rangoContador = sheet.getRange(filaContador, 1, 1, 28);
  rangoContador.merge();
  rangoContador.setValue(textoContador);
  rangoContador.setBackground(COSMOS.NIGHT_DARK);
  rangoContador.setFontColor(COSMOS.TEXT_GOLD);
  rangoContador.setFontWeight("bold");
  rangoContador.setFontSize(11);
  rangoContador.setFontFamily("Consolas");
  rangoContador.setHorizontalAlignment("center");
  sheet.setRowHeight(filaContador, 30);

  // --- Leyenda cosmica ---
  var filaLey = filaContador + 2;

  // Titulo leyenda
  var rangoTitLey = sheet.getRange(filaLey, 1, 1, 4);
  rangoTitLey.merge();
  rangoTitLey.setValue(SYM.CONSTEL + " LEYENDA COSMICA:");
  rangoTitLey.setBackground(COSMOS.NIGHT_DARK);
  rangoTitLey.setFontColor(COSMOS.TEXT_GOLD);
  rangoTitLey.setFontWeight("bold");
  rangoTitLey.setFontSize(10);

  // Activo
  sheet.getRange(filaLey + 1, 1).setValue(SYM.STAR + " Activo");
  sheet.getRange(filaLey + 1, 1).setBackground(COSMOS.ACTIVO_BG).setFontColor(COSMOS.ACTIVO_TEXT).setFontWeight("bold");
  sheet.getRange(filaLey + 1, 2, 1, 3).merge();
  sheet.getRange(filaLey + 1, 2).setValue("Equipo operativo - Cosmo encendido");
  sheet.getRange(filaLey + 1, 2).setBackground(COSMOS.NIGHT).setFontColor(COSMOS.TEXT_LIGHT);

  // En proceso
  sheet.getRange(filaLey + 2, 1).setValue(SYM.SPARK + " En proceso");
  sheet.getRange(filaLey + 2, 1).setBackground(COSMOS.PROCESO_BG).setFontColor(COSMOS.PROCESO_TEXT).setFontWeight("bold");
  sheet.getRange(filaLey + 2, 2, 1, 3).merge();
  sheet.getRange(filaLey + 2, 2).setValue("Configuracion en progreso - Elevando el cosmo");
  sheet.getRange(filaLey + 2, 2).setBackground(COSMOS.NIGHT).setFontColor(COSMOS.TEXT_LIGHT);

  // Baja
  sheet.getRange(filaLey + 3, 1).setValue(SYM.CROSS + " Baja");
  sheet.getRange(filaLey + 3, 1).setBackground(COSMOS.BAJA_BG).setFontColor(COSMOS.BAJA_TEXT).setFontWeight("bold");
  sheet.getRange(filaLey + 3, 2, 1, 3).merge();
  sheet.getRange(filaLey + 3, 2).setValue("Equipo dado de baja - Cosmo extinguido");
  sheet.getRange(filaLey + 3, 2).setBackground(COSMOS.NIGHT).setFontColor(COSMOS.TEXT_LIGHT);

  // --- Frase final epica ---
  var filaFrase = filaLey + 5;
  var rangoFrase = sheet.getRange(filaFrase, 1, 1, 28);
  rangoFrase.merge();
  rangoFrase.setValue(SYM.STAR + " Los Caballeros de Informatica protegen esta red " + SYM.STAR + "  |  " + SYM.SHIELD + " Enciende tu cosmo! " + SYM.SHIELD);
  rangoFrase.setBackground(COSMOS.NIGHT_DARK);
  rangoFrase.setFontColor(COSMOS.ATHENA_PURPLE);
  rangoFrase.setFontWeight("bold");
  rangoFrase.setFontSize(10);
  rangoFrase.setFontFamily("Trebuchet MS");
  rangoFrase.setHorizontalAlignment("center");
  sheet.setRowHeight(filaFrase, 28);
}

// =============================================================================
// TEMA REPORTE SISTEMA
// =============================================================================

function aplicarTemaReporteSistema(sheet) {
  var lastRow = sheet.getLastRow();

  // Encabezados cosmicos
  var headersReporte = [
    SYM.SHIELD + " Inv. ST",
    SYM.STAR + " Nombre Equipo",
    SYM.SPARK + " MAC Ethernet",
    SYM.COSMOS + " Impresoras",
    SYM.DIAMOND + " Usuarios",
    SYM.CONSTEL + " Apps Instaladas",
    SYM.ARROW + " Accesos Escritorio",
    SYM.TRIANGLE + " Espacio Libre",
    SYM.FLAME + " Limpieza",
    SYM.STAR + " Ult. Reporte"
  ];

  sheet.getRange(1, 1, 1, headersReporte.length).setValues([headersReporte]);
  var rangoHeaders = sheet.getRange(1, 1, 1, headersReporte.length);
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setFontSize(10);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setVerticalAlignment("middle");
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 32);

  // Congelar encabezado
  sheet.setFrozenRows(1);

  // Formatear filas de datos
  for (var fila = 2; fila <= lastRow; fila++) {
    formatearFilaReporteCosmica(sheet, fila);
  }

  // Color de pestana
  sheet.setTabColor(COSMOS.PEGASUS_CYAN);
}

function formatearFilaReporteCosmica(sheet, fila) {
  var numCols = 10;
  var rango = sheet.getRange(fila, 1, 1, numCols);

  // Fondo alterno oscuro
  var bgColor = (fila % 2 == 0) ? COSMOS.NIGHT : COSMOS.NIGHT_LIGHT;
  rango.setBackground(bgColor);
  rango.setFontColor(COSMOS.TEXT_LIGHT);
  rango.setBorder(true, true, true, true, true, true, COSMOS.BORDER_DARK, SpreadsheetApp.BorderStyle.SOLID);
  rango.setVerticalAlignment("middle");
  rango.setWrap(true);

  // Centrar primeras columnas
  sheet.getRange(fila, 1, 1, 3).setHorizontalAlignment("center").setFontFamily("Consolas");
  sheet.getRange(fila, 8, 1, 3).setHorizontalAlignment("center").setFontFamily("Consolas");

  // Listas alineadas a la izquierda
  sheet.getRange(fila, 4, 1, 4).setHorizontalAlignment("left").setFontSize(9).setFontFamily("Consolas");

  // Inv ST en dorado
  sheet.getRange(fila, 1).setFontColor(COSMOS.TEXT_GOLD).setFontWeight("bold");

  // MAC en cyan
  sheet.getRange(fila, 3).setFontColor(COSMOS.PEGASUS_CYAN);

  // Espacio libre con colores cosmicos
  var espacio = String(sheet.getRange(fila, 8).getValue());
  var gbLibre = parseFloat(espacio.replace(/[^0-9.]/g, ""));
  if (gbLibre > 50) {
    sheet.getRange(fila, 8).setBackground(COSMOS.DISCO_OK_BG).setFontColor(COSMOS.DISCO_OK_TEXT);
  } else if (gbLibre > 20) {
    sheet.getRange(fila, 8).setBackground(COSMOS.DISCO_WARN_BG).setFontColor(COSMOS.DISCO_WARN_TEXT);
  } else if (gbLibre > 0) {
    sheet.getRange(fila, 8).setBackground(COSMOS.DISCO_CRIT_BG).setFontColor(COSMOS.DISCO_CRIT_TEXT);
  }

  // Limpieza en cyan
  sheet.getRange(fila, 9).setFontColor(COSMOS.PEGASUS_CYAN);
}

// =============================================================================
// TEMA INVENTARIO SOFTWARE
// =============================================================================

function aplicarTemaInventarioSoftware(sheet) {
  var lastRow = sheet.getLastRow();

  // Encabezados cosmicos
  var headersSoftware = [
    SYM.SHIELD + " Inv. ST",
    SYM.STAR + " Nombre Equipo",
    SYM.COSMOS + " Windows",
    SYM.CIRCLE + " Build",
    SYM.STAR + " Activado",
    SYM.CONSTEL + " Product Key",
    SYM.DIAMOND + " Office",
    SYM.DIAMOND + " Chrome",
    SYM.DIAMOND + " Acrobat",
    SYM.DIAMOND + " .NET 3.5",
    SYM.SPARK + " Dedalus",
    SYM.SHIELD + " ESET",
    SYM.CIRCLE + " WinRAR",
    SYM.ARROW + " Otro Software",
    SYM.CIRCLE + " Usuario Windows",
    SYM.STAR + " Fecha Config",
    SYM.ARROW + " Notas"
  ];

  sheet.getRange(1, 1, 1, headersSoftware.length).setValues([headersSoftware]);
  var rangoHeaders = sheet.getRange(1, 1, 1, headersSoftware.length);
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setFontSize(10);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setVerticalAlignment("middle");
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 32);

  // Congelar encabezado
  sheet.setFrozenRows(1);

  // Formatear filas de datos
  for (var fila = 2; fila <= lastRow; fila++) {
    formatearFilaSoftwareCosmica(sheet, fila);
  }

  // Color de pestana
  sheet.setTabColor(COSMOS.ATHENA_PURPLE);
}

function formatearFilaSoftwareCosmica(sheet, fila) {
  var numCols = 17;
  var rango = sheet.getRange(fila, 1, 1, numCols);

  // Fondo alterno oscuro
  var bgColor = (fila % 2 == 0) ? COSMOS.NIGHT : COSMOS.NIGHT_LIGHT;
  rango.setBackground(bgColor);
  rango.setFontColor(COSMOS.TEXT_LIGHT);
  rango.setFontFamily("Consolas");
  rango.setBorder(true, true, true, true, true, true, COSMOS.BORDER_DARK, SpreadsheetApp.BorderStyle.SOLID);
  rango.setHorizontalAlignment("center");
  rango.setVerticalAlignment("middle");

  // Inv ST en dorado
  sheet.getRange(fila, 1).setFontColor(COSMOS.TEXT_GOLD).setFontWeight("bold");

  // Product Key pequeno
  sheet.getRange(fila, 6).setFontSize(8).setFontColor("#888888");

  // Columnas Si/No con colores cosmicos (Activado=5, Chrome=8, Acrobat=9, .NET=10, Dedalus=11, ESET=12, WinRAR=13)
  var colsSiNo = [5, 8, 9, 10, 11, 12, 13];
  for (var c = 0; c < colsSiNo.length; c++) {
    var col = colsSiNo[c];
    var val = String(sheet.getRange(fila, col).getValue()).trim();
    // Usar indexOf para detectar Si/No incluso si ya tiene simbolo decorativo
    if (val.indexOf("Si") >= 0 || val.indexOf("S\u00ed") >= 0 || val == "Yes") {
      sheet.getRange(fila, col).setBackground(COSMOS.SI_BG).setFontColor(COSMOS.SI_TEXT).setFontWeight("bold");
      sheet.getRange(fila, col).setValue(SYM.STAR + " Si");
    } else if (val.indexOf("No") >= 0) {
      sheet.getRange(fila, col).setBackground(COSMOS.NO_BG).setFontColor(COSMOS.NO_TEXT);
      sheet.getRange(fila, col).setValue(SYM.CROSS + " No");
    }
  }
}

// =============================================================================
// API - doPost / doGet (mantiene funcionalidad v3)
// =============================================================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return jsonResponse({ status: "BUSY", mensaje: "Servidor ocupado, reintentar" });
    }

    var sheetRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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
    } else if (accion == "diagnostico") {
      resultado = reporteDiagnostico(data);
    } else if (accion == "tema") {
      // Aplicar tema sin UI (getUi() falla en contexto web app)
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var registro = ss.getSheets()[0];
      if (registro) aplicarTemaRegistro(registro);
      var reporte = ss.getSheetByName("Reporte_Sistema");
      if (reporte) aplicarTemaReporteSistema(reporte);
      var software = ss.getSheetByName(HOJA_SOFTWARE);
      if (software) aplicarTemaInventarioSoftware(software);
      var diagnostico = ss.getSheetByName("Diagnostico_Salud");
      if (diagnostico) aplicarTemaDiagnosticoSalud(diagnostico);
      resultado = jsonResponse({ status: "OK", mensaje: "Tema cosmico aplicado" });
    } else {
      resultado = jsonResponse({ status: "ERROR", mensaje: "Accion no reconocida" });
    }

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
    mensaje: SYM.STAR + " API Registro Cosmico de Equipos HCG v4.0 " + SYM.STAR,
    fecha: new Date().toLocaleString("es-MX")
  });
}

// =============================================================================
// CREAR REGISTRO (con formato cosmico)
// =============================================================================

function crearRegistro(sheet, data) {
  var filaExistente = buscarEquipoPorInvST(sheet, data.InvST);
  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 23).setValue("En proceso");
    formatearFilaCosmica(sheet, filaExistente);
    return jsonResponse({
      status: "OK",
      accion: "actualizado",
      row: filaExistente,
      mensaje: "Equipo ya existia, actualizado a En proceso"
    });
  }

  limpiarExtras(sheet);
  var ultimaFilaDatos = obtenerUltimaFilaDatos(sheet);
  var newRow = ultimaFilaDatos + 1;
  var numConsecutivo = newRow - 4;

  var resultadoFAA = data.FAA || "";
  if (!resultadoFAA || resultadoFAA == "") {
    try { resultadoFAA = buscarEnSeriesFAA(data.Serie); } catch (e) { resultadoFAA = ""; }
  }

  var rowData = [
    numConsecutivo,
    data.Fecha || Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy"),
    data.InvST,
    data.Marca || "Lenovo",
    data.Modelo || "ThinkCentre M70s Gen 5",
    data.Serie,
    data.Procesador || "Intel Core i5-14500 vPro",
    data.Nucleos || 14,
    formatearRAM(data.RAM),
    formatearDisco(data.Disco, data.DiscoTipo),
    data.Graficos || "Intel UHD 770",
    data.WiFi || "Wi-Fi 6",
    data.BT || "5.1",
    data.SO || "Win 11 Pro",
    formatearMAC(data.MACEthernet),
    formatearMAC(data.MACWiFi),
    data.ProductKey || "",
    data.FechaFab || data.Fecha,
    data.Garantia || calcularGarantia(),
    data.Ubicacion || "",
    data.Departamento || "",
    data.Usuario || "",
    "En proceso",
    resultadoFAA
  ];

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  formatearFilaCosmica(sheet, newRow);
  ordenarPorInventario(sheet);
  var filaFinal = buscarEquipoPorInvST(sheet, data.InvST);
  colocarContadorYLeyendaCosmica(sheet);

  return jsonResponse({
    status: "OK",
    accion: "crear",
    row: filaFinal > 0 ? filaFinal : newRow,
    faa: resultadoFAA,
    serie: data.Serie,
    inventario: data.InvST
  });
}

// =============================================================================
// ACTUALIZAR REGISTRO
// =============================================================================

function actualizarRegistro(sheet, data) {
  var fila = buscarEquipoPorInvST(sheet, data.InvST);
  if (fila <= 0) {
    return jsonResponse({ status: "ERROR", mensaje: "Equipo no encontrado: " + data.InvST });
  }

  sheet.getRange(fila, 23).setValue("Activo");
  if (data.Ubicacion) sheet.getRange(fila, 20).setValue(data.Ubicacion);
  if (data.Departamento) sheet.getRange(fila, 21).setValue(data.Departamento);
  if (data.Usuario) sheet.getRange(fila, 22).setValue(data.Usuario);

  formatearFilaCosmica(sheet, fila);
  colocarContadorYLeyendaCosmica(sheet);

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
  asegurarColumnasIP(sheet);
  var lastRow = obtenerUltimaFilaDatos(sheet);
  var macBuscada = String(data.MACEthernet || "").toUpperCase().replace(/[^A-F0-9]/g, "");

  if (!macBuscada || macBuscada.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC Ethernet invalida" });
  }

  for (var i = 5; i <= lastRow; i++) {
    var macCelda = String(sheet.getRange(i, 15).getValue()).toUpperCase().replace(/[^A-F0-9]/g, "");
    if (macCelda === macBuscada) {
      sheet.getRange(i, 25).setValue(data.IPEthernet || "");
      sheet.getRange(i, 26).setValue(data.IPWiFi || "");
      sheet.getRange(i, 27).setValue(data.SSIDWiFi || "");
      sheet.getRange(i, 28).setValue(data.FechaReporte || "");

      // Formato cosmico para IPs
      var rangoIP = sheet.getRange(i, 25, 1, 4);
      rangoIP.setHorizontalAlignment("center");
      rangoIP.setVerticalAlignment("middle");
      rangoIP.setFontColor(COSMOS.PEGASUS_CYAN);
      rangoIP.setFontFamily("Consolas");
      rangoIP.setBorder(true, true, true, true, true, true, COSMOS.BORDER_DARK, SpreadsheetApp.BorderStyle.SOLID);

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
  if (!headerY || String(headerY).indexOf("IP") < 0) {
    var ipHeaders = [
      SYM.COSMOS + " IP Ethernet",
      SYM.COSMOS + " IP WiFi",
      SYM.COSMOS + " Red WiFi",
      SYM.COSMOS + " Ult. Conexion"
    ];
    sheet.getRange(4, 25, 1, 4).setValues([ipHeaders]);

    var rangoHeaders = sheet.getRange(4, 25, 1, 4);
    rangoHeaders.setFontWeight("bold");
    rangoHeaders.setHorizontalAlignment("center");
    rangoHeaders.setVerticalAlignment("middle");
    rangoHeaders.setBackground(COSMOS.GOLD_DARK);
    rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
    rangoHeaders.setFontFamily("Trebuchet MS");
    rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);

    sheet.setColumnWidth(25, 110);
    sheet.setColumnWidth(26, 110);
    sheet.setColumnWidth(27, 100);
    sheet.setColumnWidth(28, 130);
  }
}

// =============================================================================
// REPORTE DE SISTEMA
// =============================================================================

function reporteSistema(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reporte_Sistema");
  if (!sheet) { sheet = crearHojaReporteSistema(ss); }

  var macBuscada = String(data.MACEthernet || "").toUpperCase().replace(/[^A-F0-9]/g, "");
  if (!macBuscada || macBuscada.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC invalida" });
  }

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

  var filaExistente = -1;
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var macCelda = String(sheet.getRange(i, 3).getValue()).toUpperCase().replace(/[^A-F0-9]/g, "");
    if (macCelda === macBuscada) { filaExistente = i; break; }
  }

  var rowData = [
    invST,
    data.NombreEquipo || "",
    formatearMAC(data.MACEthernet),
    data.Impresoras || "Ninguna",
    data.Usuarios || "",
    data.AppsInstaladas || "",
    data.AccesosEscritorio || "",
    data.EspacioLibreGB || "",
    data.MBLimpiados ? data.MBLimpiados + " MB limpiados" : "",
    data.FechaReporte || ""
  ];

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaReporteCosmica(sheet, filaExistente);
  } else {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaReporteCosmica(sheet, newRow);
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
    SYM.SHIELD + " Inv. ST",
    SYM.STAR + " Nombre Equipo",
    SYM.SPARK + " MAC Ethernet",
    SYM.COSMOS + " Impresoras",
    SYM.DIAMOND + " Usuarios",
    SYM.CONSTEL + " Apps Instaladas",
    SYM.ARROW + " Accesos Escritorio",
    SYM.TRIANGLE + " Espacio Libre",
    SYM.FLAME + " Limpieza",
    SYM.STAR + " Ult. Reporte"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var rangoHeaders = sheet.getRange(1, 1, 1, headers.length);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 32);

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 400);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 130);

  sheet.setFrozenRows(1);
  sheet.setTabColor(COSMOS.PEGASUS_CYAN);

  return sheet;
}

// =============================================================================
// INVENTARIO DE SOFTWARE
// =============================================================================

function registrarSoftware(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_SOFTWARE);
  if (!sheet) { sheet = crearHojaSoftware(ss); }

  var filaExistente = buscarEquipoEnSoftware(sheet, data.InvST);
  var fecha = data.FechaConfig || Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy HH:mm");

  var rowData = [
    data.InvST,
    data.NombreEquipo || "",
    data.WindowsVersion || "",
    data.WindowsBuild || "",
    data.WindowsActivado || "",
    data.ProductKey || "",
    data.Office || "",
    data.Chrome || "",
    data.Acrobat || "",
    data.DotNet35 || "",
    data.Dedalus || "",
    data.ESET || "",
    data.WinRAR || "",
    data.OtroSoftware || "",
    data.UsuarioWindows || "",
    fecha,
    data.Notas || ""
  ];

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaSoftwareCosmica(sheet, filaExistente);
    return jsonResponse({
      status: "OK",
      accion: "software_actualizado",
      row: filaExistente,
      inventario: data.InvST
    });
  } else {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaSoftwareCosmica(sheet, newRow);

    if (newRow > 2) {
      var rango = sheet.getRange(2, 1, newRow - 1, 17);
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

  var headers = [
    SYM.SHIELD + " Inv. ST",
    SYM.STAR + " Nombre Equipo",
    SYM.COSMOS + " Windows",
    SYM.CIRCLE + " Build",
    SYM.STAR + " Activado",
    SYM.CONSTEL + " Product Key",
    SYM.DIAMOND + " Office",
    SYM.DIAMOND + " Chrome",
    SYM.DIAMOND + " Acrobat",
    SYM.DIAMOND + " .NET 3.5",
    SYM.SPARK + " Dedalus",
    SYM.SHIELD + " ESET",
    SYM.CIRCLE + " WinRAR",
    SYM.ARROW + " Otro Software",
    SYM.CIRCLE + " Usuario Windows",
    SYM.STAR + " Fecha Config",
    SYM.ARROW + " Notas"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var rangoHeaders = sheet.getRange(1, 1, 1, headers.length);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 80);
  sheet.setColumnWidth(10, 70);
  sheet.setColumnWidth(11, 70);
  sheet.setColumnWidth(12, 70);
  sheet.setColumnWidth(13, 70);
  sheet.setColumnWidth(14, 150);
  sheet.setColumnWidth(15, 120);
  sheet.setColumnWidth(16, 130);
  sheet.setColumnWidth(17, 150);

  sheet.setFrozenRows(1);
  sheet.setTabColor(COSMOS.ATHENA_PURPLE);

  return sheet;
}

// =============================================================================
// DIAGNOSTICO DE SALUD
// =============================================================================

function reporteDiagnostico(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Diagnostico_Salud");
  if (!sheet) { sheet = crearHojaDiagnosticoSalud(ss); }

  var macBuscada = String(data.MACEthernet || "").toUpperCase().replace(/[^A-F0-9]/g, "");
  if (!macBuscada || macBuscada.length < 12) {
    return jsonResponse({ status: "ERROR", mensaje: "MAC invalida" });
  }

  // Buscar Inv. ST en hoja principal
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

  // Buscar fila existente por MAC
  var filaExistente = -1;
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    var macCelda = String(sheet.getRange(i, 3).getValue()).toUpperCase().replace(/[^A-F0-9]/g, "");
    if (macCelda === macBuscada) { filaExistente = i; break; }
  }

  var rowData = [
    invST,
    data.NombreEquipo || "",
    formatearMAC(data.MACEthernet),
    data.RAMTotalGB || 0,
    data.RAMUsadaGB || 0,
    data.RAMLibreGB || 0,
    data.RAMPct || 0,
    data.Top5Procesos || "",
    data.ChromeMB || 0,
    data.ChromeProcs || 0,
    data.DedalusMB || 0,
    data.DedalusProcs || 0,
    data.TotalProcs || 0,
    data.CPUPct || 0,
    data.PageFileUsado || 0,
    data.PageFileTotal || 0,
    data.UptimeDias || 0,
    data.DiscoLibreGB || "",
    data.Estado || "OK",
    data.Recomendacion || "",
    data.FechaReporte || ""
  ];

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaDiagnosticoCosmica(sheet, filaExistente);
  } else {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    formatearFilaDiagnosticoCosmica(sheet, newRow);
  }

  return jsonResponse({
    status: "OK",
    accion: "diagnostico_actualizado",
    invST: invST,
    estado: data.Estado || "OK",
    mensaje: "Diagnostico de salud actualizado"
  });
}

function crearHojaDiagnosticoSalud(ss) {
  var sheet = ss.insertSheet("Diagnostico_Salud");

  var headers = [
    SYM.SHIELD + " Inv. ST",
    SYM.STAR + " Nombre",
    SYM.SPARK + " MAC",
    SYM.TRIANGLE + " RAM Total",
    SYM.TRIANGLE + " RAM Usada",
    SYM.TRIANGLE + " RAM Libre",
    SYM.FLAME + " RAM%",
    SYM.COSMOS + " Top5 Procesos",
    SYM.DIAMOND + " Chrome MB",
    SYM.DIAMOND + " Chrome Procs",
    SYM.DIAMOND + " Dedalus MB",
    SYM.DIAMOND + " Dedalus Procs",
    SYM.CIRCLE + " Total Procs",
    SYM.FLAME + " CPU%",
    SYM.ARROW + " PageFile Usado",
    SYM.ARROW + " PageFile Total",
    SYM.STAR_OPEN + " Uptime Dias",
    SYM.TRIANGLE + " Disco Libre",
    SYM.STAR + " Estado",
    SYM.CONSTEL + " Recomendacion",
    SYM.COSMOS + " Ult. Reporte"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var rangoHeaders = sheet.getRange(1, 1, 1, headers.length);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setVerticalAlignment("middle");
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setFontSize(10);
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 32);

  // Anchos de columna
  sheet.setColumnWidth(1, 70);    // Inv ST
  sheet.setColumnWidth(2, 110);   // Nombre
  sheet.setColumnWidth(3, 130);   // MAC
  sheet.setColumnWidth(4, 80);    // RAM Total
  sheet.setColumnWidth(5, 80);    // RAM Usada
  sheet.setColumnWidth(6, 80);    // RAM Libre
  sheet.setColumnWidth(7, 60);    // RAM%
  sheet.setColumnWidth(8, 350);   // Top5 Procesos
  sheet.setColumnWidth(9, 80);    // Chrome MB
  sheet.setColumnWidth(10, 90);   // Chrome Procs
  sheet.setColumnWidth(11, 80);   // Dedalus MB
  sheet.setColumnWidth(12, 95);   // Dedalus Procs
  sheet.setColumnWidth(13, 80);   // Total Procs
  sheet.setColumnWidth(14, 60);   // CPU%
  sheet.setColumnWidth(15, 100);  // PageFile Usado
  sheet.setColumnWidth(16, 100);  // PageFile Total
  sheet.setColumnWidth(17, 85);   // Uptime Dias
  sheet.setColumnWidth(18, 100);  // Disco Libre
  sheet.setColumnWidth(19, 80);   // Estado
  sheet.setColumnWidth(20, 350);  // Recomendacion
  sheet.setColumnWidth(21, 130);  // Ult Reporte

  sheet.setFrozenRows(1);
  sheet.setTabColor("#FF4500"); // Naranja cosmico

  return sheet;
}

function aplicarTemaDiagnosticoSalud(sheet) {
  var lastRow = sheet.getLastRow();

  // Encabezados cosmicos
  var headersDiag = [
    SYM.SHIELD + " Inv. ST",
    SYM.STAR + " Nombre",
    SYM.SPARK + " MAC",
    SYM.TRIANGLE + " RAM Total",
    SYM.TRIANGLE + " RAM Usada",
    SYM.TRIANGLE + " RAM Libre",
    SYM.FLAME + " RAM%",
    SYM.COSMOS + " Top5 Procesos",
    SYM.DIAMOND + " Chrome MB",
    SYM.DIAMOND + " Chrome Procs",
    SYM.DIAMOND + " Dedalus MB",
    SYM.DIAMOND + " Dedalus Procs",
    SYM.CIRCLE + " Total Procs",
    SYM.FLAME + " CPU%",
    SYM.ARROW + " PageFile Usado",
    SYM.ARROW + " PageFile Total",
    SYM.STAR_OPEN + " Uptime Dias",
    SYM.TRIANGLE + " Disco Libre",
    SYM.STAR + " Estado",
    SYM.CONSTEL + " Recomendacion",
    SYM.COSMOS + " Ult. Reporte"
  ];

  sheet.getRange(1, 1, 1, headersDiag.length).setValues([headersDiag]);
  var rangoHeaders = sheet.getRange(1, 1, 1, headersDiag.length);
  rangoHeaders.setBackground(COSMOS.GOLD_DARK);
  rangoHeaders.setFontColor(COSMOS.TEXT_GOLD);
  rangoHeaders.setFontWeight("bold");
  rangoHeaders.setFontSize(10);
  rangoHeaders.setFontFamily("Trebuchet MS");
  rangoHeaders.setHorizontalAlignment("center");
  rangoHeaders.setVerticalAlignment("middle");
  rangoHeaders.setBorder(true, true, true, true, true, true, COSMOS.GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 32);

  // Congelar encabezado
  sheet.setFrozenRows(1);

  // Formatear filas de datos
  for (var fila = 2; fila <= lastRow; fila++) {
    formatearFilaDiagnosticoCosmica(sheet, fila);
  }

  // Color de pestana
  sheet.setTabColor("#FF4500");
}

function formatearFilaDiagnosticoCosmica(sheet, fila) {
  var numCols = 21;
  var rango = sheet.getRange(fila, 1, 1, numCols);

  // Fondo alterno oscuro
  var bgColor = (fila % 2 == 0) ? COSMOS.NIGHT : COSMOS.NIGHT_LIGHT;
  rango.setBackground(bgColor);
  rango.setFontColor(COSMOS.TEXT_LIGHT);
  rango.setFontFamily("Consolas");
  rango.setBorder(true, true, true, true, true, true, COSMOS.BORDER_DARK, SpreadsheetApp.BorderStyle.SOLID);
  rango.setHorizontalAlignment("center");
  rango.setVerticalAlignment("middle");

  // Inv ST en dorado
  sheet.getRange(fila, 1).setFontColor(COSMOS.TEXT_GOLD).setFontWeight("bold");

  // MAC en cyan
  sheet.getRange(fila, 3).setFontColor(COSMOS.PEGASUS_CYAN).setFontSize(9);

  // Top5 procesos alineado a la izquierda
  sheet.getRange(fila, 8).setHorizontalAlignment("left").setFontSize(9);

  // Recomendacion alineada a la izquierda
  sheet.getRange(fila, 20).setHorizontalAlignment("left").setFontSize(9);

  // --- Color por estado (columna 19) ---
  var estado = String(sheet.getRange(fila, 19).getValue()).trim();
  var celdaEstado = sheet.getRange(fila, 19);
  if (estado == "OK") {
    celdaEstado.setBackground(COSMOS.ACTIVO_BG).setFontColor(COSMOS.ACTIVO_TEXT).setFontWeight("bold");
  } else if (estado == "Atencion") {
    celdaEstado.setBackground(COSMOS.PROCESO_BG).setFontColor(COSMOS.PROCESO_TEXT).setFontWeight("bold");
  } else if (estado == "Critico") {
    celdaEstado.setBackground(COSMOS.BAJA_BG).setFontColor(COSMOS.BAJA_TEXT).setFontWeight("bold");
  }

  // --- RAM% con color-coding (columna 7) ---
  var ramPct = parseFloat(sheet.getRange(fila, 7).getValue()) || 0;
  var celdaRAM = sheet.getRange(fila, 7);
  if (ramPct > 85) {
    celdaRAM.setBackground(COSMOS.BAJA_BG).setFontColor(COSMOS.BAJA_TEXT).setFontWeight("bold");
  } else if (ramPct > 70) {
    celdaRAM.setBackground(COSMOS.PROCESO_BG).setFontColor(COSMOS.PROCESO_TEXT).setFontWeight("bold");
  } else {
    celdaRAM.setBackground(COSMOS.ACTIVO_BG).setFontColor(COSMOS.ACTIVO_TEXT).setFontWeight("bold");
  }

  // --- CPU% con color-coding (columna 14) ---
  var cpuPct = parseFloat(sheet.getRange(fila, 14).getValue()) || 0;
  var celdaCPU = sheet.getRange(fila, 14);
  if (cpuPct > 85) {
    celdaCPU.setBackground(COSMOS.BAJA_BG).setFontColor(COSMOS.BAJA_TEXT).setFontWeight("bold");
  } else if (cpuPct > 70) {
    celdaCPU.setBackground(COSMOS.PROCESO_BG).setFontColor(COSMOS.PROCESO_TEXT).setFontWeight("bold");
  } else {
    celdaCPU.setBackground(COSMOS.ACTIVO_BG).setFontColor(COSMOS.ACTIVO_TEXT).setFontWeight("bold");
  }

  // --- Chrome MB con alerta si > 1500 (columna 9) ---
  var chromeMB = parseFloat(sheet.getRange(fila, 9).getValue()) || 0;
  if (chromeMB > 1500) {
    sheet.getRange(fila, 9).setBackground(COSMOS.BAJA_BG).setFontColor(COSMOS.BAJA_TEXT);
  } else if (chromeMB > 800) {
    sheet.getRange(fila, 9).setBackground(COSMOS.PROCESO_BG).setFontColor(COSMOS.PROCESO_TEXT);
  }

  // --- Uptime con alerta si > 15 dias (columna 17) ---
  var uptime = parseFloat(sheet.getRange(fila, 17).getValue()) || 0;
  if (uptime > 15) {
    sheet.getRange(fila, 17).setBackground(COSMOS.BAJA_BG).setFontColor(COSMOS.BAJA_TEXT);
  } else if (uptime > 7) {
    sheet.getRange(fila, 17).setBackground(COSMOS.PROCESO_BG).setFontColor(COSMOS.PROCESO_TEXT);
  }

  // --- Disco libre con colores (columna 18) ---
  var discoStr = String(sheet.getRange(fila, 18).getValue());
  var discoLibre = parseFloat(discoStr.replace(/[^0-9.]/g, "")) || 0;
  if (discoLibre < 20) {
    sheet.getRange(fila, 18).setBackground(COSMOS.DISCO_CRIT_BG).setFontColor(COSMOS.DISCO_CRIT_TEXT);
  } else if (discoLibre < 50) {
    sheet.getRange(fila, 18).setBackground(COSMOS.DISCO_WARN_BG).setFontColor(COSMOS.DISCO_WARN_TEXT);
  } else {
    sheet.getRange(fila, 18).setBackground(COSMOS.DISCO_OK_BG).setFontColor(COSMOS.DISCO_OK_TEXT);
  }
}

// =============================================================================
// FUNCIONES AUXILIARES
// =============================================================================

function buscarEquipoPorInvST(sheet, invST) {
  var data = sheet.getDataRange().getValues();
  for (var i = 4; i < data.length; i++) {
    if (String(data[i][2]) == String(invST)) {
      return i + 1;
    }
  }
  return -1;
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

function obtenerUltimaFilaDatos(sheet) {
  var data = sheet.getDataRange().getValues();
  var ultimaFila = 4;
  for (var i = 4; i < data.length; i++) {
    var invST = data[i][2]; // Columna C: Inv. ST
    if (invST && String(invST).trim() != "") {
      ultimaFila = i + 1;
    }
  }
  return ultimaFila;
}

function limpiarExtras(sheet) {
  var lastRow = sheet.getLastRow();
  var ultimaFilaDatos = obtenerUltimaFilaDatos(sheet);
  if (lastRow > ultimaFilaDatos) {
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearContent();
    sheet.getRange(ultimaFilaDatos + 1, 1, lastRow - ultimaFilaDatos, 28).clearFormat();
    for (var fila = lastRow; fila > ultimaFilaDatos; fila--) {
      try { sheet.deleteRow(fila); } catch (e) {}
    }
  }
}

// Mantener compatibilidad con v3
function colocarContadorYLeyenda(sheet) {
  colocarContadorYLeyendaCosmica(sheet);
}

function formatearFila(sheet, fila) {
  formatearFilaCosmica(sheet, fila);
}

function formatearFilaReporte(sheet, fila) {
  formatearFilaReporteCosmica(sheet, fila);
}

function formatearFilaSoftware(sheet, fila) {
  formatearFilaSoftwareCosmica(sheet, fila);
}

function ordenarPorInventario(sheet) {
  var lastRow = obtenerUltimaFilaDatos(sheet);
  if (lastRow <= 5) return;

  var rango = sheet.getRange(5, 1, lastRow - 4, 28);
  rango.sort({ column: 3, ascending: false });

  var numFilas = lastRow - 4;
  for (var i = 0; i < numFilas; i++) {
    sheet.getRange(5 + i, 1).setValue(i + 1);
  }
}

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
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    for (var row = 1; row < data.length; row++) {
      for (var col = 0; col < data[row].length; col++) {
        var cellValue = String(data[row][col]).trim().toUpperCase();
        var searchValue = String(numeroDeSerie).trim().toUpperCase();

        if (cellValue == searchValue) {
          var numSI = data[row][0];
          return "SI #" + numSI;
        }
      }
    }
    return "NO ENCONTRADO";
  } catch (error) {
    return "Error FAA: " + error.message;
  }
}

function testBuscarFAA() {
  var resultado = buscarEnSeriesFAA("MZ02W5K4");
  Logger.log("Resultado: " + resultado);
  SpreadsheetApp.getUi().alert("Resultado: " + resultado);
}

// =============================================================================
// FUNCIONES MANUALES - Ejecutar desde el menu
// =============================================================================

function formatearTodasLasFilas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = obtenerUltimaFilaDatos(sheet);
  for (var fila = 5; fila <= lastRow; fila++) {
    formatearFilaCosmica(sheet, fila);
  }
  SpreadsheetApp.getUi().alert(SYM.STAR + " Todas las filas han sido formateadas con el tema cosmico");
}

function ordenarManual() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  ordenarPorInventario(sheet);
  SpreadsheetApp.getUi().alert(SYM.STAR + " Ordenado por inventario (mayor a menor)");
}

function limpiarManual() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  limpiarYFormatear(sheet);
  SpreadsheetApp.getUi().alert(SYM.STAR + " Hoja limpiada y formateada con tema cosmico");
}

function limpiarYFormatear(sheet) {
  limpiarExtras(sheet);
  var ultimaFila = obtenerUltimaFilaDatos(sheet);
  var numEquipos = ultimaFila - 4;

  for (var fila = 5; fila <= ultimaFila; fila++) {
    formatearFilaCosmica(sheet, fila);
  }

  for (var i = 0; i < numEquipos; i++) {
    sheet.getRange(5 + i, 1).setValue(i + 1);
  }

  ordenarPorInventario(sheet);
  colocarContadorYLeyendaCosmica(sheet);

  return jsonResponse({
    status: "OK",
    mensaje: "Hoja limpiada y formateada con tema cosmico",
    equipos: numEquipos
  });
}

function formatearHojaCompleta(sheet) {
  return limpiarYFormatear(sheet);
}
