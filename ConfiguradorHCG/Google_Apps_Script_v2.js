// =============================================================================
// GOOGLE APPS SCRIPT - REGISTRO DE EQUIPOS HCG v2.0
// =============================================================================
// Copiar este codigo en: Extensiones > Apps Script
// Luego: Implementar > Nueva implementacion > Aplicacion web
// Ejecutar como: Yo mismo | Acceso: Cualquier persona
// =============================================================================

// Configuracion
const HOJA_REGISTRO = "Registro";
const HOJA_SERIES_FAA = "SeriesFAA";
const FILA_ENCABEZADOS = 4;
const FILA_DATOS_INICIO = 5;

// Columnas (ajustar segun tu hoja)
const COL = {
  No: 1,
  Fecha: 2,
  InvST: 3,
  Marca: 4,
  Modelo: 5,
  NoSerie: 6,
  Procesador: 7,
  Nucleos: 8,
  RAM: 9,
  Disco: 10,
  Graficos: 11,
  WiFi: 12,
  BT: 13,
  SO: 14,
  MACEthernet: 15,
  MACWiFi: 16,
  ProductKey: 17,
  Fab: 18,
  Garantia: 19,
  Ubicacion: 20,
  Departamento: 21,
  Usuario: 22,
  Estado: 23,
  FAA: 24
};

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const accion = datos.Accion || "crear";

    if (accion === "crear") {
      return crearRegistro(datos);
    } else if (accion === "actualizar") {
      return actualizarRegistro(datos);
    }

    return jsonResponse({ status: "ERROR", message: "Accion no reconocida" });
  } catch (error) {
    return jsonResponse({ status: "ERROR", message: error.toString() });
  }
}

function doGet(e) {
  return jsonResponse({
    status: "OK",
    message: "API de Registro de Equipos HCG v2.0",
    timestamp: new Date().toLocaleString("es-MX")
  });
}

function crearRegistro(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_REGISTRO);

  // Buscar si ya existe el equipo por InvST o Serie
  const filaExistente = buscarEquipo(hoja, datos.InvST, datos.Serie);
  if (filaExistente > 0) {
    // Actualizar la fila existente con los nuevos datos
    hoja.getRange(filaExistente, COL.Estado).setValue("En proceso");
    formatearFila(hoja, filaExistente);

    return jsonResponse({
      status: "OK",
      message: "Equipo ya existia, actualizado a En proceso",
      row: filaExistente,
      faa: hoja.getRange(filaExistente, COL.FAA).getValue() || "",
      faaEncontrado: hoja.getRange(filaExistente, COL.FAA).getValue() ? true : false,
      action: "updated",
      inventario: datos.InvST,
      serie: datos.Serie
    });
  }

  // Buscar en lista FAA
  const resultadoFAA = buscarEnSeriesFAA(datos.Serie);
  const faaEncontrado = resultadoFAA && resultadoFAA !== "";

  // Obtener siguiente numero
  const ultimaFila = hoja.getLastRow();
  let siguienteNo = 1;
  if (ultimaFila >= FILA_DATOS_INICIO) {
    const numeros = hoja.getRange(FILA_DATOS_INICIO, COL.No, ultimaFila - FILA_DATOS_INICIO + 1, 1).getValues();
    siguienteNo = Math.max(...numeros.flat().filter(n => !isNaN(n))) + 1;
  }

  // Preparar fila con todos los datos
  const nuevaFila = new Array(24).fill("");

  nuevaFila[COL.No - 1] = siguienteNo;
  nuevaFila[COL.Fecha - 1] = datos.Fecha || Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy");
  nuevaFila[COL.InvST - 1] = formatearInventario(datos.InvST);
  nuevaFila[COL.Marca - 1] = datos.Marca || "Lenovo";
  nuevaFila[COL.Modelo - 1] = datos.Modelo || "ThinkCentre M70s Gen 5";
  nuevaFila[COL.NoSerie - 1] = datos.Serie || "";
  nuevaFila[COL.Procesador - 1] = formatearProcesador(datos.Procesador);
  nuevaFila[COL.Nucleos - 1] = datos.Nucleos || 14;
  nuevaFila[COL.RAM - 1] = formatearRAM(datos.RAM);
  nuevaFila[COL.Disco - 1] = formatearDisco(datos.Disco, datos.DiscoTipo);
  nuevaFila[COL.Graficos - 1] = datos.Graficos || "Intel UHD 770";
  nuevaFila[COL.WiFi - 1] = datos.WiFi || "Wi-Fi 6";
  nuevaFila[COL.BT - 1] = datos.BT || "5.1";
  nuevaFila[COL.SO - 1] = datos.SO || "Win 11 Pro";
  nuevaFila[COL.MACEthernet - 1] = formatearMAC(datos.MACEthernet);
  nuevaFila[COL.MACWiFi - 1] = formatearMAC(datos.MACWiFi);
  nuevaFila[COL.ProductKey - 1] = datos.ProductKey || "";
  nuevaFila[COL.Fab - 1] = datos.FechaFab || datos.Fecha;
  nuevaFila[COL.Garantia - 1] = datos.Garantia || calcularGarantia(datos.Fecha);
  nuevaFila[COL.Ubicacion - 1] = datos.Ubicacion || "";
  nuevaFila[COL.Departamento - 1] = datos.Departamento || "";
  nuevaFila[COL.Usuario - 1] = datos.Usuario || "";
  nuevaFila[COL.Estado - 1] = "En proceso";
  nuevaFila[COL.FAA - 1] = resultadoFAA;

  // Insertar fila
  const nuevaFilaNum = ultimaFila + 1;
  hoja.getRange(nuevaFilaNum, 1, 1, nuevaFila.length).setValues([nuevaFila]);

  // Aplicar formato a la nueva fila
  formatearFila(hoja, nuevaFilaNum);

  // Ordenar por inventario (mayor a menor)
  ordenarPorInventario(hoja);

  // Buscar la fila final despues de ordenar
  const filaFinal = buscarEquipo(hoja, datos.InvST, datos.Serie);

  // Actualizar contador
  actualizarContador(hoja);

  return jsonResponse({
    status: "OK",
    message: faaEncontrado ? "Equipo registrado - Serie ENCONTRADA en FAA" : "Equipo registrado - Serie NO encontrada en FAA",
    row: filaFinal > 0 ? filaFinal : nuevaFilaNum,
    faa: resultadoFAA,
    faaEncontrado: faaEncontrado,
    action: "created",
    inventario: datos.InvST,
    serie: datos.Serie
  });
}

function actualizarRegistro(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_REGISTRO);

  // Buscar el equipo
  const fila = buscarEquipo(hoja, datos.InvST, datos.Serie);

  if (fila <= 0) {
    return jsonResponse({ status: "ERROR", message: "Equipo no encontrado" });
  }

  // Actualizar estado a Activo
  hoja.getRange(fila, COL.Estado).setValue("Activo");

  // Actualizar otros campos si se proporcionan
  if (datos.Ubicacion) hoja.getRange(fila, COL.Ubicacion).setValue(datos.Ubicacion);
  if (datos.Departamento) hoja.getRange(fila, COL.Departamento).setValue(datos.Departamento);
  if (datos.Usuario) hoja.getRange(fila, COL.Usuario).setValue(datos.Usuario);

  // Aplicar formato de activo (fondo verde claro)
  const rangoFila = hoja.getRange(fila, 1, 1, 24);
  rangoFila.setBackground("#E8F5E9");

  // Actualizar contador
  actualizarContador(hoja);

  return jsonResponse({
    status: "OK",
    message: "Equipo actualizado a Activo",
    row: fila,
    action: "updated"
  });
}

function buscarEquipo(hoja, invST, serie) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < FILA_DATOS_INICIO) return -1;

  const datos = hoja.getRange(FILA_DATOS_INICIO, 1, ultimaFila - FILA_DATOS_INICIO + 1, 24).getValues();

  for (let i = 0; i < datos.length; i++) {
    const invSTActual = String(datos[i][COL.InvST - 1]).replace(/\D/g, "");
    const serieActual = String(datos[i][COL.NoSerie - 1]);

    if ((invST && invSTActual === String(invST)) || (serie && serieActual === serie)) {
      return FILA_DATOS_INICIO + i;
    }
  }

  return -1;
}

function buscarEnSeriesFAA(serie) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaFAA = ss.getSheetByName(HOJA_SERIES_FAA);

    if (!hojaFAA) return "";

    const datos = hojaFAA.getDataRange().getValues();

    for (let i = 0; i < datos.length; i++) {
      for (let j = 0; j < datos[i].length; j++) {
        if (String(datos[i][j]).toUpperCase() === String(serie).toUpperCase()) {
          // Buscar el numero FAA en la misma fila o columna cercana
          // Asumiendo que el formato es "SI #XXX" o similar
          const fila = datos[i];
          for (let k = 0; k < fila.length; k++) {
            const celda = String(fila[k]);
            if (celda.match(/SI\s*#?\d+/i)) {
              return celda;
            }
          }
          return "SI #" + (i + 1);
        }
      }
    }

    return "";
  } catch (e) {
    return "";
  }
}

// =============================================================================
// FUNCIONES DE FORMATO
// =============================================================================

function formatearInventario(inv) {
  if (!inv) return "";
  const num = String(inv).replace(/\D/g, "");
  return num.padStart(5, "0");
}

function formatearMAC(mac) {
  if (!mac) return "";
  // Quitar separadores existentes y formatear con dos puntos
  const limpia = String(mac).replace(/[^A-Fa-f0-9]/g, "").toUpperCase();
  if (limpia.length !== 12) return mac;
  return limpia.match(/.{2}/g).join(":");
}

function formatearProcesador(proc) {
  if (!proc) return "Intel Core i5-14500 vPro";
  // Limpiar y formatear nombre del procesador
  let limpio = String(proc)
    .replace(/\(R\)/gi, "")
    .replace(/\(TM\)/gi, "")
    .replace(/CPU/gi, "")
    .replace(/@.*/, "")
    .replace(/\s+/g, " ")
    .trim();
  return limpio || "Intel Core i5-14500 vPro";
}

function formatearRAM(ram) {
  if (!ram) return "8 GB DDR5";
  const num = String(ram).replace(/\D/g, "");
  if (num) {
    return num + " GB DDR5";
  }
  return ram;
}

function formatearDisco(disco, tipo) {
  if (!disco) return "512 GB SSD";
  const num = String(disco).replace(/\D/g, "");
  const tipoStr = tipo || "SSD";
  if (num) {
    return num + " GB " + tipoStr;
  }
  return disco;
}

function calcularGarantia(fechaCompra) {
  const fecha = fechaCompra ? parseDate(fechaCompra) : new Date();
  fecha.setFullYear(fecha.getFullYear() + 3);
  return Utilities.formatDate(fecha, "America/Mexico_City", "dd/MM/yyyy");
}

function parseDate(dateStr) {
  if (!dateStr) return new Date();
  const parts = String(dateStr).split("/");
  if (parts.length === 3) {
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateStr);
}

function formatearFila(hoja, fila) {
  const rango = hoja.getRange(fila, 1, 1, 24);

  // Bordes
  rango.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);

  // Alineacion
  rango.setHorizontalAlignment("center");
  rango.setVerticalAlignment("middle");

  // Columnas especificas
  hoja.getRange(fila, COL.Procesador).setHorizontalAlignment("left");
  hoja.getRange(fila, COL.ProductKey).setFontSize(9);

  // Color segun estado
  const estado = hoja.getRange(fila, COL.Estado).getValue();
  if (estado === "Activo") {
    rango.setBackground("#E8F5E9"); // Verde claro
  } else if (estado === "En proceso") {
    rango.setBackground("#FFF3E0"); // Naranja claro
  } else if (estado === "Baja") {
    rango.setBackground("#FFEBEE"); // Rojo claro
  }
}

function ordenarPorInventario(hoja) {
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= FILA_DATOS_INICIO) return;

  const rango = hoja.getRange(FILA_DATOS_INICIO, 1, ultimaFila - FILA_DATOS_INICIO + 1, 24);

  // Ordenar por columna Inv. ST de mayor a menor (descendente)
  rango.sort({ column: COL.InvST, ascending: false });

  // Renumerar la columna No.
  const numFilas = ultimaFila - FILA_DATOS_INICIO + 1;
  const numeros = [];
  for (let i = 1; i <= numFilas; i++) {
    numeros.push([i]);
  }
  hoja.getRange(FILA_DATOS_INICIO, COL.No, numFilas, 1).setValues(numeros);
}

function actualizarContador(hoja) {
  const ultimaFila = hoja.getLastRow();
  const totalEquipos = ultimaFila >= FILA_DATOS_INICIO ? ultimaFila - FILA_DATOS_INICIO + 1 : 0;

  // Contar activos
  let activos = 0;
  let enProceso = 0;

  if (totalEquipos > 0) {
    const estados = hoja.getRange(FILA_DATOS_INICIO, COL.Estado, totalEquipos, 1).getValues();
    estados.forEach(e => {
      if (e[0] === "Activo") activos++;
      else if (e[0] === "En proceso") enProceso++;
    });
  }

  const pendientes = 150 - totalEquipos; // Meta de 150 equipos
  const progreso = Math.round((totalEquipos / 150) * 100);
  const fechaActualizacion = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy");

  // Actualizar celda de resumen (ajustar posicion segun tu hoja)
  const resumen = `Total: ${totalEquipos} de 150 equipos | Activos: ${activos} | En proceso: ${enProceso} | Pendientes: ${pendientes} | Progreso: ${progreso}% | Actualizado: ${fechaActualizacion}`;

  // Buscar celda de resumen existente o crear en fila 17
  hoja.getRange("N17").setValue(resumen);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// FUNCIONES MANUALES (Ejecutar desde el editor)
// =============================================================================

function ordenarHojaManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_REGISTRO);
  ordenarPorInventario(hoja);
  SpreadsheetApp.getUi().alert("Hoja ordenada por inventario (mayor a menor)");
}

function formatearTodasLasFilas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_REGISTRO);
  const ultimaFila = hoja.getLastRow();

  for (let fila = FILA_DATOS_INICIO; fila <= ultimaFila; fila++) {
    formatearFila(hoja, fila);
  }

  SpreadsheetApp.getUi().alert("Todas las filas han sido formateadas");
}

function actualizarContadorManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_REGISTRO);
  actualizarContador(hoja);
  SpreadsheetApp.getUi().alert("Contador actualizado");
}

// =============================================================================
// CONFIGURACION INICIAL - EJECUTAR UNA VEZ
// =============================================================================

function configurarHojaCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJA_REGISTRO);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_REGISTRO);
  }

  // Titulo principal (Fila 1)
  hoja.getRange("A1:X1").merge();
  hoja.getRange("A1").setValue("HOSPITAL CIVIL DE GUADALAJARA - FRAY ANTONIO ALCALDE");
  hoja.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  hoja.getRange("A1").setBackground("#1565C0").setFontColor("#FFFFFF");

  // Subtitulo (Fila 2)
  hoja.getRange("A2:X2").merge();
  hoja.getRange("A2").setValue("REGISTRO DE EQUIPOS DE COMPUTO - COORDINACION GENERAL DE INFORMATICA - EXT. 54425");
  hoja.getRange("A2").setFontSize(11).setHorizontalAlignment("center");
  hoja.getRange("A2").setBackground("#1976D2").setFontColor("#FFFFFF");

  // Fila 3 vacia como separador
  hoja.setRowHeight(3, 10);

  // Encabezados de columnas (Fila 4)
  const encabezados = [
    "No.",           // A - Numero consecutivo
    "Fecha",         // B - Fecha de registro
    "Inv. ST",       // C - Numero de inventario
    "Marca",         // D - Fabricante
    "Modelo",        // E - Modelo del equipo
    "No. Serie",     // F - Numero de serie
    "Procesador",    // G - CPU
    "Nucleos",       // H - Nucleos del CPU
    "RAM",           // I - Memoria RAM
    "Disco",         // J - Almacenamiento
    "Graficos",      // K - Tarjeta grafica
    "WiFi",          // L - Version WiFi
    "BT",            // M - Version Bluetooth
    "S.O.",          // N - Sistema operativo
    "MAC Ethernet",  // O - Direccion MAC Ethernet
    "MAC WiFi",      // P - Direccion MAC WiFi
    "Product Key",   // Q - Clave de Windows
    "Fab.",          // R - Fecha de fabricacion
    "Garantia",      // S - Fecha fin garantia
    "Ubicacion",     // T - Ubicacion fisica
    "Departamento",  // U - Departamento asignado
    "Usuario",       // V - Usuario asignado
    "Estado",        // W - Estado del equipo
    "FAA"            // X - Referencia compra FAA
  ];

  hoja.getRange(FILA_ENCABEZADOS, 1, 1, encabezados.length).setValues([encabezados]);

  // Formato de encabezados
  const rangoEncabezados = hoja.getRange(FILA_ENCABEZADOS, 1, 1, encabezados.length);
  rangoEncabezados.setFontWeight("bold");
  rangoEncabezados.setHorizontalAlignment("center");
  rangoEncabezados.setVerticalAlignment("middle");
  rangoEncabezados.setBackground("#0D47A1");
  rangoEncabezados.setFontColor("#FFFFFF");
  rangoEncabezados.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Agregar notas/comentarios a cada encabezado
  const descripciones = [
    "Numero consecutivo automatico",
    "Fecha de registro del equipo (dd/mm/aaaa)",
    "Numero de inventario de Soporte Tecnico (5 digitos)",
    "Fabricante del equipo (ej: Lenovo, HP, Dell)",
    "Modelo especifico del equipo",
    "Numero de serie unico del fabricante",
    "Modelo del procesador (CPU)",
    "Cantidad de nucleos del procesador",
    "Cantidad de memoria RAM instalada",
    "Capacidad y tipo de disco (SSD/HDD)",
    "Tarjeta grafica o GPU integrada",
    "Version de WiFi soportada",
    "Version de Bluetooth",
    "Sistema operativo instalado",
    "Direccion MAC de la tarjeta Ethernet",
    "Direccion MAC de la tarjeta WiFi",
    "Clave de producto de Windows (OEM)",
    "Fecha de fabricacion del equipo",
    "Fecha de vencimiento de garantia",
    "Ubicacion fisica (edificio, piso, area)",
    "Departamento al que esta asignado",
    "Nombre del usuario responsable",
    "Estado actual: Activo, En proceso, Sin asignar, Baja",
    "Numero de referencia de compra FAA (SI #XXX)"
  ];

  for (let i = 0; i < descripciones.length; i++) {
    hoja.getRange(FILA_ENCABEZADOS, i + 1).setNote(descripciones[i]);
  }

  // Anchos de columna optimizados
  const anchos = [40, 90, 60, 70, 160, 120, 180, 60, 80, 90, 100, 60, 40, 80, 130, 130, 180, 85, 85, 120, 120, 100, 80, 70];
  for (let i = 0; i < anchos.length; i++) {
    hoja.setColumnWidth(i + 1, anchos[i]);
  }

  // Congelar filas de encabezado
  hoja.setFrozenRows(FILA_ENCABEZADOS);

  // Crear leyenda de estados (despues de los datos)
  crearLeyenda(hoja);

  SpreadsheetApp.getUi().alert("Hoja configurada correctamente con encabezados y descripciones");
}

function crearLeyenda(hoja) {
  // Leyenda de estados en la fila 19 (o despues de los datos)
  const filaLeyenda = 19;

  hoja.getRange(filaLeyenda, 1).setValue("LEYENDA DE ESTADOS:");
  hoja.getRange(filaLeyenda, 1).setFontWeight("bold");

  // Activo
  hoja.getRange(filaLeyenda + 1, 1).setValue("Activo");
  hoja.getRange(filaLeyenda + 1, 1).setBackground("#E8F5E9");
  hoja.getRange(filaLeyenda + 1, 2).setValue("Equipo operativo y funcionando");

  // Almac
  hoja.getRange(filaLeyenda + 2, 1).setValue("Almac");
  hoja.getRange(filaLeyenda + 2, 1).setBackground("#FFF9C4");
  hoja.getRange(filaLeyenda + 2, 2).setValue("Sin asignar / En almacen");

  // Baja
  hoja.getRange(filaLeyenda + 3, 1).setValue("Baja");
  hoja.getRange(filaLeyenda + 3, 1).setBackground("#FFCDD2");
  hoja.getRange(filaLeyenda + 3, 2).setValue("Dado de baja / Fuera de servicio");

  // En proceso
  hoja.getRange(filaLeyenda + 4, 1).setValue("En proceso");
  hoja.getRange(filaLeyenda + 4, 1).setBackground("#FFF3E0");
  hoja.getRange(filaLeyenda + 4, 2).setValue("Configuracion en progreso");
}

function crearHojaDescripciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Descripcion_Campos");

  if (!hoja) {
    hoja = ss.insertSheet("Descripcion_Campos");
  } else {
    hoja.clear();
  }

  // Titulo
  hoja.getRange("A1:C1").merge();
  hoja.getRange("A1").setValue("DESCRIPCION DE CAMPOS - REGISTRO DE EQUIPOS");
  hoja.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#1565C0").setFontColor("#FFFFFF");

  // Encabezados
  hoja.getRange("A3:C3").setValues([["COLUMNA", "CAMPO", "DESCRIPCION"]]);
  hoja.getRange("A3:C3").setFontWeight("bold").setBackground("#E3F2FD");

  // Datos
  const campos = [
    ["A", "No.", "Numero consecutivo automatico asignado a cada equipo"],
    ["B", "Fecha", "Fecha en que se registro el equipo en el sistema (formato: dd/mm/aaaa)"],
    ["C", "Inv. ST", "Numero de inventario de Soporte Tecnico - 5 digitos que identifican el equipo"],
    ["D", "Marca", "Fabricante del equipo (ej: Lenovo, HP, Dell, Acer)"],
    ["E", "Modelo", "Modelo especifico del equipo (ej: ThinkCentre M70s Gen 5)"],
    ["F", "No. Serie", "Numero de serie unico asignado por el fabricante - Se encuentra en la etiqueta del equipo"],
    ["G", "Procesador", "Modelo del procesador/CPU instalado (ej: Intel Core i5-14500 vPro)"],
    ["H", "Nucleos", "Cantidad de nucleos fisicos del procesador"],
    ["I", "RAM", "Cantidad de memoria RAM instalada (ej: 8 GB DDR5)"],
    ["J", "Disco", "Capacidad y tipo de almacenamiento (ej: 512 GB SSD)"],
    ["K", "Graficos", "Tarjeta grafica o GPU (ej: Intel UHD 770)"],
    ["L", "WiFi", "Version de WiFi soportada (ej: Wi-Fi 6)"],
    ["M", "BT", "Version de Bluetooth (ej: 5.1)"],
    ["N", "S.O.", "Sistema operativo instalado (ej: Win 11 Pro)"],
    ["O", "MAC Ethernet", "Direccion MAC de la tarjeta de red Ethernet - Identificador unico de red"],
    ["P", "MAC WiFi", "Direccion MAC de la tarjeta WiFi - Identificador unico de red inalambrica"],
    ["Q", "Product Key", "Clave de producto de Windows (licencia OEM del fabricante)"],
    ["R", "Fab.", "Fecha de fabricacion del equipo"],
    ["S", "Garantia", "Fecha de vencimiento de la garantia del fabricante"],
    ["T", "Ubicacion", "Ubicacion fisica del equipo (edificio, piso, area, consultorio)"],
    ["U", "Departamento", "Departamento o servicio al que esta asignado el equipo"],
    ["V", "Usuario", "Nombre del usuario responsable del equipo"],
    ["W", "Estado", "Estado actual del equipo: Activo, En proceso, Almac (sin asignar), Baja"],
    ["X", "FAA", "Numero de referencia de la compra FAA (SI #XXX) - Del archivo seriesFAA.xlsx"]
  ];

  hoja.getRange(4, 1, campos.length, 3).setValues(campos);

  // Formato
  hoja.setColumnWidth(1, 70);
  hoja.setColumnWidth(2, 120);
  hoja.setColumnWidth(3, 500);

  // Bordes
  hoja.getRange(3, 1, campos.length + 1, 3).setBorder(true, true, true, true, true, true);

  SpreadsheetApp.getUi().alert("Hoja de descripciones creada");
}
