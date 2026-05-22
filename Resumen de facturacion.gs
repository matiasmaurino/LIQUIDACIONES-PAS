function consolidarResumenMensual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ORIGEN: Datos detallados de las liquidaciones actuales
  var hojaDetalle = ss.getSheetByName("LIQUIDACIONES AGRUPADAS");
  if (!hojaDetalle) {
    console.warn("No se encontró la solapa 'LIQUIDACIONES AGRUPADAS'.");
    return;
  }
  var datosDetalle = hojaDetalle.getDataRange().getValues();
  if (datosDetalle.length <= 1) {
    console.warn("La hoja 'LIQUIDACIONES AGRUPADAS' no contiene filas de datos.");
    return;
  }
  
  // Mapeo dinámico basado en los encabezados reales de la hoja de origen (LIQUIDACIONES AGRUPADAS)
  var encabezadosDetalle = datosDetalle[0];
  var idxFecha = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "FECHA"); // Col B
  var idxPrima = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "PRIMA"); // Col F
  var idxPremio = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "PREMIO"); // Col G
  var idxComision = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "COMISION"); // Col H
  var idxCia = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "CIA"); // Col J
  var idxPas = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "PAS AGRUPADO"); // Col K
  var idxRamo = encabezadosDetalle.findIndex(h => h.toString().trim().toUpperCase() === "RAMO_NOMBRE"); // Col L
  
  if (idxFecha === -1 || idxPrima === -1 || idxPremio === -1 || idxComision === -1 || idxCia === -1 || idxPas === -1 || idxRamo === -1) {
    throw new Error("Faltan columnas críticas en 'LIQUIDACIONES AGRUPADAS'. Verificá que existan: FECHA, PRIMA, PREMIO, COMISION, CIA, PAS AGRUPADO y RAMO_NOMBRE.");
  }
  
  // Procesar y agrupar las filas en memoria
  var mapaNovedades = {};
  
  for (var i = 1; i < datosDetalle.length; i++) {
    var fila = datosDetalle[i];
    
    var rawPas = fila[idxPas] ? fila[idxPas].toString().trim().toUpperCase() : "";
    var rawFecha = fila[idxFecha];
    var ramo = fila[idxRamo] ? fila[idxRamo].toString().trim() : "";
    var cia = fila[idxCia] ? fila[idxCia].toString().trim() : "";
    
    var prima = parseFloat(fila[idxPrima]) || 0;
    var premio = parseFloat(fila[idxPremio]) || 0;
    var comision = parseFloat(fila[idxComision]) || 0;
    
    if (!rawFecha || rawFecha === "FECHA" || rawFecha === "") continue;
    
    // Traducción normalizada del PAS
    var pasTraducido = "";
    if (rawPas.includes("DANIEL GUILLERMO") || rawPas === "DGM") {
      pasTraducido = "DGM";
    } else if (rawPas.includes("MATIAS") || rawPas === "MATIAS") {
      pasTraducido = "MATIAS";
    } else {
      pasTraducido = rawPas; 
    }
    
    // Convertir la fecha a formato YYYY-MM
    var fechaStr = "";
    if (rawFecha instanceof Date) {
      fechaStr = Utilities.formatDate(rawFecha, ss.getSpreadsheetTimeZone(), "yyyy-MM");
    } else {
      fechaStr = rawFecha.toString().substring(0, 7);
    }
    
    if (fechaStr.length !== 7 || !fechaStr.includes("-")) continue;
    
    var clave = fechaStr + "_" + pasTraducido + "_" + cia + "_" + ramo;
    
    if (!mapaNovedades[clave]) {
      mapaNovedades[clave] = {
        fecha: fechaStr,
        cia: cia,
        pas: pasTraducido,
        ramo: ramo,
        primaTotal: 0,
        premioTotal: 0,
        comisionTotal: 0,
        cantidadPolizas: 0
      };
    }
    
    mapaNovedades[clave].primaTotal += prima;
    mapaNovedades[clave].premioTotal += premio;
    mapaNovedades[clave].comisionTotal += comision;
    mapaNovedades[clave].cantidadPolizas += 1; 
  }
  
  // Transformar a matriz estructurada de 8 columnas (CON EL ORDEN DE PAS Y CIA CORREGIDO)
  var filasNuevas = [];
  for (var c in mapaNovedades) {
    var item = mapaNovedades[c];
    
    filasNuevas.push([
      item.fecha,            // Col A: Año-mes
      item.pas,              // Col B: PAS <-- CORREGIDO (Antes estaba item.cia)
       item.ramo,             // Col D: RAMO_NOMBRE
      item.cia,              // Col C: CIA <-- CORREGIDO (Antes estaba item.pas)
      item.primaTotal,       // Col E: PRIMA
      item.premioTotal,      // Col F: PREMIO
      item.comisionTotal,    // Col G: COMISION
      item.cantidadPolizas   // Col H: CANT POLIZAS
    ]);
  }
  
  // Ordenar el nuevo bloque por Fecha, PAS y CIA para homogeneidad
  filasNuevas.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    if (a[1] !== b[1]) return a[1].localeCompare(b[1]);
    if (a[2] !== b[2]) return a[2].localeCompare(b[2]);
    return a[3].localeCompare(b[3]);
  });
  
  // 2. DESTINO: Tu hoja real histórica (LIQUIDACIONES AGRUPADAS POR MES)
  var hojaDestino = ss.getSheetByName("LIQUIDACIONES AGRUPADAS POR MES");
  if (!hojaDestino) {
    throw new Error("No se encontró la hoja 'LIQUIDACIONES AGRUPADAS POR MES'.");
  }
  
  // Detectar el final de los datos para no pisar el histórico manual de arriba
  var ultimaFila = hojaDestino.getLastRow();
  if (ultimaFila === 0) {
    var encabezadosDestino = ["Año-mes", "PAS", "CIA", "RAMO_NOMBRE", "PRIMA", "PREMIO", "COMISION", "CANT POLIZAS"];
    hojaDestino.appendRow(encabezadosDestino);
    ultimaFila = 1;
  }
  
  // 3. INSERCIÓN PURA AL FINAL
  if (filasNuevas.length > 0) {
    var rangoDestino = hojaDestino.getRange(ultimaFila + 1, 1, filasNuevas.length, filasNuevas[0].length);
    rangoDestino.setValues(filasNuevas);
    
    // Formatos visuales de Moneda y Números enteros
    hojaDestino.getRange(ultimaFila + 1, 5, filasNuevas.length, 3).setNumberFormat("$#,##0"); 
    hojaDestino.getRange(ultimaFila + 1, 8, filasNuevas.length, 1).setNumberFormat("#,##0");  
    
    console.log("¡Hecho! Se acoplaron con éxito " + filasNuevas.length + " filas nuevas al final del histórico.");
  } else {
    console.log("No se detectaron novedades válidas para ingresar.");
  }
}