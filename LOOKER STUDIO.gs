function actualizarHistoricoDetalleLooker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaOrigen = "LIQUIDACIONES AGRUPADAS";
  const nombreHojaDestino = "HISTORICO_DETALLE_LOOKER";
  
  let hojaOrigen = ss.getSheetByName(nombreHojaOrigen);
  if (!hojaOrigen) {
    console.error("No se encontró la hoja de origen: " + nombreHojaOrigen);
    return;
  }
  
  let datosOrigen = hojaOrigen.getDataRange().getValues();
  if (datosOrigen.length <= 1) {
    console.log("La hoja de origen está vacía o solo tiene cabecera.");
    return;
  }
  
  // 1. IDENTIFICAR POSICIONES DE LOS ENCABEZADOS EN LA HOJA ORIGEN
  let encabezadosOrigen = datosOrigen[0].map(h => String(h).trim().toUpperCase());
  
  // Definimos las ÚNICAS 7 columnas que van a existir en Looker Studio
  let columnasDeseadas = ["FECHA", "NOMBRE", "RAMO", "POLIZA", "COMISION", "CIA", "PAS AGRUPADO","RAMO_NOMBRE"];
  
  // Buscamos los índices de cada columna en la hoja de origen
  let indicesColumnas = columnasDeseadas.map(col => encabezadosOrigen.indexOf(col));
  
  // La cabecera ahora tiene estrictamente 7 elementos
  let cabeceraDestino = [...columnasDeseadas];
  let matrizFinalLooker = [cabeceraDestino];
  
  // 2. PROCESAR Y FILTRAR TODAS LAS FILAS EN MEMORIA
  for (let i = 1; i < datosOrigen.length; i++) {
    let filaOriginal = datosOrigen[i];
    let filaFiltrada = [];
    
    // Agregamos solo las 7 columnas deseadas en el orden correcto
    indicesColumnas.forEach(idx => {
      filaFiltrada.push(idx !== -1 ? filaOriginal[idx] : "");
    });
    
    matrizFinalLooker.push(filaFiltrada);
  }
  
  // 3. SOBREESCRIBIR POR COMPLETO LA HOJA DESTINO
  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(nombreHojaDestino);
  }
  
  // Limpiamos todo el contenido viejo y formatos previos
  hojaDestino.clear();
  
  // Escribimos la nueva matriz de datos limpia (ahora de 7 columnas de ancho)
  hojaDestino.getRange(1, 1, matrizFinalLooker.length, matrizFinalLooker[0].length).setValues(matrizFinalLooker);
  
  // 4. APLICAR FORMATOS AUTOMÁTICOS A LAS NUEVAS COLUMNAS
  const totalFilas = hojaDestino.getLastRow();
  if (totalFilas > 1) {
    let colFechaDestino = cabeceraDestino.indexOf("FECHA") + 1;
    let colComisionDestino = cabeceraDestino.indexOf("COMISION") + 1;
    
    // Formato de Fecha (dd/mm/yyyy)
    if (colFechaDestino > 0) {
      hojaDestino.getRange(2, colFechaDestino, totalFilas - 1).setNumberFormat("dd/mm/yyyy");
    }
    // Formato de Comisión sin decimales (#,##0)
    if (colComisionDestino > 0) {
      hojaDestino.getRange(2, colComisionDestino, totalFilas - 1).setNumberFormat("#,##0");
    }
    
    // Ajuste automático de ancho para las 7 columnas resultantes
    hojaDestino.autoResizeColumns(1, cabeceraDestino.length);
  }
  
  console.log("🚀 Reemplazo total completado desde LIQUIDACIONES AGRUPADAS. Estructura pura de 7 columnas para Looker.");
}