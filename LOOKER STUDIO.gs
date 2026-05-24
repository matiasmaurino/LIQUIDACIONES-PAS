/**
 * FUNCIÓN ULTRA-RESILIENTE Y DINÁMICA: Sincroniza el detalle analítico hacia Looker.
 * Se adapta automáticamente al tamaño real de columnas (12, 13 o más) que tengan 
 * tus hojas de origen, evitando errores de desajuste de intervalos.
 */
function actualizarHistoricoDetalleLooker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombresHojasOrigen = ["FP", "RIV", "PS"];
  const nombreHojaDestino = "HISTORICO_DETALLE_LOOKER";
  
  let todosLosDatosEntrantes = [];
  let idsLoteNuevos = {};
  let cabeceraEstandar = [];

  // 1. ESCANEAR Y RECOLECTAR NOVEDADES DE LAS PESTAÑAS INDIVIDUALES
  nombresHojasOrigen.forEach(nombre => {
    let hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    
    let datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return; // Vacía o sólo cabecera

    // Guardamos la cabecera real de la hoja para saber la estructura
    if (cabeceraEstandar.length === 0 || datos[0].length > cabeceraEstandar.length) {
      cabeceraEstandar = datos[0];
    }

    // Buscamos dinámicamente en qué columna está el "ID_PROCESAMIENTO"
    let idxIdProcesamiento = datos[0].indexOf("ID_PROCESAMIENTO");
    if (idxIdProcesamiento === -1) {
      // Si no encuentra el texto, por defecto es la última columna
      idxIdProcesamiento = datos[0].length - 1;
    }

    for (let i = 1; i < datos.length; i++) {
      let fila = datos[i];
      let idFila = fila[idxIdProcesamiento]; 
      
      if (idFila && idFila.toString().trim() !== "") {
        idsLoteNuevos[idFila.toString().trim()] = true;
        todosLosDatosEntrantes.push(fila); 
      }
    }
  });

  if (todosLosDatosEntrantes.length === 0) {
    console.log("⚠️ No se encontraron IDs de lote válidos en las pestañas de origen.");
    SpreadsheetApp.getUi().alert("⚠️ Atención", "No se detectaron datos listos para subir en las hojas de origen.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Determinamos el ancho real final en base a la cabecera detectada
  const anchoColumnasReal = cabeceraEstandar.length;

  // 2. PREPARAR U OPTIMIZAR LA HOJA DE DESTINO
  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(nombreHojaDestino);
  }
  
  let datosHistoricos = hojaDestino.getDataRange().getValues();
  let historicoAConservar = [];
  
  // Identificamos dónde buscar el ID en el histórico existente
  let idxIdHist = datosHistoricos.length > 0 ? datosHistoricos[0].indexOf("ID_PROCESAMIENTO") : -1;
  if (idxIdHist === -1) idxIdHist = anchoColumnasReal - 1;

  if (datosHistoricos.length > 1) {
    for (let i = 1; i < datosHistoricos.length; i++) {
      let filaHist = datosHistoricos[i];
      let idHist = filaHist[idxIdHist]; 
      
      // Si el registro NO pertenece al lote que estamos inyectando hoy, se queda intacto
      if (!idHist || !idsLoteNuevos[idHist.toString().trim()]) {
        // Normalizamos el ancho de la fila vieja por si acaso
        while (filaHist.length < anchoColumnasReal) filaHist.push("");
        historicoAConservar.push(filaHist.slice(0, anchoColumnasReal));
      }
    }
  }

  // Aseguramos que todas las filas entrantes tengan exactamente el mismo ancho que la cabecera
  let novedadesNormalizadas = todosLosDatosEntrantes.map(fila => {
    while (fila.length < anchoColumnasReal) fila.push("");
    return fila.slice(0, anchoColumnasReal);
  });

  // 3. COMBINAR TODO EL BLOQUE EN MEMORIA RAM
  let matrizFinalLooker = [cabeceraEstandar].concat(historicoAConservar).concat(novedadesNormalizadas);
  
  // Limpieza absoluta de la solapa destino para reconfigurar el nuevo ancho sin colisiones
  hojaDestino.clear();
  
  // ESCRITURA DINÁMICA: Ajusta las columnas en base al valor real en memoria (ej. 13)
  hojaDestino.getRange(1, 1, matrizFinalLooker.length, anchoColumnasReal).setValues(matrizFinalLooker);
  
  // 4. APLICAR FORMATOS RÁPIDOS EN BLOQUE
  const totalFilas = hojaDestino.getLastRow();
  if (totalFilas > 1) {
    hojaDestino.getRange(2, 2, totalFilas - 1, 1).setNumberFormat("dd/mm/yyyy"); // Fecha en Col B
    hojaDestino.getRange(2, 4, totalFilas - 1, 5).setNumberFormat("#,##0");       // Formato numérico de corrido para importes
    hojaDestino.getRange(1, 1, 1, anchoColumnasReal).setFontWeight("bold").setBackground("#f3f3f3"); // Cabecera destacada
    hojaDestino.autoResizeColumns(1, anchoColumnasReal);
  }
  
  SpreadsheetApp.getUi().alert("🎉 ¡Sincronización Exitosa!", "Se registraron correctamente todos los datos en 'HISTORICO_DETALLE_LOOKER'.\nLotes procesados: " + Object.keys(idsLoteNuevos).join(", "), SpreadsheetApp.getUi().ButtonSet.OK);
}