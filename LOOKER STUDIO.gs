/**
 * REVOLUCIONARIA FUNCIÓN MULTI-COMPAÑÍA: Sincroniza el detalle analítico hacia la hoja de Looker.
 * Escanea de forma inteligente las solapas de las compañías actuales y purga el gran histórico
 * basándose estrictamente en los ID_PROCESAMIENTO individuales (ej. 2026-05_RIV, 2026-05_PS).
 */
function actualizarHistoricoDetalleLooker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombresHojasOrigen = ["FP", "RIV", "PS"];
  const nombreHojaDestino = "HISTORICO_DETALLE_LOOKER";
  
  let todosLosDatosEntrantes = [];
  let idsLoteNuevos = {};
  let cabeceraEstandar = ["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"];

  // 1. ESCANEAR Y RECOLECTAR NOVEDADES DE LAS PESTAÑAS INDIVIDUALES
  nombresHojasOrigen.forEach(nombre => {
    let hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    
    let datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return; // Si solo tiene cabecera, pasamos de largo

    // Forzamos a que use la primera cabecera válida encontrada como molde estructural
    cabeceraEstandar = datos[0];

    for (let i = 1; i < datos.length; i++) {
      let idFila = datos[i][11]; // Columna L (ID_PROCESAMIENTO)
      if (idFila) {
        idsLoteNuevos[idFila] = true;
        todosLosDatosEntrantes.push(datos[i]);
      }
    }
  });

  if (todosLosDatosEntrantes.length === 0) {
    console.log("No se detectaron novedades transaccionales en las pestañas de las compañías.");
    return;
  }

  // 2. LEER GRAN HISTÓRICO DE LOOKER Y FILTRAR QUIRÚRGICAMENTE EN MEMORIA
  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(nombreHojaDestino);
    hojaDestino.appendRow(cabeceraEstandar);
  }
  
  let datosHistoricos = hojaDestino.getDataRange().getValues();
  let historicoAConservar = [];

  if (datosHistoricos.length > 1) {
    cabeceraEstandar = datosHistoricos[0];
    for (let i = 1; i < datosHistoricos.length; i++) {
      let idHist = datosHistoricos[i][11]; // Columna L en el histórico total
      
      // Si el registro NO pertenece al lote específico de la compañía/mes que estamos inyectando, se queda
      if (!idsLoteNuevos[idHist]) {
        historicoAConservar.push(datosHistoricos[i]);
      }
    }
  }

  // 3. AMALGAMAR Y REESCRIBIR DE UN SOLO TIRO
  let matrizFinalLooker = [cabeceraEstandar].concat(historicoAConservar).concat(todosLosDatosEntrantes);
  
  hojaDestino.clearContents();
  hojaDestino.getRange(1, 1, matrizFinalLooker.length, matrizFinalLooker[0].length).setValues(matrizFinalLooker);
  
  // Formatos en bloque súper veloces
  const totalFilas = hojaDestino.getLastRow();
  if (totalFilas > 1) {
    hojaDestino.getRange(2, 2, totalFilas - 1).setNumberFormat("dd/mm/yyyy");
    hojaDestino.getRange(2, 6, totalFilas - 1, 3).setNumberFormat("#,##0"); // Formatea Prima, Premio y Comisión de corrido
    hojaDestino.autoResizeColumns(1, 12);
  }
  console.log("🚀 Sincronización a Looker completada con éxito. Procesados lotes: " + Object.keys(idsLoteNuevos).join(", "));
}