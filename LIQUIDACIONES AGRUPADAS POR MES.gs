/**
 * FUNCIÓN ACTUALIZADA: Genera el resumen consolidado dentro del mismo Google Sheet.
 * Toma los datos origen únicamente de la pestaña "LIQUIDACIONES AGRUPADAS".
 * AJUSTE: Ahora mapea e importa el campo "RAMO_NOMBRE" desde la columna L de la hoja agrupada.
 */
function consolidarResumenMensual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaAgrupada = ss.getSheetByName("LIQUIDACIONES AGRUPADAS");
  
  if (!hojaAgrupada) {
    SpreadsheetApp.getUi().alert("⚠️ Error", "No se encontró la pestaña 'LIQUIDACIONES AGRUPADAS'. Primero debés correr el Paso 5.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const datos = hojaAgrupada.getDataRange().getValues();
  if (datos.length <= 1) {
    SpreadsheetApp.getUi().alert("⚠️ Atención", "La pestaña 'LIQUIDACIONES AGRUPADAS' está vacía o solo tiene encabezados.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Detectar los índices de columnas dinámicamente de la hoja agrupada
  const cabecera = datos[0];
  const idxProductor = cabecera.indexOf("PRODUCTOR_ASOCIADO") !== -1 ? cabecera.indexOf("PRODUCTOR_ASOCIADO") : cabecera.indexOf("Nombre Productor");
  const idxPrima = cabecera.indexOf("PRIMA_PESOS") !== -1 ? cabecera.indexOf("PRIMA_PESOS") : cabecera.indexOf("PRI.COB");
  const idxPremio = cabecera.indexOf("PREMIO_PESOS") !== -1 ? cabecera.indexOf("PREMIO_PESOS") : cabecera.indexOf("PAGOS");
  const idxComision = cabecera.indexOf("IMPORTE_COMISION_PESOS") !== -1 ? cabecera.indexOf("IMPORTE_COMISION_PESOS") : cabecera.indexOf("COM.DEV");
  const idxPasAgrupado = cabecera.indexOf("PAS AGRUPADO");
  
  // AJUSTE SOLICITADO: Mapeamos dinámicamente RAMO_NOMBRE (Debería ser la Col L / índice 11)
  let idxRamoNombre = cabecera.indexOf("RAMO_NOMBRE");
  if (idxRamoNombre === -1) {
    // Si por alguna razón el encabezado exacto no coincide, forzamos el índice 11 (Col L) como indicaste
    idxRamoNombre = 11; 
  }
  
  // Agrupar y sumar montos por Productor y Ramo en memoria RAM
  let resumenProductores = {};

  for (let i = 1; i < datos.length; i++) {
    let fila = datos[i];
    let productor = fila[idxProductor] ? fila[idxProductor].toString().trim() : "SIN NOMBRE";
    let ramoNombre = fila[idxRamoNombre] ? fila[idxRamoNombre].toString().trim() : "SIN RAMO";
    let prima = parseFloat(fila[idxPrima]) || 0;
    let premio = parseFloat(fila[idxPremio]) || 0;
    let comision = parseFloat(fila[idxComision]) || 0;
    let pasAgrupado = idxPasAgrupado !== -1 && fila[idxPasAgrupado] ? fila[idxPasAgrupado].toString().trim() : "";
    
    // Creamos una clave combinada única por si el mismo productor tiene múltiples ramos y querés ver el desglose limpio
    let claveUnica = productor + "_" + ramoNombre;
    
    if (!resumenProductores[claveUnica]) {
      resumenProductores[claveUnica] = { 
        productor: productor,
        ramo: ramoNombre,
        prima: 0, 
        premio: 0, 
        comision: 0, 
        pas: pasAgrupado 
      };
    }
    
    resumenProductores[claveUnica].prima += prima;
    resumenProductores[claveUnica].premio += premio;
    resumenProductores[claveUnica].comision += comision;
  }
  
  // Construir la matriz de salida con la nueva columna incorporada
  let matrizResumen = [["PRODUCTOR ASOCIADO", "PAS AGRUPADO", "RAMO NOMBRE", "TOTAL PRIMA", "TOTAL PREMIO", "TOTAL COMISIÓN"]];
  
  for (let clave in resumenProductores) {
    matrizResumen.push([
      resumenProductores[clave].productor,
      resumenProductores[clave].pas,
      resumenProductores[clave].ramo, // Campo importado de la Col L
      Number(resumenProductores[clave].prima.toFixed(2)),
      Number(resumenProductores[clave].premio.toFixed(2)),
      Number(resumenProductores[clave].comision.toFixed(2))
    ]);
  }
  
  // Destino: Buscamos o creamos la pestaña "LIQUIDACIONES AGRUPADAS POR MES"
  const nombreHojaDestino = "LIQUIDACIONES AGRUPADAS POR MES";
  let sheetDestino = ss.getSheetByName(nombreHojaDestino);
  if (!sheetDestino) {
    sheetDestino = ss.insertSheet(nombreHojaDestino);
  }
  
  // Limpiar datos anteriores y volcar la nueva matriz de 6 columnas
  sheetDestino.clear();
  sheetDestino.getRange(1, 1, matrizResumen.length, matrizResumen[0].length).setValues(matrizResumen);
  
  // Formatos estéticos aplicados al rango (ahora de 6 columnas)
  sheetDestino.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e0e0e0");
  if (matrizResumen.length > 1) {
    sheetDestino.getRange(2, 4, matrizResumen.length - 1, 3).setNumberFormat("#,##0.00");
  }
  sheetDestino.autoResizeColumns(1, 6);
  
  SpreadsheetApp.getUi().alert("🎉 Proceso Completado", "Se actualizó la pestaña '" + nombreHojaDestino + "' incorporando con éxito la columna 'RAMO NOMBRE' desde la hoja unificada.", SpreadsheetApp.getUi().ButtonSet.OK);
}