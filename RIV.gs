function consolidarResumenMensual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasOrigen = ["FP", "RIV", "PS"];
  const hojaDestino = ss.getSheetByName("LIQUIDACIONES AGRUPADAS POR MES");
  if (!hojaDestino) {
    throw new Error("No se encontró la hoja de destino 'LIQUIDACIONES AGRUPADAS POR MES'");
  }

  // 1. DETECTAR QUÉ MESES Y COMPAÑÍAS ESTAMOS PROCESANDO EN ESTA TANDA
  let lotesNuevosMap = {};
  let resumenesNuevosMap = {};

  hojasOrigen.forEach(nombreHoja => {
    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;

    // LECTURA SEGURA: Leemos explícitamente hasta la columna L (12 columnas) para evitar desfases con la columna M
    let ultimaFila = sheet.getLastRow();
    if (ultimaFila <= 1) return; 
    let dataRaw = sheet.getRange(1, 1, ultimaFila, 12).getValues();

    // MAPEO UNIFICADO Y REAL DE COLUMNAS (Índices basados en estructura de la A a la L)
    let colFecha = 1;     // Columna B
    let colRamo = 3;      // Columna D
    let colPrima = 5;     // Columna F (Aplica igual para FP, RIV y PS en la estructura limpia)
    let colPremio = 6;    // Columna G
    let colComision = 7;  // Columna H
    let colCia = 9;       // Columna J
    let colPAS = 10;      // Columna K

    for (let i = 1; i < dataRaw.length; i++) {
      let fila = dataRaw[i];
      let pas = fila[colPAS] ? fila[colPAS].toString().trim().toUpperCase() : "";
      let cia = fila[colCia] ? fila[colCia].toString().trim().toUpperCase() : nombreHoja;
      let ramo = fila[colRamo] ? fila[colRamo].toString().trim() : "";
      let fechaRaw = fila[colFecha];
      
      if (!pas || !fechaRaw) continue;

      let mesAnioStr = "";
      if (fechaRaw instanceof Date) {
        let mm = (fechaRaw.getMonth() + 1).toString().padStart(2, '0');
        mesAnioStr = fechaRaw.getFullYear() + "-" + mm;
      } else if (fechaRaw.toString().includes("/")) {
        let partes = fechaRaw.toString().split("/");
        if (partes.length === 3) mesAnioStr = partes[2].trim() + "-" + partes[1].trim().padStart(2, '0');
      }

      if (mesAnioStr === "" || mesAnioStr.length !== 7) continue;
      
      // Marcamos este lote (Mes + Compañía) como "activo" para purgarlo del histórico y no duplicar
      let claveLote = mesAnioStr + "_" + cia;
      lotesNuevosMap[claveLote] = true;

      // Acumulamos la agrupación en memoria RAM
      let claveAgrupado = mesAnioStr + "_" + pas + "_" + cia + "_" + ramo;
      if (resumenesNuevosMap[claveAgrupado]) {
        resumenesNuevosMap[claveAgrupado].prima += parseFloat(fila[colPrima]) || 0;
        resumenesNuevosMap[claveAgrupado].premio += parseFloat(fila[colPremio]) || 0;
        resumenesNuevosMap[claveAgrupado].comision += parseFloat(fila[colComision]) || 0;
        resumenesNuevosMap[claveAgrupado].polizas += 1;
      } else {
        resumenesNuevosMap[claveAgrupado] = {
          mesAnio: mesAnioStr, pas: pas, ramo: ramo, cia: cia,
          prima: parseFloat(fila[colPrima]) || 0,
          premio: parseFloat(fila[colPremio]) || 0,
          comision: parseFloat(fila[colComision]) || 0,
          polizas: 1
        };
      }
    }
  });

  // 2. LEER EL HISTÓRICO EXISTENTE Y FILTRAR CON CUIDADO (CONSERVANDO LO MANUAL)
  let datosDestinoRaw = hojaDestino.getDataRange().getValues();
  let encabezadosDestino = ["Año-mes", "PAS", "RAMO_NOMBRE", "CIA", "PRIMA", "PREMIO", "COMISION", "CANT POLIZAS"];
  let historicoAConservar = [];
  if (datosDestinoRaw.length > 1) {
    encabezadosDestino = datosDestinoRaw[0];
    for (let i = 1; i < datosDestinoRaw.length; i++) {
      let filaHist = datosDestinoRaw[i];
      let mesAnioEx = filaHist[0] ? filaHist[0].toString().trim() : "";
      let ciaEx = filaHist[3] ? filaHist[3].toString().trim().toUpperCase() : "";
      let claveFilaHist = mesAnioEx + "_" + ciaEx;

      // Protege cargas manuales de meses anteriores o registros históricos puros.
      if (!lotesNuevosMap[claveFilaHist]) {
        historicoAConservar.push(filaHist);
      }
    }
  }

  // 3. TRANSFORMAR LAS NOVEDADES A ARRAY
  let matrizNovedades = [];
  Object.keys(resumenesNuevosMap).forEach(clave => {
    let item = resumenesNuevosMap[clave];
    matrizNovedades.push([
      item.mesAnio, item.pas, item.ramo, item.cia,
      item.prima, item.premio, item.comision, item.polizas
    ]);
  });

  // 4. UNIFICAR Y ESCRIBIR TODO EL BLOQUE PROTEGIDO
  let matrizFinalCompleta = [encabezadosDestino].concat(historicoAConservar).concat(matrizNovedades);
  
  // Limpiamos y reescribimos el consolidado final protegido
  hojaDestino.clear();
  hojaDestino.getRange(1, 1, matrizFinalCompleta.length, 8).setValues(matrizFinalCompleta);
  
  // Formatos rápidos en lote
  const ultimaFilaDestino = hojaDestino.getLastRow();
  if (ultimaFilaDestino > 1) {
    hojaDestino.getRange(2, 1, ultimaFilaDestino - 1, 4).setNumberFormat("@"); 
    hojaDestino.getRange(2, 5, ultimaFilaDestino - 1, 3).setNumberFormat("$#,##0");
    hojaDestino.getRange(2, 8, ultimaFilaDestino - 1, 1).setNumberFormat("#,##0"); 
    hojaDestino.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f3f3f3");
    hojaDestino.autoResizeColumns(1, 8);
  }

  console.log("✅ ¡Resumen Mensual actualizado correctamente con montos reales!");
}