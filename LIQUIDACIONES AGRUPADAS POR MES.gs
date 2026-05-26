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

    let ultimaFila = sheet.getLastRow();
    if (ultimaFila <= 1) return; 
    
    // AMPLIACIÓN CRÍTICA: Leemos 13 columnas (de la A a la M) para poder capturar la columna M (índice 12)
    let dataRaw = sheet.getRange(1, 1, ultimaFila, 13).getValues();

    // MAPEO DE COLUMNAS CORREGIDO
    let colFecha = 1;     // Columna B
    let colRamo = 12;     // SOLUCIÓN: Columna M (índice 12 en base 0) trae la DESCRIPCIÓN del Ramo
    let colPrima = 4;     // Columna E
    let colPremio = 5;    // Columna F
    let colComision = 6;  // Columna G
    let colCia = 9;       // Columna J
    let colPAS = 10;      // Columna K

    for (let i = 1; i < dataRaw.length; i++) {
      let fila = dataRaw[i];
      let pas = fila[colPAS] ? fila[colPAS].toString().trim().toUpperCase() : "";
      let cia = fila[colCia] ? fila[colCia].toString().trim().toUpperCase() : nombreHoja;
      let ramo = fila[colRamo] ? fila[colRamo].toString().trim() : "";
      let fechaRaw = fila[colFecha];
      
      if (!pas || !fechaRaw) continue;

      // Si el ramo por algún motivo viene vacío de la fórmula auxiliar, usamos un genérico para que no se pierda
      if (ramo === "") ramo = "SIN ESPECIFICAR";

      let mesAnioStr = "";
      if (fechaRaw instanceof Date) {
        let mm = (fechaRaw.getMonth() + 1).toString().padStart(2, '0');
        mesAnioStr = fechaRaw.getFullYear() + "-" + mm;
      } else if (fechaRaw.toString().includes("/")) {
        let partes = fechaRaw.toString().split("/");
        if (partes.length === 3) mesAnioStr = partes[2].trim() + "-" + partes[1].trim().padStart(2, '0');
      }

      if (mesAnioStr === "" || mesAnioStr.length !== 7) continue;
      
      let claveLote = mesAnioStr + "_" + cia;
      lotesNuevosMap[claveLote] = true;

      // Agrupamos usando el nombre del ramo (Columna M)
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

  // 2. LEER EL HISTÓRICO EXISTENTE EN LA HOJA DE DESTINO SIN TOCAR FILAS DE TOTALES
  let datosDestinoRaw = hojaDestino.getDataRange().getValues();
  let encabezadosDestino = ["Año-mes", "PAS", "RAMO_NOMBRE", "CIA", "PRIMA", "PREMIO", "COMISION", "CANT POLIZAS"];
  let historicoAConservar = [];

  if (datosDestinoRaw.length > 1) {
    encabezadosDestino = datosDestinoRaw[0];
    for (let i = 1; i < datosDestinoRaw.length; i++) {
      let filaHist = datosDestinoRaw[i];
      let mesAnioEx = filaHist[0] ? filaHist[0].toString().trim() : "";
      let ciaEx = filaHist[3] ? filaHist[3].toString().trim().toUpperCase() : "";
      
      // Evitamos procesar filas de totales o vacías en la matriz dinámica
      if (mesAnioEx.toLowerCase().includes("total") || mesAnioEx === "") {
        continue;
      }

      let claveFilaHist = mesAnioEx + "_" + ciaEx;
      if (!lotesNuevosMap[claveFilaHist]) {
        historicoAConservar.push(filaHist.slice(0, 8));
      }
    }
  }

  // 3. PASAR LAS NOVEDADES PROCESADAS A MATRIZ
  let matrizNovedades = [];
  Object.keys(resumenesNuevosMap).forEach(clave => {
    let item = resumenesNuevosMap[clave];
    matrizNovedades.push([
      item.mesAnio, item.pas, item.ramo, item.cia,
      item.prima, item.premio, item.comision, item.polizas
    ]);
  });

  // Combinamos el contenido estructurado
  let matrizFinalCompleta = [encabezadosDestino.slice(0, 8)].concat(historicoAConservar).concat(matrizNovedades);
  
  // 4. LIMPIEZA Y ESCRITURA EN BLOQUE DE LAS COLUMNAS A-H
  let ultimaFilaDestinoOriginal = hojaDestino.getLastRow();
  if (ultimaFilaDestinoOriginal > 0) {
    hojaDestino.getRange(1, 1, ultimaFilaDestinoOriginal, 8).clearContent();
  }
  
  // Volcamos los datos correctos
  hojaDestino.getRange(1, 1, matrizFinalCompleta.length, 8).setValues(matrizFinalCompleta);
  
  // 5. REAPLICAR FORMATOS AUTOMÁTICOS
  const ultimaFilaFinal = hojaDestino.getLastRow();
  if (ultimaFilaFinal > 1) {
    hojaDestino.getRange(2, 1, ultimaFilaFinal - 1, 4).setNumberFormat("@"); 
    hojaDestino.getRange(2, 5, ultimaFilaFinal - 1, 3).setNumberFormat("$#,##0"); 
    hojaDestino.getRange(2, 8, ultimaFilaFinal - 1, 1).setNumberFormat("#,##0"); 
    hojaDestino.autoResizeColumns(1, 8);
  }

  console.log("✅ ¡Resumen Mensual consolidado con descripciones de ramo correctas!");
}