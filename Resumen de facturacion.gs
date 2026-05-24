/**
 * VERSÍON ULTRA-OPTIMIZADA EN MEMORIA (CORREGIDA - 8 COLUMNAS): 
 * Agrupa los datos detallados de las pestañas individuales de las compañías (FP, RIV, PS) 
 * y actualiza de forma incremental la hoja "LIQUIDACIONES AGRUPADAS POR MES".
 * Respeta estrictamente la estructura histórica de 8 columnas.
 */
function consolidarResumenMensual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasOrigen = ["FP", "RIV", "PS"];
  
  const hojaDestino = ss.getSheetByName("LIQUIDACIONES AGRUPADAS POR MES");
  if (!hojaDestino) {
    throw new Error("No se encontró la hoja de destino 'LIQUIDACIONES AGRUPADAS POR MES'");
  }

  // 1. LEER TODO EL RESUMEN EXISTENTE A LA MEMORIA (ARRAY)
  let datosDestinoRaw = hojaDestino.getDataRange().getValues();
  // Estructura nativa exacta de 8 columnas de tu histórico
  let encabezadosDestino = ["Año-mes", "PAS", "RAMO_NOMBRE", "CIA", "PRIMA", "PREMIO", "COMISION", "CANT POLIZAS"];
  let resumenesMap = {}; 

  if (datosDestinoRaw.length > 1) {
    encabezadosDestino = datosDestinoRaw[0];
    for (let i = 1; i < datosDestinoRaw.length; i++) {
      let mesAnioEx = datosDestinoRaw[i][0] ? datosDestinoRaw[i][0].toString().trim() : "";
      let pasEx = datosDestinoRaw[i][1] ? datosDestinoRaw[i][1].toString().trim().toUpperCase() : "";
      let ramoEx = datosDestinoRaw[i][2] ? datosDestinoRaw[i][2].toString().trim() : "";
      let ciaEx = datosDestinoRaw[i][3] ? datosDestinoRaw[i][3].toString().trim().toUpperCase() : "";
      
      // Clave única combinando las 4 dimensiones para evitar mezclar datos
      let clave = mesAnioEx + "_" + pasEx + "_" + ciaEx + "_" + ramoEx;
      
      resumenesMap[clave] = {
        mesAnio: mesAnioEx,
        pas: pasEx,
        ramo: ramoEx,
        cia: ciaEx,
        prima: parseFloat(datosDestinoRaw[i][4]) || 0,
        premio: parseFloat(datosDestinoRaw[i][5]) || 0,
        comision: parseFloat(datosDestinoRaw[i][6]) || 0,
        polizas: parseInt(datosDestinoRaw[i][7]) || 0
      };
    }
  }

  // 2. PROCESAR CADA COMPAÑÍA COMPLETAMENTE EN MEMORIA
  hojasOrigen.forEach(nombreHoja => {
    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;

    let dataRaw = sheet.getDataRange().getValues();
    if (dataRaw.length <= 1) return; 

    // Índices de columnas estables según tus archivos RIV, PS y FP optimizados
    let colFecha = 1;      // Columna B (FECHA)
    let colRamo = 3;       // Columna D (RAMO)
    let colPrima = nombreHoja === "FP" ? 3 : 5;     // D para FP, F para RIV/PS
    let colPremio = nombreHoja === "FP" ? 4 : 6;    // E para FP, G para RIV/PS
    let colComision = nombreHoja === "FP" ? 5 : 7;  // F para FP, H para RIV/PS
    let colCia = 9;        // Columna J (CIA)
    let colPAS = 10;       // Columna K (PAS AGRUPADO)

    for (let i = 1; i < dataRaw.length; i++) {
      let fila = dataRaw[i];
      let pas = fila[colPAS] ? fila[colPAS].toString().trim().toUpperCase() : "";
      let cia = fila[colCia] ? fila[colCia].toString().trim().toUpperCase() : "";
      let ramo = fila[colRamo] ? fila[colRamo].toString().trim() : "";
      let fechaRaw = fila[colFecha];
      
      if (!pas || !fechaRaw) continue;

      let mesAnioStr = "";
      // Formatear fecha de forma segura a YYYY-MM
      if (fechaRaw instanceof Date) {
        let mm = (fechaRaw.getMonth() + 1).toString().padStart(2, '0');
        mesAnioStr = fechaRaw.getFullYear() + "-" + mm;
      } else if (fechaRaw.toString().includes("/")) {
        let partes = fechaRaw.toString().split("/");
        if (partes.length === 3) {
          mesAnioStr = partes[2].trim() + "-" + partes[1].trim().padStart(2, '0');
        }
      }

      if (mesAnioStr === "" || mesAnioStr.length !== 7) continue;

      let claveAgrupado = mesAnioStr + "_" + pas + "_" + cia + "_" + ramo;
      let prima = parseFloat(fila[colPrima]) || 0;
      let premio = parseFloat(fila[colPremio]) || 0;
      let comision = parseFloat(fila[colComision]) || 0;

      // Acumulación matemática en memoria RAM
      if (resumenesMap[claveAgrupado]) {
        resumenesMap[claveAgrupado].prima += prima;
        resumenesMap[claveAgrupado].premio += premio;
        resumenesMap[claveAgrupado].comision += comision;
        resumenesMap[claveAgrupado].polizas += 1;
      } else {
        resumenesMap[claveAgrupado] = {
          mesAnio: mesAnioStr,
          pas: pas,
          ramo: ramo,
          cia: cia,
          prima: prima,
          premio: premio,
          comision: comision,
          polizas: 1
        };
      }
    }
  });

  // 3. RECONSTRUIR LA MATRIZ FINAL (8 COLUMNAS EXACTAS)
  let resultadoMatrizFinal = [encabezadosDestino];
  
  Object.keys(resumenesMap).forEach(clave => {
    let item = resumenesMap[clave];
    resultadoMatrizFinal.push([
      item.mesAnio, // Col A
      item.pas,     // Col B
      item.ramo,    // Col C
      item.cia,     // Col D
      item.prima,   // Col E
      item.premio,  // Col F
      item.comision,// Col G
      item.polizas  // Col H
    ]);
  });

  // 4. ESCRIBUIR EL BLOQUE COMPLETO DE UN SOLO GOLPE
  hojaDestino.clearContents();
  hojaDestino.getRange(1, 1, resultadoMatrizFinal.length, resultadoMatrizFinal[0].length).setValues(resultadoMatrizFinal);

  // Formatos rápidos en lote
  const ultimaFilaDestino = hojaDestino.getLastRow();
  if (ultimaFilaDestino > 1) {
    hojaDestino.getRange(2, 1, ultimaFilaDestino - 1, 4).setNumberFormat("@"); // Columnas de texto congeladas
    hojaDestino.getRange(2, 5, ultimaFilaDestino - 1, 3).setNumberFormat("$#,##0"); // Formato moneda para valores económicos
    hojaDestino.getRange(2, 8, ultimaFilaDestino - 1, 1).setNumberFormat("#,##0"); // Cantidad de pólizas entero
    hojaDestino.autoResizeColumns(1, 8);
  }

  console.log("✅ ¡Éxito absoluto! Resumen Mensual consolidado respetando las 8 columnas históricas.");
}