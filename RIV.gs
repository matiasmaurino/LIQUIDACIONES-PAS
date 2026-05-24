function consolidarYLimpiarRIV() {
  const folderId = '1cL4zPoAD7YJjhhHMZvg33vBo0kB5OsHZ'; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName("RIV");
  if (!sheet) {
    sheet = ss.insertSheet("RIV");
    sheet.appendRow(["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"]);
  }
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const columnasInteres = [0, 3, 4, 5, 10, 8, 11, 2];
    let todosLosDatosNuevos = [];
    let idsProcesamientoNuevos = {};

    while (files.hasNext()) {
      let file = files.next();
      let fileName = file.getName().toLowerCase();
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        let blob = file.getBlob();
        let resource = {
          title: "[TEMP] " + file.getName(),
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [{id: folderId}]
        };
        
        let tempFile = Drive.Files.create(resource, blob);
        let tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
        let tempSheet = tempSpreadsheet.getSheets()[0];
        let fullData = tempSheet.getDataRange().getValues();
        
        DriveApp.getFileById(tempFile.id).setTrashed(true);
        if (!fullData || fullData.length <= 1) continue;

        let datosProcesados = fullData.slice(1).map((fila) => {
          if (fila.length === 0) return null;

          let fechaFormateada = "";
          let mesAnioStr = "INDEFINIDO";

          let columnasExtraidas = columnasInteres.map(idx => {
            let valor = fila[idx];
            if (valor === undefined || valor === null) return "";
            
            if (columnasInteres.indexOf(idx) === 0) {
              let dateObj = null;
              if (valor instanceof Date) {
                dateObj = valor;
              } else if (typeof valor === 'number') {
                dateObj = new Date(1899, 11, 30);
                dateObj.setDate(dateObj.getDate() + valor);
              }
              if (dateObj) {
                fechaFormateada = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
                mesAnioStr = dateObj.getFullYear() + "-" + (dateObj.getMonth() + 1).toString().padStart(2, '0');
                return fechaFormateada;
              }
            }
            return valor.toString().trim();
          });

          let tieneDato = fila[0] !== "" && fila[0] != null; 
          let nombreProductor = tieneDato ? "MAURIÑO DANIEL GUILLERMO" : "";
          let compania = tieneDato ? "RIV" : "";
          let pasAgrupado = tieneDato ? "DGM" : "";
          
          // Generar ID de lote específico para Rivadavia
          let idProcesamiento = mesAnioStr + "_RIV";
          if (tieneDato) idsProcesamientoNuevos[idProcesamiento] = true;

          return [nombreProductor].concat(columnasExtraidas).concat([compania, pasAgrupado, idProcesamiento]);
        }).filter(fila => fila !== null);

        todosLosDatosNuevos = todosLosDatosNuevos.concat(datosProcesados);
      }
    }

    // FILTRADO SEGURO EN MEMORIA (SÚPER VELOZ)
    if (todosLosDatosNuevos.length > 0) {
      let datosHistoricos = sheet.getDataRange().getValues();
      let encabezados = ["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"];
      let historicoConservado = [];

      if (datosHistoricos.length > 1) {
        encabezados = datosHistoricos[0];
        // Columna L es índice 11
        for (let i = 1; i < datosHistoricos.length; i++) {
          let idHist = datosHistoricos[i][11]; 
          if (!idsProcesamientoNuevos[idHist]) {
            historicoConservado.push(datosHistoricos[i]);
          }
        }
      }

      let resultadoFinal = [encabezados].concat(historicoConservado).concat(todosLosDatosNuevos);

      sheet.clearContents();
      sheet.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);

      const ultimaFila = sheet.getLastRow();
      if (ultimaFila > 1) {
        sheet.getRange(2, 2, ultimaFila - 1, 1).setNumberFormat("dd/mm/yyyy");
        sheet.getRange(2, 6, ultimaFila - 1, 3).setNumberFormat("#,##0"); 
        sheet.getRange(2, 9, ultimaFila - 1, 1).setNumberFormat("@"); 
        sheet.autoResizeColumns(1, 12);
      }
      console.log("✅ RIV procesado de forma incremental mediante ID de procesamiento.");
    }
  } catch (e) {
    console.log("Error en RIV: " + e.message);
  }
}