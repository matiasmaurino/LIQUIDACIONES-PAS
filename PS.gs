function consolidarYLimpiarPS() {
  const folderId = '15na20q6-kJxYyzINun1UhqAMS38Qy3Ru'; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName("PS");
  if (!sheet) {
    sheet = ss.insertSheet("PS");
    sheet.appendRow(["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO"]);
  }

  try {
    const folder = DriveApp.getFolderById(folderId.trim());
    const files = folder.getFilesByType(MimeType.CSV);
    
    const columnasInteres = [17, 6, 5, 1, 2, -1, 12, 14, 4]; 
    let todosLosDatosNuevos = [];

    while (files.hasNext()) {
      let file = files.next();
      let csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'), ';');
      
      let datosProcesados = csvData.slice(4).map((fila) => {
        if (!fila[17] || fila[17].toString().trim() === "" || fila[17].toString().includes("PRODUCTOR")) return null;

        let filaNueva = [];
        columnasInteres.forEach(idx => {
          if (idx === -1) {
            filaNueva.push(0); 
          } else {
            let valor = fila[idx];
            if (idx === 6) { // Procesar Fecha de PS
              valor = traducirMesYAnioPS(valor);
            }
            filaNueva.push(valor);
          }
        });

        let tieneDato = filaNueva[0] !== "" && filaNueva[0] != null;
        filaNueva.push(tieneDato ? "PS" : "");

        let nombreProductor = filaNueva[0] ? filaNueva[0].toString().trim().toUpperCase() : "";
        if (nombreProductor === "MAURIÑO MATIAS") {
          filaNueva.push("MATIAS");
        } else {
          filaNueva.push(tieneDato ? "DGM" : "");
        }
        return filaNueva;
      }).filter(fila => fila !== null);

      if (datosProcesados.length > 0) {
        todosLosDatosNuevos = todosLosDatosNuevos.concat(datosProcesados);
      }
    }

    if (todosLosDatosNuevos.length > 0) {
      // Identificar meses entrantes para limpiar preventivamente el histórico
      let mesesAQuitar = {};
      todosLosDatosNuevos.forEach(fila => {
        let f = fila[1]; // Índice 1 es FECHA
        let mAnio = "";
        if (f instanceof Date) {
          mAnio = f.getFullYear() + "-" + (f.getMonth() + 1).toString().padStart(2, '0');
        } else {
          // Si viene mapeado como dd/mm/yyyy del traductor
          let partes = f.toString().split("/");
          if (partes.length === 3) mAnio = partes[2] + "-" + partes[1].padStart(2, '0');
        }
        if (mAnio.length === 7) mesesAQuitar[mAnio] = true;
      });

      // Limpiar del histórico existente sólo las filas que coincidan con los meses nuevos
      let datosHistoricos = sheet.getDataRange().getValues();
      for (let i = datosHistoricos.length - 1; i >= 1; i--) {
        let fHist = datosHistoricos[i][1];
        let mHist = "";
        if (fHist instanceof Date) {
          mHist = fHist.getFullYear() + "-" + (fHist.getMonth() + 1).toString().padStart(2, '0');
        } else {
          let partes = fHist.toString().split("/");
          if (partes.length === 3) mHist = partes[2] + "-" + partes[1].padStart(2, '0');
        }
        if (mesesAQuitar[mHist]) {
          sheet.deleteRow(i + 1);
        }
      }

      // Append acumulado de lo nuevo
      let ultimaFilaDestino = sheet.getLastRow();
      sheet.getRange(ultimaFilaDestino + 1, 1, todosLosDatosNuevos.length, todosLosDatosNuevos[0].length).setValues(todosLosDatosNuevos);
    }

    const ultimaFilaPS = sheet.getLastRow();
    if (ultimaFilaPS > 1) {
      sheet.getRange(2, 2, ultimaFilaPS - 1).setNumberFormat("dd/mm/yyyy"); 
      sheet.getRange(2, 7, ultimaFilaPS - 1, 2).setNumberFormat("#,##0"); 
    }
    console.log("✅ PS procesado en modo incremental de manera exitosa.");
  } catch (e) {
    console.error("Error en PS: " + e.message);
  }
}