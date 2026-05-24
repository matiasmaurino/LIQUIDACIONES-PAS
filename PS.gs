function consolidarYLimpiarPS() {
  const folderId = '15na20q6-kJxYyzINun1UhqAMS38Qy3Ru'; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName("PS");
  if (!sheet) {
    sheet = ss.insertSheet("PS");
    sheet.appendRow(["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"]);
  }

  try {
    const folder = DriveApp.getFolderById(folderId.trim());
    const files = folder.getFilesByType(MimeType.CSV);
    const columnasInteres = [17, 6, 5, 1, 2, -1, 12, 14, 4]; 
    let todosLosDatosNuevos = [];
    let idsProcesamientoNuevos = {};

    while (files.hasNext()) {
      let file = files.next();
      let csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'), ';');
      
      let datosProcesados = csvData.slice(4).map((fila) => {
        if (!fila[17] || fila[17].toString().trim() === "" || fila[17].toString().includes("PRODUCTOR")) return null;

        let filaNueva = [];
        let mesAnioStr = "INDEFINIDO";

        columnasInteres.forEach(idx => {
          if (idx === -1) {
            filaNueva.push(0); 
          } else {
            let valor = fila[idx];
            if (idx === 6) { 
              // Obtenemos el formato estandarizado y calculamos el mesAnioStr
              valor = traducirMesYAnioPS(valor); // Mantiene tu lógica nativa de traducción
              if (valor && valor.includes("/")) {
                let partes = valor.split("/");
                if (partes.length === 3) mesAnioStr = partes[2] + "-" + partes[1].padStart(2, '0');
              }
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

        // Generar ID de lote específico para PS
        let idProcesamiento = mesAnioStr + "_PS";
        if (tieneDato) idsProcesamientoNuevos[idProcesamiento] = true;
        
        filaNueva.push(idProcesamiento);
        return filaNueva;
      }).filter(fila => fila !== null);

      if (datosProcesados.length > 0) {
        todosLosDatosNuevos = todosLosDatosNuevos.concat(datosProcesados);
      }
    }

    // FILTRADO QUIRÚRGICO EN MEMORIA (REEMPLAZA AL DELETE_ROW LENTO)
    if (todosLosDatosNuevos.length > 0) {
      let datosHistoricos = sheet.getDataRange().getValues();
      let encabezados = ["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"];
      let historicoConservado = [];

      if (datosHistoricos.length > 1) {
        encabezados = datosHistoricos[0];
        for (let i = 1; i < datosHistoricos.length; i++) {
          let idHist = datosHistoricos[i][11]; // Columna L
          if (!idsProcesamientoNuevos[idHist]) {
            historicoConservado.push(datosHistoricos[i]);
          }
        }
      }

      let resultadoFinal = [encabezados].concat(historicoConservado).concat(todosLosDatosNuevos);

      sheet.clearContents();
      sheet.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);

      const ultimaFilaPS = sheet.getLastRow();
      if (ultimaFilaPS > 1) {
        sheet.getRange(2, 2, ultimaFilaPS - 1).setNumberFormat("dd/mm/yyyy");
        sheet.getRange(2, 7, ultimaFilaPS - 1, 2).setNumberFormat("#,##0"); 
        sheet.autoResizeColumns(1, 12);
      }
      console.log("✅ PS procesado de forma incremental mediante ID de procesamiento.");
    }
  } catch (e) {
    console.error("Error en PS: " + e.message);
  }
}