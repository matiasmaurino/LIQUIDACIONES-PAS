function consolidarYLimpiarPS() {
  const folderId = '15na20q6-kJxYyzINun1UhqAMS38Qy3Ru'; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName("PS");
  if (!sheet) {
    sheet = ss.insertSheet("PS");
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
              valor = traducirMesYAnioPS(valor); 
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

        let idProcesamiento = mesAnioStr + "_PS";
        if (tieneDato) idsProcesamientoNuevos[idProcesamiento] = true;
        
        filaNueva.push(idProcesamiento);
        return filaNueva;
      }).filter(fila => fila !== null);
      if (datosProcesados.length > 0) {
        todosLosDatosNuevos = todosLosDatosNuevos.concat(datosProcesados);
      }
    }

    if (todosLosDatosNuevos.length > 0) {
      let datosHistoricos = sheet.getDataRange().getValues();
      let encabezados = ["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"];
      let historicoConservado = [];
      
      if (datosHistoricos.length > 1) {
        encabezados = datosHistoricos[0];
        for (let i = 1; i < datosHistoricos.length; i++) {
          let idHist = datosHistoricos[i][11];
          if (idHist && !idsProcesamientoNuevos[idHist]) {
            historicoConservado.push(datosHistoricos[i].slice(0, 12));
          }
        }
      }

      let resultadoFinal = [encabezados.slice(0, 12)].concat(historicoConservado).concat(todosLosDatosNuevos);
      
      // 1. LIMPIEZA TOTAL
      sheet.clearContents();
      
      // 2. ESCRITURA DE DATOS
      sheet.getRange(1, 1, resultadoFinal.length, 12).setValues(resultadoFinal);

      // 3. INYECCIÓN DE LA FÓRMULA ARRAY EN M1 Y M2
      sheet.getRange("M1").setValue("AUXILIAR_BUSQUEDA");
      sheet.getRange("M2").setFormula(`=ARRAYFORMULA(IF(A2:A<>"";XLOOKUP(D2:D;AUX!A2:A;AUX!d2:d;"");""))`);

      const ultimaFilaPS = sheet.getLastRow();
      if (ultimaFilaPS > 1) {
        sheet.getRange(2, 2, ultimaFilaPS - 1).setNumberFormat("dd/mm/yyyy");
        sheet.getRange(2, 7, ultimaFilaPS - 1, 2).setNumberFormat("#,##0"); 
        sheet.autoResizeColumns(1, 13);
      }
      console.log("✅ PS procesado e inyectada su ARRAYFORMULA en M2.");
    }
  } catch (e) {
    console.error("Error en PS: " + e.message);
  }
}

function traducirMesYAnioPS(valorTexto) {
  if (!valorTexto) return "";
  var str = valorTexto.toString().trim().toUpperCase();
  var partes = str.split("-");
  if (partes.length !== 2) return valorTexto;
  
  var mesTexto = partes[0];
  var anioTexto = partes[1];
  
  var meses = {
    "ENE": "01", "FEB": "02", "MAR": "03", "ABR": "04", 
    "MAY": "05", "JUN": "06", "JUL": "07", "AGO": "08", 
    "SEP": "09", "OCT": "10", "NOV": "11", "DIC": "12"
  };
  
  var mesNum = meses[mesTexto];
  if (!mesNum) return valorTexto;
  
  var anioCompleto = anioTexto.length === 2 ? "20" + anioTexto : anioTexto;
  return "01/" + mesNum + "/" + anioCompleto;
}