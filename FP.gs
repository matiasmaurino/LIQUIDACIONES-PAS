function consolidarYLimpiarFP() {
  const folderId = '1MFWeyrluXJdDA8pJyuzAGAeRRAOeIHbV';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaClientes = ss.getSheetByName("CLIENTES FP");
  const clientesMap = {};
  if (hojaClientes) {
    const dataClientes = hojaClientes.getDataRange().getValues();
    for (let i = 1; i < dataClientes.length; i++) {
      let nombreCliente = dataClientes[i][2] ? dataClientes[i][2].toString().trim().toUpperCase() : ""; 
      let idFedPatronal = dataClientes[i][0]; 
      clientesMap[nombreCliente] = idFedPatronal;
    }
  }

  let nombreHoja = "FP";
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) {
    sheet = ss.insertSheet(nombreHoja);
  }
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.CSV);
  const columnasInteres = [2, 3, 6, 7, 8, 16, 17, 18]; 
  
  let todosLosDatosNuevos = [];
  while (files.hasNext()) {
    let file = files.next();
    let csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'));
    for (let i = 1; i < csvData.length; i++) {
      let fila = csvData[i];
      if (!fila[2] || fila[2].toString().trim() === "" || fila[2].toString().includes("PRODUCTOR")) continue;
      
      let filaNueva = columnasInteres.map((idx, colPos) => {
        let valor = fila[idx] ? fila[idx].toString().trim() : "";
        
        if (colPos === 0 && valor.toUpperCase().includes("MAURI")) {
          return "MAURIÑO MATIAS";
        }

        if (valor.includes("Ã‘")) valor = valor.replace(/Ã‘/g, "Ñ");
        if (valor.includes("Ã³")) valor = valor.replace(/Ã³/g, "ó");

        if (colPos === 1 && valor.includes("-")) {
          let partes = valor.substring(0, 10).split('-');
          if (partes.length === 3) return partes[2] + "/" + partes[1] + "/" + partes[0];
        }
        return valor;
      });

      let fechaFila = filaNueva[1];
      let mesAnioStr = "INDEFINIDO";
      if (fechaFila && fechaFila.includes("/")) {
        let partesF = fechaFila.split("/");
        if (partesF.length === 3) mesAnioStr = partesF[2].trim() + "-" + partesF[1].trim().padStart(2, '0');
      }
      let idProcesamiento = mesAnioStr + "_FP";

      let nombreAsegurado = filaNueva[2] ? filaNueva[2].toString().trim().toUpperCase() : "";
      let encontrado = clientesMap[nombreAsegurado] || ""; 
      if (encontrado && typeof encontrado === "string") {
        encontrado = encontrado.replace("M-", "");
      }
      filaNueva.push(encontrado);

      let tieneDato = filaNueva[0] !== "" && filaNueva[0] != null;
      filaNueva.push(tieneDato ? "FP" : "");

      let nombreProductor = filaNueva[0] ? filaNueva[0].toString().trim().toUpperCase() : "";
      if (nombreProductor.includes("MATIAS")) {
        filaNueva.push("MATIAS");
      } else {
        filaNueva.push(tieneDato ? "DGM" : "");
      }
      
      filaNueva.push(idProcesamiento);
      todosLosDatosNuevos.push(filaNueva);
    }
  }

  if (todosLosDatosNuevos.length > 0) {
    let encabezados = ["PRODUCTOR_ASOCIADO", "FECHA_EMISION", "NRO_POLIZA", "PRIMA_PESOS", "PREMIO_PESOS", "IMPORTE_COMISION_PESOS", "COD_ASEGURADO", "RAMO_CODIGO", "ID FedPatronal", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"];
    let resultadoFinal = [encabezados].concat(todosLosDatosNuevos);

    // 1. LIMPIEZA TOTAL
    sheet.clearContents();
    
    // 2. ESCRITURA DE DATOS
    sheet.getRange(1, 1, resultadoFinal.length, 12).setValues(resultadoFinal);

    // 3. INYECCIÓN DE LA FÓRMULA ARRAY EN M1 Y M2
    sheet.getRange("M1").setValue("AUXILIAR_BUSQUEDA");
    sheet.getRange("M2").setFormula(`=ARRAYFORMULA(IF(A2:A<>"";XLOOKUP(D2:D;AUX!A2:A;AUX!B2:B;"");""))`);

    const ultimaFila = sheet.getLastRow();
    sheet.getRange(2, 2, ultimaFila - 1, 1).setNumberFormat("dd/mm/yyyy"); 
    sheet.getRange(2, 4, ultimaFila - 1, 3).setNumberFormat("#,##0"); 
    sheet.autoResizeColumns(1, 13);
    console.log("✅ FP procesado e inyectada su ARRAYFORMULA en M2.");
  } else {
    console.log("No se encontraron archivos o datos nuevos para procesar en FP.");
  }
}