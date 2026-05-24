/**
 * FUNCIÓN ACTUALIZADA (CON PARSEO DE Ñ): Procesa los CSV de Federación Patronal, 
 * los relaciona con el ID de cliente y realiza la carga pura al final.
 * Corrige el error de codificación transformando "MAURIÃ‘O" en "MAURIÑO".
 */
function consolidarYLimpiarFP() {
  const folderId = '1MFWeyrluXJdDA8pJyuzAGAeRRAOeIHbV';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. CARGAR DICCIONARIO DE CLIENTES (En memoria)
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
    sheet.appendRow(["PRODUCTOR_ASOCIADO", "FECHA_EMISION", "NRO_POLIZA", "PRIMA_PESOS", "PREMIO_PESOS", "IMPORTE_COMISION_PESOS", "COD_ASEGURADO", "RAMO_CODIGO", "ID FedPatronal", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"]);
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
      
      // Extraer columnas base (8 campos) y limpiar la codificación rota de caracteres
      let filaNueva = columnasInteres.map((idx, colPos) => {
        let valor = fila[idx] ? fila[idx].toString().trim() : "";
        
        // REPARACIÓN CRÍTICA DE Ñ: Traduce el error de codificación del CSV original
        if (valor.includes("Ã‘")) {
          valor = valor.replace(/Ã‘/g, "Ñ");
        }
        if (valor.includes("Ã³")) {
          valor = valor.replace(/Ã³/g, "ó");
        }

        if (colPos === 1 && valor.includes("-")) {
          let partes = valor.substring(0, 10).split('-');
          if (partes.length === 3) return partes[2] + "/" + partes[1] + "/" + partes[0];
        }
        return valor;
      });

      // Calcular el Año-Mes para el ID de lote
      let fechaFila = filaNueva[1];
      let mesAnioStr = "INDEFINIDO";
      if (fechaFila && fechaFila.includes("/")) {
        let partesF = fechaFila.split("/");
        if (partesF.length === 3) mesAnioStr = partesF[2].trim() + "-" + partesF[1].trim().padStart(2, '0');
      }
      let idProcesamiento = mesAnioStr + "_FP";

      // 9. ID FedPatronal
      let nombreAsegurado = filaNueva[2] ? filaNueva[2].toString().trim().toUpperCase() : "";
      let encontrado = clientesMap[nombreAsegurado] || ""; 
      if (encontrado && typeof encontrado === "string") {
        encontrado = encontrado.replace("M-", "");
      }
      filaNueva.push(encontrado);

      // 10. CIA
      let tieneDato = filaNueva[0] !== "" && filaNueva[0] != null;
      filaNueva.push(tieneDato ? "FP" : "");

      // 11. PAS AGRUPADO
      let nombreProductor = filaNueva[0] ? filaNueva[0].toString().trim().toUpperCase() : "";
      
      // También corregimos la validación por si viene roto el texto en la primera columna
      if (nombreProductor === "MAURIÑO MATIAS" || nombreProductor === "MAURIÑO MATIAS" || nombreProductor.includes("MATIAS")) {
        filaNueva.push("MATIAS");
      } else {
        filaNueva.push(tieneDato ? "DGM" : "");
      }
      
      // 12. ID_PROCESAMIENTO
      filaNueva.push(idProcesamiento);
      
      todosLosDatosNuevos.push(filaNueva);
    }
  }

  // 2. INSERCIÓN DIRECTA AL FINAL DE LA HOJA
  if (todosLosDatosNuevos.length > 0) {
    sheet.getRange(1, 12).setValue("ID_PROCESAMIENTO");
    sheet.getRange(sheet.getLastRow() + 1, 1, todosLosDatosNuevos.length, todosLosDatosNuevos[0].length).setValues(todosLosDatosNuevos);

    const ultimaFila = sheet.getLastRow();
    const filasAgregadas = todosLosDatosNuevos.length;
    const filaInicioFormatos = ultimaFila - filasAgregadas + 1;

    sheet.getRange(filaInicioFormatos, 2, filasAgregadas, 1).setNumberFormat("dd/mm/yyyy"); 
    sheet.getRange(filaInicioFormatos, 4, filasAgregadas, 3).setNumberFormat("#,##0"); 
    sheet.autoResizeColumns(1, 12); 
    
    console.log("✅ FP procesado correctamente con nombres normalizados.");
  }
}