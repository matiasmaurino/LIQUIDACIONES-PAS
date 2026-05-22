/**
 * FUNCIÓN 1: Procesa los CSV de Federación Patronal, 
 * los relaciona con el ID de cliente y realiza la carga incremental mensual.
 */
function consolidarYLimpiarFP() {
  const folderId = '1MFWeyrluXJdDA8pJyuzAGAeRRAOeIHbV';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. CARGAR DICCIONARIO DE CLIENTES
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
  if (!sheet) sheet = ss.insertSheet(nombreHoja);
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.CSV);
  const columnasInteres = [2, 3, 6, 7, 8, 16, 17, 18]; 
  
  let todosLosDatosNuevos = [];
  let mesesAQuitar = {}; 

  while (files.hasNext()) {
    let file = files.next();
    let csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'));
    
    // Recorremos los datos del CSV saltando SIEMPRE la fila de títulos nativos (index === 0)
    for (let i = 1; i < csvData.length; i++) {
      let fila = csvData[i];
      
      // Validación original para saltar filas vacías o defectuosas
      if (!fila[2] || fila[2].toString().trim() === "" || fila[2].toString().includes("PRODUCTOR")) continue;

      let filaNueva = columnasInteres.map((idx, colPos) => {
        let valor = fila[idx] ? fila[idx].toString().trim() : "";
        
        // Limpieza de caracteres especiales
        valor = valor.replace(/Ã³/g, "ó").replace(/Ã/g, "Ñ").replace(/[^\x20-\x7E\xC0-\xFF/]/g, "");

        if (colPos === 2 && valor.includes("undefined")) {
          valor = valor.split('/').pop();
        }

        if (colPos === 1 && valor.includes("-")) {
          let partes = valor.substring(0, 10).split('-');
          if (partes.length === 3) return partes[2] + "/" + partes[1] + "/" + partes[0];
        }

        if (colPos >= 4) {
          let num = parseFloat(valor.replace(',', '')); 
          return isNaN(num) ? valor : Math.trunc(num);
        }
        return valor;
      });

      // --- DETECCIÓN DEL MES PARA LA PURGA INCREMENTAL ---
      let fechaFila = filaNueva[1]; 
      if (fechaFila && fechaFila.includes("/")) {
        let partesF = fechaFila.split("/");
        if (partesF.length === 3) {
          let mAnio = partesF[2].trim() + "-" + partesF[1].trim().padStart(2, '0');
          mesesAQuitar[mAnio] = true;
        }
      }

      // 1. Columna I: ID FedPatronal
      let nombreAsegurado = filaNueva[2] ? filaNueva[2].toString().trim().toUpperCase() : "";
      let encontrado = clientesMap[nombreAsegurado] || ""; 
      if (encontrado && typeof encontrado === "string") {
        encontrado = encontrado.replace("M-", "");
      }
      filaNueva.push(encontrado);

      // 2. Columna J: CIA (Siempre FP si la fila tiene datos)
      let tieneDato = filaNueva[0] !== "" && filaNueva[0] != null;
      filaNueva.push(tieneDato ? "FP" : "");

      // 3. Columna K: PAS AGRUPADO (Lógica según el Productor)
      let nombreProductor = filaNueva[0] ? filaNueva[0].toString().trim().toUpperCase() : "";
      if (nombreProductor === "MAURIÑO MATIAS") {
        filaNueva.push("MATIAS");
      } else {
        filaNueva.push(tieneDato ? "DGM" : "");
      }
      
      todosLosDatosNuevos.push(filaNueva);
    }
  }

  // --- CONTROL Y BORRADO DE MESES REPETIDOS ---
  if (todosLosDatosNuevos.length > 0) {
    let datosHistoricos = sheet.getDataRange().getValues();
    
    if (datosHistoricos.length > 1) {
      for (let i = datosHistoricos.length - 1; i >= 1; i--) {
        let fHist = datosHistoricos[i][1]; // Columna B (FECHA)
        let mHist = "";
        if (fHist instanceof Date) {
          mHist = fHist.getFullYear() + "-" + (fHist.getMonth() + 1).toString().padStart(2, '0');
        } else if (fHist && fHist.toString().includes("/")) {
          let partesH = fHist.toString().split("/");
          if (partesH.length === 3) {
            mHist = partesH[2].trim() + "-" + partesH[1].trim().padStart(2, '0');
          }
        }
        
        if (mesesAQuitar[mHist]) {
          sheet.deleteRow(i + 1);
        }
      }
    }

    // Si por alguna razón la hoja se quedó sin filas, le reponemos tus encabezados limpios originales
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["PRODUCTOR_ASOCIADO", "FECHA_EMISION", "NRO_POLIZA", "PRIMA_PESOS", "PREMIO_PESOS", "IMPORTE_COMISION_PESOS", "COD_ASEGURADO", "RAMO_CODIGO", "ID FedPatronal", "CIA", "PAS AGRUPADO"]);
    }

    // Pegamos las filas de datos puros exactamente abajo del último registro
    sheet.getRange(sheet.getLastRow() + 1, 1, todosLosDatosNuevos.length, todosLosDatosNuevos[0].length).setValues(todosLosDatosNuevos);
  }

  // FORMATOS VISUALES
  const ultimaFila = sheet.getLastRow();
  if (ultimaFila > 1) {
    sheet.getRange(2, 2, ultimaFila - 1).setNumberFormat("dd/mm/yyyy"); 
    sheet.getRange(2, 5, ultimaFila - 1, 4).setNumberFormat("#,##0"); 
    sheet.autoResizeColumns(1, 11); 
  }

  // Se comenta la sincronización para evitar el ReferenceError ya que no está declarada en el proyecto
  // sincronizarNuevosAlCRM();

  // Alerta segura
  try {
    SpreadsheetApp.getUi().alert("✅ FP procesado correctamente de forma incremental y sin duplicar títulos.");
  } catch (e) {
    console.log("Proceso terminado: FP consolidado.");
  }
}