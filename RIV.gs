function consolidarYLimpiarRIV() {
  const folderId = '1cL4zPoAD7YJjhhHMZvg33vBo0kB5OsHZ'; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName("RIV");
  if (!sheet) sheet = ss.insertSheet("RIV");
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    // Índices originales Excel: 0(Fecha), 3(Asegurado), 4(Ramo), 5(Poliza), 10(Prima), 8(Premio), 11(Comisión), 2(Matrícula)
    const columnasInteres = [0, 3, 4, 5, 10, 8, 11, 2]; 
    
    let todosLosDatos = [];

    while (files.hasNext()) {
      let file = files.next();
      let fileName = file.getName().toLowerCase();
      
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        let blob = file.getBlob();
        
        // Configuración compatible con Drive API v3 usando recurso básico
        let resource = {
          title: "[TEMP] " + file.getName(),
          mimeType: MimeType.GOOGLE_SHEETS,
          parents: [{id: folderId}]
        };
        
        // Conversión nativa segura
        let tempFile = Drive.Files.create(resource, blob);
        let tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
        let tempSheet = tempSpreadsheet.getSheets()[0];
        let fullData = tempSheet.getDataRange().getValues();
        
        // Eliminamos el archivo temporal inmediatamente después de leerlo
        DriveApp.getFileById(tempFile.id).setTrashed(true);

        if (!fullData || fullData.length <= 1) continue;

        // Recorremos los registros saltando siempre la fila 0 del Excel original (sus títulos nativos)
        let datosProcesados = fullData.slice(1).map((fila) => {
          if (fila.length === 0) return null;

          let columnasExtraidas = columnasInteres.map(idx => {
            let valor = fila[idx];
            if (valor === undefined || valor === null) return "";
            
            // Tratamiento de la fecha original de Rivadavia
            if (columnasInteres.indexOf(idx) === 0) {
              if (valor instanceof Date) {
                return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
              }
              if (typeof valor === 'number') {
                let fechaBase = new Date(1899, 11, 30);
                fechaBase.setDate(fechaBase.getDate() + valor);
                return Utilities.formatDate(fechaBase, Session.getScriptTimeZone(), "dd/MM/yyyy");
              }
            }
            return valor.toString().trim();
          });

          let tieneDato = fila[0] !== "" && fila[0] != null; 
          let nombreProductor = tieneDato ? "MAURIÑO DANIEL GUILLERMO" : "";
          
          let compania = tieneDato ? "RIV" : "";
          let pasAgrupado = tieneDato ? "DGM" : "";

          // Retornamos la fila armada de datos puros
          return [nombreProductor].concat(columnasExtraidas).concat([compania, pasAgrupado]);
        }).filter(fila => fila !== null);

        todosLosDatos = todosLosDatos.concat(datosProcesados);
      }
    }

    if (todosLosDatos.length > 0) {
      // Limpiamos por completo la hoja para evitar restos viejos
      sheet.clear(); 
      
      // ESCRIBIMOS LOS ENCABEZADOS IDENTICOS EN LA FILA 1
      sheet.appendRow(["PRODUCTOR", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRIMA", "PREMIO", "COMISION", "MATRICULA", "CIA", "PAS AGRUPADO"]);
      
      // Pegamos los datos procesados exactamente a partir de la fila 2
      sheet.getRange(2, 1, todosLosDatos.length, todosLosDatos[0].length).setValues(todosLosDatos);
      
      const ultimaFila = sheet.getLastRow();
      if (ultimaFila > 1) {
        sheet.getRange(2, 2, ultimaFila - 1, 1).setNumberFormat("dd/mm/yyyy");
        sheet.getRange(2, 6, ultimaFila - 1, 3).setNumberFormat("#,##0"); 
        sheet.getRange(2, 9, ultimaFila - 1, 1).setNumberFormat("@"); 
        sheet.autoResizeColumns(1, 11);
      }
    }

    console.log("✅ RIV procesado correctamente con sus encabezados fijos en la primera fila.");

  } catch (e) {
    console.log("Error en RIV: " + e.message);
  }
}