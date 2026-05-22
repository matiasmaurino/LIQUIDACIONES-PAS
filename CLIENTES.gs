function importarCSVsDesdeCarpetas() {
  // Configuración
  const folderIds = [
    '1QT78KtXO8svwN21HwCdkFWjSu9hF86i0', 
    '1a7gtjBghm5jWCG92GmMPFeJIouroeS30'
  ];
  const nombreHojaDestino = "CLIENTES FP";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nombreHojaDestino);
  
  // Si la hoja no existe, la crea
  if (!sheet) {
    sheet = ss.insertSheet(nombreHojaDestino);
  } else {
    sheet.clearContents(); 
  }

  let todosLosDatos = [];
  let esPrimerArchivo = true;

  folderIds.forEach(folderId => {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.CSV);

    while (files.hasNext()) {
      const file = files.next();
      let csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      
      if (csvData.length > 0) {
        // Determinamos desde qué fila empezar para evitar repetir encabezados
        let filaInicio = esPrimerArchivo ? 0 : 1;

        for (let i = filaInicio; i < csvData.length; i++) {
          let fila = csvData[i];
          
          // LÓGICA PARA LIMPIAR LA COLUMNA A (Matrícula)
          // Si el dato existe, quitamos "M-" o cualquier "M" y "-" al inicio
          if (fila[0]) {
            fila[0] = fila[0].toString().replace(/^M-/, "").replace(/^M/, "");
          }
          
          todosLosDatos.push(fila);
        }
        esPrimerArchivo = false;
      }
    }
  });

  // Pegar los datos en la hoja
  if (todosLosDatos.length > 0) {
    sheet.getRange(1, 1, todosLosDatos.length, todosLosDatos[0].length).setValues(todosLosDatos);
    SpreadsheetApp.getUi().alert('Importación completada. Se han limpiado las matrículas.');
  } else {
    SpreadsheetApp.getUi().alert('No se encontraron datos en las carpetas.');
  }
}