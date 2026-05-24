function consolidarYLimpiarRIV() {
  const folderId = '1cL4zPoAD7YJjhhHMZvg33vBo0kB5OsHZ'; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet = ss.getSheetByName("RIV");
  if (!sheet) sheet = ss.insertSheet("RIV");
  
  try {
    const query = `'${folderId}' in parents and trashed = false`;
    const filesList = Drive.Files.list({
      q: query,
      fields: 'files(id, name)'
    }).files;

    if (!filesList || filesList.length === 0) {
      SpreadsheetApp.getUi().alert("No se encontraron archivos.");
      return;
    }

    // Índices originales Excel: 0(Fecha), 3(Asegurado), 4(Ramo), 5(Poliza), 10(Prima), 8(Premio), 11(Comisión), 2(Matrícula)
    const columnasInteres = [0, 3, 4, 5, 10, 8, 11, 2]; 
    
    let todosLosDatos = [];
    let isFirstFile = true;

    for (const file of filesList) {
      let fileName = file.name.toLowerCase();
      
      if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        let fileSource = DriveApp.getFileById(file.id);
        let blob = fileSource.getBlob();
        let resource = { name: "temp_convert", mimeType: MimeType.GOOGLE_SHEETS };
        
        let tempFile = Drive.Files.create(resource, blob);
        let tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
        let tempSheet = tempSpreadsheet.getSheets()[0];
        let excelData = tempSheet.getDataRange().getValues();
        
        Drive.Files.remove(tempFile.id);

        let datosProcesados = excelData.map((fila, index) => {
          if (index === 0 && !isFirstFile) return null;

          let columnasExtraidas = columnasInteres.map((idx, colPos) => {
            let valor = fila[idx] != null ? fila[idx] : "";

            // 1. LIMPIEZA DE ASEGURADO (Quitar coma entre apellido y nombre)
            if (colPos === 1 && index > 0) {
              valor = String(valor).replace(/,/g, "").trim();
            }

            // Limpieza caracteres UTF-8
            if (typeof valor === "string") {
              valor = valor.replace(/Ã³/g, "ó").replace(/Ã¡/g, "á")
                           .replace(/Ã©/g, "é").replace(/Ã/g, "Ñ")
                           .replace(/[^\x20-\x7E\xC0-\xFF/]/g, "").trim();
            }

            // 2. FECHA
            if (colPos === 0 && valor !== "" && index > 0) {
              let d = (valor instanceof Date) ? valor : new Date(valor);
              if (!isNaN(d.getTime())) return d;
            }

            // 3. PÓLIZA (Quitar prefijo 47-)
            if (colPos === 3 && typeof valor === "string" && index > 0) {
              value = valor.replace(/^\d{2}-/, "");
            }

            // 4. CORRECCIÓN DE NÚMEROS (PRI.COB, PAGOS, COM.DEV)
            if ((colPos >= 4 && colPos <= 6) && valor !== "" && index > 0) {
              let strValor = String(valor).trim();
              if (strValor.includes(",") && strValor.includes(".")) {
                strValor = strValor.replace(/\./g, "").replace(",", ".");
              } 
              else if (strValor.includes(",")) {
                strValor = strValor.replace(",", ".");
              }
              
              let num = parseFloat(strValor);
              return isNaN(num) ? valor : Math.trunc(num);
            }

            // 5. MATRÍCULA (Solo números)
            if (colPos === 7 && valor !== "" && index > 0) {
              return String(valor).replace(/\D/g, "");
            }

            return valor;
          });

          // SECCIÓN DE ARMADO DE FILA LOGICAL (Columnas A hasta K)
          if (index === 0) {
            // Encabezados exactos para la fila 1
            return ["Nombre Productor", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRI.COB", "PAGOS", "COM.DEV", "MATRIC.", "CIA", "PAS AGRUPADO"];
          } else {
            let tieneDato = fila[0] !== "" && fila[0] != null; 
            let nombreProductor = tieneDato ? "MAURIÑO DANIEL GUILLERMO" : "";
            
            // Si hay productor, completamos CIA (J) y PAS AGRUPADO (K)
            let compania = tieneDato ? "RIV" : "";
            let pasAgrupado = tieneDato ? "DGM" : "";

            // Unimos todo en el orden correcto: [A] + [B a I] + [J, K]
            return [nombreProductor].concat(columnasExtraidas).concat([compania, pasAgrupado]);
          }
        }).filter(fila => fila !== null);

        todosLosDatos = todosLosDatos.concat(datosProcesados);
        isFirstFile = false;
      }
    }

    if (todosLosDatos.length > 0) {
      sheet.clear(); 
      sheet.getRange(1, 1, todosLosDatos.length, todosLosDatos[0].length).setValues(todosLosDatos);
      
      const ultimaFila = sheet.getLastRow();
      if (ultimaFila > 1) {
        sheet.getRange(2, 2, ultimaFila - 1, 1).setNumberFormat("dd/mm/yyyy");
        sheet.getRange(2, 6, ultimaFila - 1, 3).setNumberFormat("#,##0"); 
        sheet.getRange(2, 9, ultimaFila - 1, 1).setNumberFormat("@"); 
        sheet.autoResizeColumns(1, 11); // Modificado a 11 columnas para incluir J y K
      }
    }
    
    SpreadsheetApp.getUi().alert("✅ Procesado correctamente. Se arreglaron los valores numéricos y se integró PAS AGRUPADO.");
    
  } catch (err) {
    Logger.log(err.toString());
    SpreadsheetApp.getUi().alert("Error: " + err.message);
  }
}