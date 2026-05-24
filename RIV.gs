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

          let mesAnioStr = "INDEFINIDO"; // Variable para rastrear el período del lote

          let columnasExtraidas = columnasInteres.map((idx, colPos) => {
            let valor = fila[idx] != null ? fila[idx] : "";

            // 1. LIMPIEZA DE ASEGURADO (Quitar coma entre apellido y nombre)
            if (colPos === 1 && index > 0) {
              valor = String(valor).replace(/,/g, "").trim();
            }

            // Limpieza caracteres UTF-8 y reparación de Ñ
            if (typeof valor === "string") {
              valor = valor.replace(/Ã³/g, "ó").replace(/Ã¡/g, "á")
                           .replace(/Ã©/g, "é").replace(/Ã‘/g, "Ñ")
                           .replace(/MAURIÃ‘O/g, "MAURIÑO")
                           .replace(/[^\x20-\x7E\xC0-\xFF/]/g, "").trim();
            }

            // 2. FECHA Y EXTRACCIÓN DE PERÍODO para ID LOTE
            if (colPos === 0 && valor !== "" && index > 0) {
              let d = (valor instanceof Date) ? valor : new Date(valor);
              if (!isNaN(d.getTime())) {
                let mm = (d.getMonth() + 1).toString().padStart(2, '0');
                mesAnioStr = d.getFullYear() + "-" + mm;
                return d;
              }
            }

            // 3. PÓLIZA (Quitar prefijo 47-) - CORREGIDO BUG TIPOGRÁFICO DE VARIABLE
            if (colPos === 3 && typeof valor === "string" && index > 0) {
              valor = valor.replace(/^\d{2}-/, "");
            }

            // 4. CORRECCIÓN DE NÚMEROS (CONSERVANDO DECIMALES REALES)
            if ((colPos >= 4 && colPos <= 6) && valor !== "" && index > 0) {
              let strValor = String(valor).trim();
              if (strValor.includes("温") || strValor === "-") strValor = "0";
              
              if (strValor.includes(",") && strValor.includes(".")) {
                strValor = strValor.replace(/\./g, "").replace(",", ".");
              } 
              else if (strValor.includes(",")) {
                strValor = strValor.replace(",", ".");
              }
              
              let num = parseFloat(strValor);
              // Conservamos centavos limitando a 2 decimales en lugar de truncar a entero
              return isNaN(num) ? valor : Number(num.toFixed(2));
            }

            // 5. MATRÍCULA (Solo números)
            if (colPos === 7 && valor !== "" && index > 0) {
              return String(valor).replace(/\D/g, "");
            }

            return valor;
          });

          // SECCIÓN DE ARMADO DE FILA DE 12 COLUMNAS (A hasta L)
          if (index === 0) {
            return ["Nombre Productor", "FECHA", "ASEGURADO", "RAMO", "POLIZA", "PRI.COB", "PAGOS", "COM.DEV", "MATRIC.", "CIA", "PAS AGRUPADO", "ID_PROCESAMIENTO"];
          } else {
            let tieneDato = fila[0] !== "" && fila[0] != null; 
            let nombreProductor = tieneDato ? "MAURIÑO DANIEL GUILLERMO" : "";
            let compania = tieneDato ? "RIV" : "";
            let pasAgrupado = tieneDato ? "DGM" : "";
            
            // Generar ID único basado en el mes extraído de la columna Fecha
            let idProcesamiento = mesAnioStr + "_RIV";

            // Retornamos las 12 columnas exactas integrando el lote al final (Col L)
            return [nombreProductor].concat(columnasExtraidas).concat([compania, pasAgrupado, idProcesamiento]);
          }
        }).filter(fila => fila !== null);

        todosLosDatos = todosLosDatos.concat(datosProcesados);
        isFirstFile = false;
      }
    }

    // 2. INYECCIÓN INCREMENTAL O ESCRITURA EN LIMPIO SEGÚN TU PREFERENCIA
    if (todosLosDatos.length > 0) {
      sheet.clear(); 
      sheet.getRange(1, 1, todosLosDatos.length, todosLosDatos[0].length).setValues(todosLosDatos);
      
      const ultimaFila = sheet.getLastRow();
      if (ultimaFila > 1) {
        sheet.getRange(2, 2, ultimaFila - 1, 1).setNumberFormat("dd/mm/yyyy");
        sheet.getRange(2, 6, ultimaFila - 1, 3).setNumberFormat("#,##0.00"); // Cambiado a .00 para ver los centavos en el Sheet
        sheet.getRange(2, 9, ultimaFila - 1, 1).setNumberFormat("@"); 
        sheet.autoResizeColumns(1, 12); // Extendible a la columna 12 (L)
      }
    }
    
    SpreadsheetApp.getUi().alert("✅ Procesado correctamente. Se integró la columna ID_PROCESAMIENTO y se mantuvieron los decimales contables.");
    
  } catch (err) {
    Logger.log(err.toString());
    SpreadsheetApp.getUi().alert("Error: " + err.message);
  }
}