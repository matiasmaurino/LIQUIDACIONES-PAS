function generarResumenMensualConPorcentajes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Obtener los datos consolidados de origen
  var hojaOrigen = ss.getSheetByName("LIQUIDACIONES AGRUPADAS");
  if (!hojaOrigen) return;
  var datosOrigen = hojaOrigen.getDataRange().getValues();
  
  // Identificar índices de columnas en el origen basándonos en tu QUERY actual:
  var idxFecha = 1;       // ColB (Fecha)
  var idxPas = 10;        // ColK (PAS AGRUPADO)
  var idxRamo = 11;       // ColL (RAMO_NOMBRE)
  var idxCia = 9;         // ColJ (CIA)
  var idxComision = 7;    // ColH (COMISION)
  
  // Objeto para agrupar los valores únicos por clave [Fecha | Cia | Ramo]
  var mapaAgrupado = {};
  
  // 2. Recorrer datos de origen (saltando el encabezado)
  for (var i = 1; i < datosOrigen.length; i++) {
    var fila = datosOrigen[i];
    
    var pas = fila[idxPas] ? fila[idxPas].toString().trim().toUpperCase() : "";
    var rawFecha = fila[idxFecha];
    var ramo = fila[idxRamo];
    var cia = fila[idxCia];
    var comision = parseFloat(fila[idxComision]) || 0;
    
    if (!pas || !rawFecha || rawFecha === "FECHA") continue;
    
    // Normalizar formato fecha a YYYY-MM si viene como tipo Date o texto largo
    var fechaStr = "";
    if (rawFecha instanceof Date) {
      fechaStr = Utilities.formatDate(rawFecha, ss.getSpreadsheetTimeZone(), "yyyy-MM");
    } else {
      fechaStr = rawFecha.toString().substring(0, 7); // Por si es texto completo
    }
    
    // Crear una clave única combinada de fila
    var clave = fechaStr + "_" + cia + "_" + ramo;
    
    if (!mapaAgrupado[clave]) {
      mapaAgrupado[clave] = {
        fecha: fechaStr,
        cia: cia,
        ramo: ramo,
        dgm: 0,
        matias: 0
      };
    }
    
    // Asignar e ir sumando comisiones según corresponda al PAS AGRUPADO
    if (pas === "DGM") {
      mapaAgrupado[clave].dgm += comision;
    } else if (pas === "MATIAS") {
      mapaAgrupado[clave].matias += comision;
    }
  }
  
  // 3. Procesar los datos agrupados y calcular totales/porcentajes
  var filasResultado = [];
  
  for (var k in mapaAgrupado) {
    var item = mapaAgrupado[k];
    var sumaTotal = item.dgm + item.matias;
    
    // Calcular porcentajes de forma segura (evitando división por cero)
    var pctDgm = sumaTotal > 0 ? (item.dgm / sumaTotal) : 0;
    var pctMatias = sumaTotal > 0 ? (item.matias / sumaTotal) : 0;
    
    filasResultado.push([
      item.fecha,
      item.cia,
      item.ramo,
      item.dgm,
      item.matias,
      sumaTotal,
      pctDgm,
      pctMatias
    ]);
  }
  
  // Ordenar el resultado por Fecha, Compañía y Ramo para mantener la estética
  filasResultado.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    if (a[1] !== b[1]) return a[1].localeCompare(b[1]);
    return a[2].localeCompare(b[2]);
  });
  
  // 4. Escribir los datos procesados en la solapa de destino
  var hojaDestino = ss.getSheetByName("LIQUIDACIONES AGRUPADAS POR MES") || ss.insertSheet("LIQUIDACIONES AGRUPADAS POR MES");
  
  if (filasResultado.length > 0) {
    hojaDestino.clear(); // Solo limpia si realmente hay datos para escribir
    var encabezados = ["FECHA", "CIA", "RAMO_NOMBRE", "DGM", "MATIAS", "Suma total", "% DGM", "% MATIAS"];
    hojaDestino.appendRow(encabezados);
    
    hojaDestino.getRange(2, 1, filasResultado.length, filasResultado[0].length).setValues(filasResultado);
    
    // 5. Dar formato profesional automático a las columnas de porcentaje y dinero
    hojaDestino.getRange(2, 4, filasResultado.length, 3).setNumberFormat("$#,##0"); // DGM, MATIAS, Suma total
    hojaDestino.getRange(2, 7, filasResultado.length, 2).setNumberFormat("0.00%");  // % DGM y % MATIAS
  }
}

// =================================================================
// FUNCIÓN NUEVA (EJECUTAR POR ÚNICA VEZ) - CORREGIDA
// =================================================================
function procesarHistoricoSheetGigante_UnaVez() {
  // 1. Conectarse al Google Sheet externo usando el ID proporcionado
  var idPlanillaOrigen = "1N8XRgzTnkHhOWHtSAP0rCEdHYlk5GEDHfEjEyBBNC9Q";
  var ssOrigen;
  
  try {
    ssOrigen = SpreadsheetApp.openById(idPlanillaOrigen);
  } catch (e) {
    Logger.log("No se pudo acceder al archivo origen. Verifica los permisos. Error: " + e.toString());
    return;
  }
  
  var hojaOrigen = ssOrigen.getSheetByName("LIQUIDACIONES AGRUPADAS");
  if (!hojaOrigen) {
    Logger.log("No se encontró la pestaña 'LIQUIDACIONES AGRUPADAS' en el archivo externo.");
    return;
  }
  
  // Trae las 226k filas a la memoria
  var datosOrigen = hojaOrigen.getDataRange().getValues();
  
  // Identificar índices de columnas (idénticos a tu función habitual)
  var idxFecha = 1;       // ColB (Fecha)
  var idxPas = 10;        // ColK (PAS AGRUPADO)
  var idxRamo = 11;       // ColL (RAMO_NOMBRE)
  var idxCia = 9;         // ColJ (CIA)
  var idxComision = 7;    // ColH (COMISION)
  
  var mapaAgrupado = {};
  var zonaHoraria = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
  // 2. Recorrer datos de origen optimizando formatos de texto y fechas
  for (var i = 1; i < datosOrigen.length; i++) {
    var fila = datosOrigen[i];
    
    // CORRECCIÓN CLAVE: Limpiar espacios y pasar a mayúsculas para asegurar coincidencia
    var pas = fila[idxPas] ? fila[idxPas].toString().trim().toUpperCase() : "";
    var rawFecha = fila[idxFecha];
    var ramo = fila[idxRamo];
    var cia = fila[idxCia];
    var comision = parseFloat(fila[idxComision]) || 0;
    
    if (!pas || !rawFecha || rawFecha === "FECHA") continue;
    
    // Normalizar formato fecha a YYYY-MM
    var fechaStr = "";
    if (rawFecha instanceof Date) {
      fechaStr = Utilities.formatDate(rawFecha, zonaHoraria, "yyyy-MM");
    } else {
      var stringFecha = rawFecha.toString().trim();
      if (stringFecha.includes("/")) {
        var partes = stringFecha.split("/"); 
        if (partes.length === 3) {
          // Asegura formato YYYY-MM manejando días/meses de un solo dígito
          var anio = partes[2].trim();
          var mes = ("0" + partes[1].trim()).slice(-2);
          fechaStr = anio + "-" + mes;
        }
      } else if (stringFecha.includes("-")) {
        fechaStr = stringFecha.substring(0, 7); 
      }
    }
    
    // Salta filas con fechas corruptas o vacías en el histórico
    if (!fechaStr || fechaStr.length !== 7) continue; 
    
    var clave = fechaStr + "_" + cia + "_" + ramo;
    
    if (!mapaAgrupado[clave]) {
      mapaAgrupado[clave] = {
        fecha: fechaStr,
        cia: cia,
        ramo: ramo,
        dgm: 0,
        matias: 0
      };
    }
    
    if (pas === "DGM") {
      mapaAgrupado[clave].dgm += comision;
    } else if (pas === "MATIAS") {
      mapaAgrupado[clave].matias += comision;
    }
  }
  
  // 3. Procesar los datos agrupados y calcular totales/porcentajes
  var filasResultado = [];
  
  for (var k in mapaAgrupado) {
    var item = mapaAgrupado[k];
    var sumaTotal = item.dgm + item.matias;
    
    var pctDgm = sumaTotal > 0 ? (item.dgm / sumaTotal) : 0;
    var pctMatias = sumaTotal > 0 ? (item.matias / sumaTotal) : 0;
    
    // Solo incluir en el resumen si hubo algún movimiento para DGM o MATIAS
    if (sumaTotal > 0) {
      filasResultado.push([
        item.fecha,
        item.cia,
        item.ramo,
        item.dgm,
        item.matias,
        sumaTotal,
        pctDgm,
        pctMatias
      ]);
    }
  }
  
  // Ordenar el resultado por Fecha, Compañía y Ramo
  filasResultado.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    if (a[1] !== b[1]) return a[1].localeCompare(b[1]);
    return a[2].localeCompare(b[2]);
  });
  
  // 4. Escribir los datos en tu planilla actual
  var ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ssDestino.getSheetByName("LIQUIDACIONES AGRUPADAS POR MES") || ssDestino.insertSheet("LIQUIDACIONES AGRUPADAS POR MES");
  
  if (filasResultado.length > 0) {
    hojaDestino.clear(); // Protegido: Solo borra si el proceso fue exitoso y hay datos
    
    var encabezados = ["FECHA", "CIA", "RAMO_NOMBRE", "DGM", "MATIAS", "Suma total", "% DGM", "% MATIAS"];
    hojaDestino.appendRow(encabezados);
    
    hojaDestino.getRange(2, 1, filasResultado.length, filasResultado[0].length).setValues(filasResultado);
    
    // 5. Formato automático a los montos y porcentajes resultantes
    hojaDestino.getRange(2, 4, filasResultado.length, 3).setNumberFormat("$#,##0"); 
    hojaDestino.getRange(2, 7, filasResultado.length, 2).setNumberFormat("0.00%");  
    Logger.log("¡Proceso histórico terminado con éxito! Filas resumidas generadas: " + filasResultado.length);
  } else {
    Logger.log("ATENCIÓN: El proceso terminó pero no encontró ninguna fila que coincida con 'DGM' o 'MATIAS' en la columna K. Revisa los datos del origen.");
  }
}