/**
 * Crea el menú personalizado en la barra de herramientas de la planilla
 * cada vez que se abre el archivo de Google Sheets.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💼 Gestión Integral de clientes y liquidaciones')
    .addItem('🚀 Ejecutar Importación Completa', 'ejecutarProcesoCompletoLiquidaciones')
    .addSeparator()
    .addItem('📊 Actualizar Solo Reporte de Variaciones', 'ejecutarVariacionesManualmente') // Función puente añadida
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ Pasos Individuales')
       .addItem('1. Importar Clientes FP (CSV)', 'importarCSVsDesdeCarpetas')
       .addItem('2. Consolidar Liquidación RIV', 'consolidarYLimpiarRIV')
       .addItem('3. Consolidar Liquidación PS', 'consolidarYLimpiarPS')
       .addItem('4. Consolidar Liquidación FP', 'consolidarYLimpiarFP')
       .addItem('5. Agrupar Clientes e IDs', 'consolidarYAgruparSeguros'))
    .addToUi();
}

/**
 * Función Orquestadora Única: Corre las importaciones y actualiza las variaciones.
 */
function ejecutarProcesoCompletoLiquidaciones() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirmación de seguridad
  const respuesta = ui.alert('💼 Confirmación', '¿Deseás procesar las tres liquidaciones, actualizar Looker Studio y el Reporte de Variaciones?', ui.ButtonSet.YES_NO);
  if (respuesta !== ui.Button.YES) return;
  
  try {
    // 1. Ejecutar las importaciones individuales de las carpetas de Drive
    console.log("Importando maestros de clientes de Federación Patronal...");
    importarCSVsDesdeCarpetas();
    
    console.log("Consolidando liquidación Rivadavia...");
    consolidarYLimpiarRIV();
    
    console.log("Consolidando liquidación Provincia Seguros...");
    consolidarYLimpiarPS();
    
    console.log("Consolidando liquidación Federación Patronal...");
    consolidarYLimpiarFP();
    
    // 2. AGRUPAR CLIENTES Y DISPARAR REPORTE ANALÍTICO
    console.log("Generando padrón unificado y cálculos de porcentajes por ramo...");
    consolidarYAgruparSeguros();
    
    // 3. EJECUTAR EL REPORTE MATRICIAL CON VARIACIONES
    console.log("Generando reporte matricial horizontal con variaciones mensuales...");
    generarTablaConVariaciones();
    
    ui.alert('✅ Proceso Terminado', 'Se importaron todas las compañías, tu fórmula consolidó la sábana histórica, se actualizó Looker Studio y el Reporte con Variaciones ya está listo.', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error("Error en proceso: " + error.toString());
    ui.alert('❌ Error', 'Ocurrió un problema durante la importación: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Función puente para cuando un usuario ejecuta SOLO el reporte desde el menú.
 * Evita errores de interfaz y añade confirmación visual de éxito.
 */
function ejecutarVariacionesManualmente() {
  try {
    generarTablaConVariaciones();
    SpreadsheetApp.getUi().alert('✅ Reporte Actualizado', 'El Reporte con Variaciones ha sido recalculado con éxito.', ui.ButtonSet.OK);
  } catch (e) {
    console.error("Error al ejecutar manualmente: " + e.toString());
  }
}

/**
 * Procesa la sábana consolidada por tu fórmula y genera la estructura matricial con variaciones.
 */
function generarTablaConVariaciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // CORRECCIÓN: Nombre de pestaña cambiado para coincidir con tu Google Sheets real ("LIQUIDACIONES AGRUPADAS POR MES")
  var nombreHojaOrigen = "LIQUIDACIONES AGRUPADAS POR MES"; 
  var hojaOrigen = ss.getSheetByName(nombreHojaOrigen);
  
  if (!hojaOrigen) {
    console.warn("No se encontró la pestaña '" + nombreHojaOrigen + "'");
    // Alerta segura por si se corre en un contexto con UI
    try { SpreadsheetApp.getUi().alert("No se encontró la pestaña '" + nombreHojaOrigen + "'"); } catch(e){}
    return;
  }
  
  var datosBase = hojaOrigen.getDataRange().getValues();
  
  var listaPas = [];
  var matrizAgrupada = {};
  
  for (var i = 1; i < datosBase.length; i++) {
    var fila = datosBase[i];
    
    var fechaOriginal = fila[1]; // Col B (FECHA)
    var poliza = fila[4];        // Col E (POLIZA)
    var comision = parseFloat(fila[7]) || 0; // Col H (COMISION)
    var cia = fila[9];           // Col J (CIA)
    var pasAgrupado = fila[10];  // Col K (PAS AGRUPADO)
    var ramo = fila[11];         // Col L (RAMO_NOMBRE)
    
    if (!pasAgrupado || !fechaOriginal || fechaOriginal === "FECHA" || fechaOriginal === "") continue;
    
    // Formatear Fecha de manera segura
    var mesAnio = "";
    if (fechaOriginal instanceof Date) {
      var mm = (fechaOriginal.getMonth() + 1).toString().padStart(2, '0');
      mesAnio = fechaOriginal.getFullYear() + "-" + mm;
    } else {
      var partes = fechaOriginal.toString().split(/[-/]/);
      if (partes.length >= 3) {
        if (partes[0].length === 4) {
          mesAnio = partes[0] + "-" + padding(partes[1]);
        } else {
          mesAnio = partes[2] + "-" + padding(partes[1]);
        }
      } else {
        mesAnio = fechaOriginal.toString().substring(0,7);
      }
    }
    
    if (listaPas.indexOf(pasAgrupado) === -1) listaPas.push(pasAgrupado);
    
    var claveFila = mesAnio + "|" + cia + "|" + ramo;
    
    if (!matrizAgrupada[claveFila]) {
      matrizAgrupada[claveFila] = {
        mesAnio: mesAnio,
        cia: cia,
        ramo: ramo,
        valoresPas: {}
      };
    }
    
    if (!matrizAgrupada[claveFila].valoresPas[pasAgrupado]) {
      matrizAgrupada[claveFila].valoresPas[pasAgrupado] = { comision: 0, certificados: 0, polizasVistas: {} };
    }
    
    matrizAgrupada[claveFila].valoresPas[pasAgrupado].comision += comision;
    
    if (!matrizAgrupada[claveFila].valoresPas[pasAgrupado].polizasVistas[poliza]) {
      matrizAgrupada[claveFila].valoresPas[pasAgrupado].polizasVistas[poliza] = true;
      matrizAgrupada[claveFila].valoresPas[pasAgrupado].certificados += 1;
    }
  }
  
  listaPas.sort();
  var clavesOrdenadas = Object.keys(matrizAgrupada).sort();
  
  var resultadoFinal = [];
  var filaCabecera = ["FECHA", "CIA", "RAMO_NOMBRE"];
  
  listaPas.forEach(function(pas) {
    filaCabecera.push(pas + " COMISION", pas + " CANT POLIZAS", pas + " VAR. COMISION", pas + " VAR. POLIZAS");
  });
  resultadoFinal.push(filaCabecera);
  
  function obtenerMesAnterior(mesAnioStr) {
    var partes = mesAnioStr.split("-");
    var año = parseInt(partes[0]);
    var mes = parseInt(partes[1]) - 1;
    if (mes === 0) {
      mes = 12;
      año -= 1;
    }
    return año + "-" + mes.toString().padStart(2, '0');
  }
  
  function padding(str) {
    return str.toString().trim().padStart(2, '0');
  }
  
  clavesOrdenadas.forEach(function(clave) {
    var nodo = matrizAgrupada[clave];
    var filaDestino = [nodo.mesAnio, nodo.cia, nodo.ramo];
    
    var mesAnterior = obtenerMesAnterior(nodo.mesAnio);
    var claveMesAnterior = mesAnterior + "|" + nodo.cia + "|" + nodo.ramo;
    var nodoAnterior = matrizAgrupada[claveMesAnterior];
    
    listaPas.forEach(function(pas) {
      var datosActuales = nodo.valoresPas[pas] || { comision: 0, certificados: 0 };
      var datosAnteriores = (nodoAnterior && nodoAnterior.valoresPas[pas]) ? nodoAnterior.valoresPas[pas] : null;
      
      var comisionActual = datosActuales.comision;
      var certActual = datosActuales.certificados;
      
      var varComision = datosAnteriores ? (comisionActual - datosAnteriores.comision) : comisionActual;
      var varCert = datosAnteriores ? (certActual - datosAnteriores.certificados) : certActual;
      
      filaDestino.push(comisionActual, certActual, varComision, varCert);
    });
    
    resultadoFinal.push(filaDestino);
  });
  
  var nombreNuevaHoja = "REPORTE CON VARIACIONES";
  var hojaDestino = ss.getSheetByName(nombreNuevaHoja);
  if (hojaDestino) {
    hojaDestino.clear(); 
  } else {
    hojaDestino = ss.insertSheet(nombreNuevaHoja);
  }
  
  hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
  hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).setBackground("#f3f3f3").setFontWeight("bold");
  hojaDestino.autoResizeColumns(1, hojaDestino.getLastColumn());
  
  console.log("¡Reporte con variaciones procesado internamente con éxito!");
}