/**
 * REPORTE CON VARIACIONES: Versión Simplificada
 * Aplica formato de número entero a todas las columnas de métricas.
 */
function generarTablaConVariaciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. OBTENER DATOS DEL RESUMEN MENSUAL
  var hojaOrigen = ss.getSheetByName("LIQUIDACIONES AGRUPADAS POR MES");
  if (!hojaOrigen) {
    SpreadsheetApp.getUi().alert("No se encontró la pestaña 'LIQUIDACIONES AGRUPADAS POR MES'");
    return;
  }
  
  var datosBase = hojaOrigen.getDataRange().getValues();
  if (datosBase.length <= 1) return; // Sin datos
  
  var matrizAgrupada = {};
  var listaPas = [];
  
  // 2. PROCESAR EL RESUMEN EN MEMORIA
  for (var i = 1; i < datosBase.length; i++) {
    var fila = datosBase[i];
    var mesAnio = fila[0] ? fila[0].toString().trim() : "";
    var pasAgrupado = fila[1] ? fila[1].toString().trim().toUpperCase() : "";
    var ramo = fila[2] ? fila[2].toString().trim() : "";
    var cia = fila[3] ? fila[3].toString().trim().toUpperCase() : "";
    var comision = parseFloat(fila[6]) || 0;
    var certificados = parseInt(fila[7]) || 0;
    
    if (!pasAgrupado || !mesAnio || mesAnio === "Año-mes" || mesAnio.toLowerCase().includes("total")) continue;
    
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
      matrizAgrupada[claveFila].valoresPas[pasAgrupado] = { comision: 0, certificados: 0 };
    }
    
    matrizAgrupada[claveFila].valoresPas[pasAgrupado].comision += comision;
    matrizAgrupada[claveFila].valoresPas[pasAgrupado].certificados += certificados;
  }
  
  listaPas.sort();
  var clavesOrdenadas = Object.keys(matrizAgrupada).sort();
  
  // 3. CONSTRUIR LA NUEVA TABLA CON VARIACIONES MENSUALES Y ANUALES (AA)
  var resultadoFinal = [];
  var filaCabecera = ["FECHA", "CIA", "RAMO_NOMBRE"];
  
  listaPas.forEach(function(pas) {
    filaCabecera.push(
      pas + " COMISION", 
      pas + " CANT POLIZAS", 
      pas + " VAR. COMISION", 
      pas + " VAR. POLIZAS",
      pas + " VAR. COMISION AA", // Año Anterior
      pas + " VAR. POLIZAS AA"   // Año Anterior
    );
  });
  resultadoFinal.push(filaCabecera);
  
  // FUNCION AUXILIAR 1: Calcula el mes inmediato anterior
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
  
  // FUNCION AUXILIAR 2: Calcula el mismo mes del año anterior (AA)
  function obtenerMesAñoAnterior(mesAnioStr) {
    var partes = mesAnioStr.split("-");
    var año = parseInt(partes[0]) - 1;
    var mes = partes[1];
    return año + "-" + mes;
  }
  
  // 4. CALCULAR DESVÍOS RELATIVOS EN MEMORIA RAM
  clavesOrdenadas.forEach(function(clave) {
    var nodo = matrizAgrupada[clave];
    var filaDestino = [nodo.mesAnio, nodo.cia, nodo.ramo];
    
    var mesAnterior = obtenerMesAnterior(nodo.mesAnio);
    var claveMesAnterior = mesAnterior + "|" + nodo.cia + "|" + nodo.ramo;
    var nodoAnterior = matrizAgrupada[claveMesAnterior];
    
    var mesAñoAnterior = obtenerMesAñoAnterior(nodo.mesAnio);
    var claveAñoAnterior = mesAñoAnterior + "|" + nodo.cia + "|" + nodo.ramo;
    var nodoAñoAnterior = matrizAgrupada[claveAñoAnterior];
    
    listaPas.forEach(function(pas) {
      var datosActuales = nodo.valoresPas[pas] || { comision: 0, certificados: 0 };
      var comisionActual = datosActuales.comision;
      var certActual = datosActuales.certificados;
      
      var datosAnteriores = (nodoAnterior && nodoAnterior.valoresPas[pas]) ? nodoAnterior.valoresPas[pas] : null;
      var varComision = datosAnteriores ? (comisionActual - datosAnteriores.comision) : comisionActual;
      var varCert = datosAnteriores ? (certActual - datosAnteriores.certificados) : certActual;
      
      var datosAñoAnterior = (nodoAñoAnterior && nodoAñoAnterior.valoresPas[pas]) ? nodoAñoAnterior.valoresPas[pas] : null;
      var varComisionAA = datosAñoAnterior ? (comisionActual - datosAñoAnterior.comision) : comisionActual;
      var varCertAA = datosAñoAnterior ? (certActual - datosAñoAnterior.certificados) : certActual;
      
      filaDestino.push(comisionActual, certActual, varComision, varCert, varComisionAA, varCertAA);
    });
    
    resultadoFinal.push(filaDestino);
  });
  
  // 5. ESCRIBIR EN LA PESTAÑA DESTINO
  var nombreNuevaHoja = "REPORTE CON VARIACIONES";
  var hojaDestino = ss.getSheetByName(nombreNuevaHoja);
  if (hojaDestino) {
    hojaDestino.clear(); 
  } else {
    hojaDestino = ss.insertSheet(nombreNuevaHoja);
  }
  
  hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
  
  // 6. APLICAR FORMATOS MASIVOS SIMPLIFICADOS
  var ultimaFila = hojaDestino.getLastRow();
  var ultimaColumna = hojaDestino.getLastColumn();
  
  if (ultimaFila > 1) {
    // Columnas A, B, C como Texto plano
    hojaDestino.getRange(2, 1, ultimaFila - 1, 3).setNumberFormat("@");

    // SOLUCIÓN TOTAL: Todo el bloque métrico completo (de la columna D hasta el final) como Número Entero
    hojaDestino.getRange(2, 4, ultimaFila - 1, ultimaColumna - 3).setNumberFormat("#,##0");
  }
  
  // Formato estético rápido de cabeceras
  hojaDestino.getRange(1, 1, 1, ultimaColumna).setBackground("#f3f3f3").setFontWeight("bold");
  hojaDestino.autoResizeColumns(1, ultimaColumna);
  console.log("✅ Reporte con variaciones generado y formateado como entero de forma simplificada.");
}