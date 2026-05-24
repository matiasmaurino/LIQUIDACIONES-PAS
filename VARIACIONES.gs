/**
 * VERSIÓN ULTRA-OPTIMIZADA: Genera el reporte con variaciones basándose
 * en los datos consolidados de "LIQUIDACIONES AGRUPADAS POR MES".
 * Ejecución en milisegundos al evitar procesar el detalle transaccional.
 */
function generarTablaConVariaciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. OBTENER DATOS DEL RESUMEN MENSUAL YA CONSOLIDADO
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
  // Estructura de origen: [Año-mes, PAS, RAMO_NOMBRE, CIA, PRIMA, PREMIO, COMISION, CANT POLIZAS]
  for (var i = 1; i < datosBase.length; i++) {
    var fila = datosBase[i];
    var mesAnio = fila[0] ? fila[0].toString().trim() : "";
    var pasAgrupado = fila[1] ? fila[1].toString().trim().toUpperCase() : "";
    var ramo = fila[2] ? fila[2].toString().trim() : "";
    var cia = fila[3] ? fila[3].toString().trim().toUpperCase() : "";
    var comision = parseFloat(fila[6]) || 0;
    var certificados = parseInt(fila[7]) || 0;
    
    if (!pasAgrupado || !mesAnio || mesAnio === "Año-mes") continue;
    
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
  
  // 3. CONSTRUIR LA NUEVA TABLA CON VARIACIONES
  var resultadoFinal = [];
  var filaCabecera = ["FECHA", "CIA", "RAMO_NOMBRE"];
  
  listaPas.forEach(function(pas) {
    filaCabecera.push(
      pas + " COMISION", 
      pas + " CANT POLIZAS", 
      pas + " VAR. COMISION", 
      pas + " VAR. POLIZAS"
    );
  });
  resultadoFinal.push(filaCabecera);
  
  // Función auxiliar rápida para calcular mes anterior
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
  
  // Calcular desvíos relativos en memoria RAM
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
  
  // 4. ESCRIBIR EN LA PESTAÑA DESTINO
  var nombreNuevaHoja = "REPORTE CON VARIACIONES";
  var hojaDestino = ss.getSheetByName(nombreNuevaHoja);
  if (hojaDestino) {
    hojaDestino.clear(); 
  } else {
    hojaDestino = ss.insertSheet(nombreNuevaHoja);
  }
  
  hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
  
  // Formato estético rápido
  hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).setBackground("#f3f3f3").setFontWeight("bold");
  hojaDestino.autoResizeColumns(1, hojaDestino.getLastColumn());
  
  console.log("✅ Reporte con variaciones generado exitosamente en bloque.");
}