function generarTablaConVariaciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var nombreHojaOrigen = "LIQUIDACIONES AGRUPADAS POR MES";
  var hojaOrigen = ss.getSheetByName(nombreHojaOrigen);
  
  if (!hojaOrigen) {
    console.warn("No se encontró la pestaña '" + nombreHojaOrigen + "'");
    try { SpreadsheetApp.getUi().alert("No se encontró la pestaña '" + nombreHojaOrigen + "'"); } catch(e){}
    return;
  }
  
  var datosBase = hojaOrigen.getDataRange().getValues();
  if (datosBase.length <= 1) return;
  
  var universoDatos = {};
  
  // PRIMERA PASADA: Indexar absolutamente todo el histórico en memoria
  for (var i = 1; i < datosBase.length; i++) {
    var fila = datosBase[i];
    var fechaOriginal = fila[0]; // Col A: Año-mes
    var pas = fila[1];           // Col B: PAS
    var cia = fila[2];           // Col C: CIA
    var ramo = fila[3];          // Col D: RAMO_NOMBRE
    
    if (!fechaOriginal || fechaOriginal === "FECHA" || fechaOriginal === "") continue;
    
    // Forzar consistencia en el formato Año-Mes (YYYY-MM)
    var mesAnio = "";
    if (fechaOriginal instanceof Date) {
      var mm = (fechaOriginal.getMonth() + 1).toString().padStart(2, '0');
      mesAnio = fechaOriginal.getFullYear() + "-" + mm;
    } else {
      mesAnio = fechaOriginal.toString().substring(0, 7).trim();
    }
    
    // Creamos una clave única por cada combinación dimensional
    var claveUnica = mesAnio + "|" + pas + "|" + cia + "|" + ramo;
    
    universoDatos[claveUnica] = {
      mesAnio: mesAnio,
      pas: pas,
      cia: cia,
      ramo: ramo,
      prima: parseFloat(fila[4]) || 0,
      premio: parseFloat(fila[5]) || 0,
      comision: parseFloat(fila[6]) || 0,
      polizas: parseFloat(fila[7]) || 0
    };
  }
  
  var clavesOrdenadas = Object.keys(universoDatos).sort();
  var resultadoFinal = [];
  
  // Encabezados con las métricas interanuales exactas que pediste
  var filaCabecera = [
    "Año-mes", "PAS", "CIA", "RAMO_NOMBRE", 
    "PRIMA", "PREMIO", "COMISION", "CANT POLIZAS",
    "VAR. MES ANTERIOR COMISION", "VAR. MES ANTERIOR POLIZAS",
    "VARIACION CANTIDAD DE POLIZAS AA", "VARIACION % CANTIDAD DE POLIZAS AA",
    "VARIACION COMISION AA", "VARIACION COMISION % AA"
  ];
  resultadoFinal.push(filaCabecera);
  
  // Motores de cálculo de fechas (Algoritmos de retroceso temporal)
  function calcularMesAnterior(mesAnioStr) {
    var partes = mesAnioStr.split("-");
    var anio = parseInt(partes[0]);
    var mes = parseInt(partes[1]) - 1;
    if (mes === 0) { mes = 12; anio -= 1; }
    return anio + "-" + mes.toString().padStart(2, '0');
  }
  
  function calcularMismoMesAnioAnterior(mesAnioStr) {
    var partes = mesAnioStr.split("-");
    var anio = parseInt(partes[0]) - 1; 
    return anio + "-" + partes[1];
  }
  
  // SEGUNDA PASADA: Cruzar datos fila por fila contra sus períodos pasados reales
  clavesOrdenadas.forEach(function(clave) {
    var actual = universoDatos[clave];
    
    // Buscar el mes anterior exacto para este PAS, CIA y Ramo
    var mesAntTarget = calcularMesAnterior(actual.mesAnio);
    var claveMesAnt = mesAntTarget + "|" + actual.pas + "|" + actual.cia + "|" + actual.ramo;
    var historicoMesAnt = universoDatos[claveMesAnt];
    
    // Buscar el mismo mes del año anterior exacto (AA)
    var anioAntTarget = calcularMismoMesAnioAnterior(actual.mesAnio);
    var claveAnioAnt = anioAntTarget + "|" + actual.pas + "|" + actual.cia + "|" + actual.ramo;
    var historicoAnioAnt = universoDatos[claveAnioAnt];
    
    // 1. Cálculos de Variación de Mes Anterior
    var varMesAntComision = historicoMesAnt ? (actual.comision - historicoMesAnt.comision) : actual.comision;
    var varMesAntPolizas = historicoMesAnt ? (actual.polizas - historicoMesAnt.polizas) : actual.polizas;
    
    // 2. Cálculos de Variación de Año Anterior (AA)
    var varAnioAntPolizas = 0;
    var varAnioAntPolizasPorc = 0;
    var varAnioAntComision = 0;
    var varAnioAntComisionPorc = 0;
    
    if (historicoAnioAnt) {
      // Diferencias absolutas reales
      varAnioAntPolizas = actual.polizas - historicoAnioAnt.polizas;
      varAnioAntComision = actual.comision - historicoAnioAnt.comision;
      
      // Diferencias porcentuales controladas contra división por cero
      varAnioAntPolizasPorc = historicoAnioAnt.polizas !== 0 ? (varAnioAntPolizas / historicoAnioAnt.polizas) : 0;
      varAnioAntComisionPorc = historicoAnioAnt.comision !== 0 ? (varAnioAntComision / historicoAnioAnt.comision) : 0;
    } else {
      // Si el año pasado no existía este ramo/compañía para este PAS, la variación es el total actual
      varAnioAntPolizas = actual.polizas;
      varAnioAntComision = actual.comision;
      varAnioAntPolizasPorc = actual.polizas !== 0 ? 1 : 0; // 100% de crecimiento
      varAnioAntComisionPorc = actual.comision !== 0 ? 1 : 0; // 100% de crecimiento
    }
    
    resultadoFinal.push([
      actual.mesAnio,
      actual.pas,
      actual.cia,
      actual.ramo,
      actual.prima,
      actual.premio,
      actual.comision,
      actual.polizas,
      varMesAntComision,
      varMesAntPolizas,
      varAnioAntPolizas,     // VARIACION CANTIDAD DE POLIZAS AA
      varAnioAntPolizasPorc,   // VARIACION % CANTIDAD DE POLIZAS AA
      varAnioAntComision,    // VARIACION COMISION AA
      varAnioAntComisionPorc   // VARIACION COMISION % AA
    ]);
  });
  
  // 4. VOLCADO FINAL EN LA PESTAÑA DE DESTINO
  var nombreNuevaHoja = "REPORTE CON VARIACIONES";
  var hojaDestino = ss.getSheetByName(nombreNuevaHoja);
  if (hojaDestino) {
    hojaDestino.clear();
  } else {
    hojaDestino = ss.insertSheet(nombreNuevaHoja);
  }
  
  hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
  
  // Estilos de diseño para legibilidad inmediata
  hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).setBackground("#f3f3f3").setFontWeight("bold");
  
  if (resultadoFinal.length > 1) {
    var maxFilas = resultadoFinal.length - 1;
    
    // Formato de Moneda ($) para comisiones y variaciones monetarias
    hojaDestino.getRange(2, 5, maxFilas, 3).setNumberFormat("$#,##0");  // PRIMA, PREMIO, COMISION
    hojaDestino.getRange(2, 9, maxFilas, 1).setNumberFormat("$#,##0");  // VAR MES ANT COMISION
    hojaDestino.getRange(2, 13, maxFilas, 1).setNumberFormat("$#,##0"); // VARIACION COMISION AA
    
    // Formato Numérico Entero (#) para conteos fijos
    hojaDestino.getRange(2, 8, maxFilas, 1).setNumberFormat("#,##0");   // CANT POLIZAS
    hojaDestino.getRange(2, 10, maxFilas, 2).setNumberFormat("#,##0");  // VAR MES ANT POL y VAR POLIZAS AA
    
    // Formato de Porcentaje Profesional (0.00%) para análisis de crecimiento
    hojaDestino.getRange(2, 12, maxFilas, 1).setNumberFormat("0.00%");   // VARIACION % CANTIDAD DE POLIZAS AA
    hojaDestino.getRange(2, 14, maxFilas, 1).setNumberFormat("0.00%");   // VARIACION COMISION % AA
  }
  
  hojaDestino.autoResizeColumns(1, hojaDestino.getLastColumn());
  console.log("¡Reporte de variaciones interanuales ejecutado con éxito y números corregidos!");
}