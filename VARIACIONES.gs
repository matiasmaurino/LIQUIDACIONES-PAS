function generarTablaConVariaciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. OBTENER DATOS DE ORIGEN
  var hojaOrigen = ss.getSheetByName("LIQUIDACIONES AGRUPADAS");
  if (!hojaOrigen) {
    SpreadsheetApp.getUi().alert("No se encontró la pestaña 'LIQUIDACIONES AGRUPADAS'");
    return;
  }
  
  var datosBase = hojaOrigen.getDataRange().getValues();
  
  // --- PASO A PASO EN MEMORIA ---
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
    
    if (!pasAgrupado || !fechaOriginal || fechaOriginal === "FECHA") continue;
    
    // Formatear Fecha de manera segura
    var mesAnio = "";
    if (fechaOriginal instanceof Date) {
      var mm = (fechaOriginal.getMonth() + 1).toString().padStart(2, '0');
      mesAnio = fechaOriginal.getFullYear() + "-" + mm;
    } else {
      var partes = fechaOriginal.toString().split(/[-/]/);
      if (partes.length >= 3) {
        if (partes[0].length === 4) {
          mesAnio = partes[0] + "-" + partes[1].padStart(2, '0');
        } else {
          mesAnio = partes[2] + "-" + partes[1].padStart(2, '0');
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
  
  // 2. CONSTRUIR LA NUEVA TABLA CON VARIACIONES
  var resultadoFinal = [];
  
  // --- CABECERA EN UNA SOLA FILA ---
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
  
  // Llenar filas de datos y calcular variaciones respecto al mes anterior
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
  
  // 3. ESCRIBIR EN LA PESTAÑA
  var nombreNuevaHoja = "REPORTE CON VARIACIONES";
  var hojaDestino = ss.getSheetByName(nombreNuevaHoja);
  if (hojaDestino) {
    hojaDestino.clear(); 
  } else {
    hojaDestino = ss.insertSheet(nombreNuevaHoja);
  }
  
  hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
  
  // Formato visual rápido (Fila 1 en negrita y fondo gris)
  hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).setBackground("#f3f3f3").setFontWeight("bold");
  hojaDestino.autoResizeColumns(1, hojaDestino.getLastColumn());
  
  SpreadsheetApp.getUi().alert("¡Reporte actualizado con éxito en una sola línea de títulos!");
}