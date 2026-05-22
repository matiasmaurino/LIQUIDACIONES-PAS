function consolidarYAgruparSeguros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaDestino = "CLIENTES AGRUPADOS";
  
  let mapaClientes = {};

  // CONFIGURACIÓN DE HOJAS
  const configuracionHojas = [
    { nombre: "RIV", colNombre: 2, colMatricula: 8, colDir: 9, colTel: 10, colMail: 11, colCuil: 12, etiqueta: "RIV" },        
    { nombre: "PS", colNombre: 2, colMatricula: 8, colDir: 9, colTel: 10, colMail: 11, colCuil: 12, etiqueta: "PS" },         
    { nombre: "CLIENTES FP", colNombre: 2, colMatricula: 0, colDir: 10, colTel: 15, colMail: 17, colCuil: 16, etiqueta: "FP" } 
  ];

  configuracionHojas.forEach(conf => {
    const hoja = ss.getSheetByName(conf.nombre);
    if (!hoja) return;

    const datos = hoja.getDataRange().getValues();
    datos.shift(); 

    datos.forEach(fila => {
      let nombreRaw = fila[conf.colNombre];
      const matricula = fila[conf.colMatricula] ? fila[conf.colMatricula].toString().trim() : "";
      const direccion = fila[conf.colDir] ? fila[conf.colDir].toString().trim() : "";
      const telefono = fila[conf.colTel] ? fila[conf.colTel].toString().trim() : "";
      const email = fila[conf.colMail] ? fila[conf.colMail].toString().trim() : "";
      const cuil = fila[conf.colCuil] ? fila[conf.colCuil].toString().trim() : "";

      if (nombreRaw === "" || typeof nombreRaw === 'number' || !isNaN(nombreRaw)) {
        return; 
      }

      let nombre = nombreRaw.toString().replace(/,/g, "").trim().toUpperCase();

      if (nombre !== "") {
        if (!mapaClientes[nombre]) {
          mapaClientes[nombre] = { 
            riv: "", ps: "", fp: "", 
            direccion: direccion, 
            telefono: telefono, 
            email: email, 
            cuil: cuil 
          };
        }

        if (!mapaClientes[nombre].direccion) mapaClientes[nombre].direccion = direccion;
        if (!mapaClientes[nombre].telefono) mapaClientes[nombre].telefono = telefono;
        if (!mapaClientes[nombre].email) mapaClientes[nombre].email = email;
        if (!mapaClientes[nombre].cuil) mapaClientes[nombre].cuil = cuil;

        if (conf.etiqueta === "RIV") mapaClientes[nombre].riv = matricula;
        if (conf.etiqueta === "PS")  mapaClientes[nombre].ps = matricula;
        if (conf.etiqueta === "FP")  mapaClientes[nombre].fp = matricula;
      }
    });
  });

  // NUEVO: El encabezado ahora incluye ID_CLIENTE al principio
  let resultadoFinal = [["ID_CLIENTE", "CLIENTE", "RIVADAVIA", "PROVINCIA", "FED. PATRONAL", "DIRECCION", "TELEFONO", "EMAIL", "CUIL"]];
  const nombresOrdenados = Object.keys(mapaClientes).sort();

  // Generador de ID correlativo
  let contadorID = 1;

  nombresOrdenados.forEach(nombre => {
    resultadoFinal.push([
      contadorID, // Nueva columna A: ID numérico
      nombre,     // Columna B: Nombre
      mapaClientes[nombre].riv,
      mapaClientes[nombre].ps,
      mapaClientes[nombre].fp,
      mapaClientes[nombre].direccion,
      mapaClientes[nombre].telefono,
      mapaClientes[nombre].email,
      mapaClientes[nombre].cuil
    ]);
    contadorID++;
  });

  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(nombreHojaDestino);
  } else {
    hojaDestino.clear();
  }

  if (resultadoFinal.length > 1) {
    hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
    hojaDestino.autoResizeColumns(1, resultadoFinal[0].length);
    hojaDestino.getRange(1, 1, 1, resultadoFinal[0].length).setFontWeight("bold").setBackground("#f3f3f3");
    
    console.log("Consolidación terminada con IDs generados.");
  }

// ... (todo tu código anterior de pegado de datos e IDs se mantiene igual)

  if (resultadoFinal.length > 1) {
    hojaDestino.getRange(1, 1, resultadoFinal.length, resultadoFinal[0].length).setValues(resultadoFinal);
    hojaDestino.autoResizeColumns(1, resultadoFinal[0].length);
    hojaDestino.getRange(1, 1, 1, resultadoFinal[0].length).setFontWeight("bold").setBackground("#f3f3f3");
    
    console.log("Consolidación terminada con IDs generados.");
  }

  // =========================================================
  // EL REMATE AUTOMÁTICO (DEBE IR ACÁ, ADENTRO DE LA FUNCIÓN)
  // =========================================================
  console.log("Iniciando generación de resumen de comisiones mensuales...");
  generarResumenMensualConPorcentajes(); 
  console.log("¡Todo el sistema se actualizó correctamente!");

} // <-- ESTA LLAVE CIERRA LA FUNCIÓN PRINCIPAL 'consolidarYAgruparSeguros'