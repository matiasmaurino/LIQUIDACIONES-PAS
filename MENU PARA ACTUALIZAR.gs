function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💼 Gestión Integral de clientes y liquidaciones')
    .addItem('📥 Paso 1: Procesar Liquidación RIV', 'ejecutarPaso1_RIV')
    .addItem('📥 Paso 2: Procesar Liquidación PS', 'ejecutarPaso2_PS')
    .addItem('📥 Paso 3: Procesar Liquidación FP', 'ejecutarPaso3_FP')
    .addSeparator()
    .addItem('👥 Paso 4: Agrupar Clientes y Asignar PAS', 'ejecutarPaso4_Agrupacion')
    .addSeparator()
    .addItem('📊 Paso 5: Actualizar Resumen Mensual', 'ejecutarPaso5_Resumen')
    .addItem('📉 Paso 6: Actualizar Tabla de Variaciones', 'ejecutarPaso6_Variaciones')
    .addItem('🚀 Paso 7: Sincronizar Detalle Looker Studio', 'ejecutarPaso7_Looker')
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ Pasos Individuales')
       .addItem('Importar Clientes FP base (CSV)', 'importarCSVsDesdeCarpetas'))
    .addToUi();
}

/**
 * PASO 1: Procesamiento exclusivo de Rivadavia
 */
function ejecutarPaso1_RIV() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando extracción RIV...");
    consolidarYLimpiarRIV();
    ui.alert('✅ Paso 1 Completado', 'La liquidación de RIV se procesó correctamente.\n\nContinuá con el "Paso 2".', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 1 (RIV)', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * PASO 2: Procesamiento exclusivo de Pronto Pago / PS
 */
function ejecutarPaso2_PS() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando extracción PS...");
    consolidarYLimpiarPS();
    ui.alert('✅ Paso 2 Completado', 'La liquidación de PS se procesó correctamente.\n\nContinuá con el "Paso 3".', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 2 (PS)', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * PASO 3: Procesamiento exclusivo de Federación Patronal
 */
function ejecutarPaso3_FP() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando extracción FP...");
    consolidarYLimpiarFP();
    ui.alert('✅ Paso 3 Completado', 'La liquidación de FP se procesó correctamente.\n\nContinuá con el "Paso 4".', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 3 (FP)', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * PASO 4: Cruce de clientes e identificación de IDs
 */
function ejecutarPaso4_Agrupacion() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando unificación de clientes...");
    consolidarYAgruparSeguros();
    ui.alert('✅ Paso 4 Completado', 'Se cruzaron los clientes y se asignaron las matrículas/PAS.\n\nYa podés correr el "Paso 5".', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 4 (Agrupación)', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * PASO 5: Actualizar Resumen Mensual
 */
function ejecutarPaso5_Resumen() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando Resumen Mensual...");
    consolidarResumenMensual();
    ui.alert('✅ Paso 5 Completado', 'El Resumen Mensual Acumulado fue actualizado con éxito.\n\nContinuá con el "Paso 6".', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 5 (Resumen)', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * PASO 6: Calcular Tabla de Variaciones
 */
function ejecutarPaso6_Variaciones() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando Tabla con Variaciones...");
    generarTablaConVariaciones();
    ui.alert('✅ Paso 6 Completado', 'El Reporte con Variaciones y desvíos fue recalculado con éxito.\n\nFinalizá ejecutando el "Paso 7".', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 6 (Variaciones)', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * PASO 7: Sincronizar con el Histórico de Looker Studio
 */
function ejecutarPaso7_Looker() {
  const ui = SpreadsheetApp.getUi();
  try {
    console.log("Iniciando Indexación para Looker Studio...");
    actualizarHistoricoDetalleLooker();
    ui.alert('🎉 ¡Proceso Totalmente Finalizado!', 'La base transaccional detallada de Looker Studio se actualizó e indexó de forma incremental sin errores.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error en Paso 7 (Looker)', error.toString(), ui.ButtonSet.OK);
  }
}