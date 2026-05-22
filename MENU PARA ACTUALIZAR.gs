function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💼 Gestión Integral de clientes y liquidaciones')
    .addItem('🚀 Ejecutar Importación Completa', 'ejecutarProcesoCompletoLiquidaciones')
    .addSeparator()
    .addItem('📊 Actualizar Resumen Mensual (Acumulado)', 'consolidarResumenMensual')
    .addItem('📉 Actualizar Reporte de Variaciones', 'ejecutarVariacionesManualmente')
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
 * Ejecuta el flujo completo automatizado de liquidaciones.
 * Al finalizar, consolida de forma inteligente el resumen mensual.
 */
function ejecutarProcesoCompletoLiquidaciones() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Ejecutar procesos de consolidación individuales por compañía
    console.log("Iniciando paso 1: Consolidación RIV...");
    consolidarYLimpiarRIV();
    
    console.log("Iniciando paso 2: Consolidación PS...");
    consolidarYLimpiarPS();
    
    console.log("Iniciando paso 3: Consolidación FP...");
    consolidarYLimpiarFP();
    
    // 2. Agrupar la información cruzando clientes e IDs de PAS
    console.log("Iniciando paso 4: Agrupación de seguros y asignación de PAS...");
    consolidarYAgruparSeguros();
    
    // 3. NUEVO LLAMADO: Construir el resumen mensual con las novedades del mes sin pisar el histórico
    console.log("Iniciando paso 5: Consolidación incremental del Resumen Mensual...");
    consolidarResumenMensual();
    
    ui.alert('✅ Proceso Completado', 'Se ejecutaron todas las importaciones, agrupaciones y el resumen mensual de forma exitosa.', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error("Error en la ejecución del proceso completo: " + error.toString());
    ui.alert('❌ Error en el Proceso', 'Ocurrió un error durante la ejecución:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Función auxiliar para permitirle al usuario actualizar el reporte de variaciones 
 * de forma manual directamente desde el menú principal.
 */
function ejecutarVariacionesManualmente() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Aquí asumimos que en VARIACIONES.gs tu función principal se llama generarReporteVariaciones o similar
    if (typeof generarReporteVariaciones === 'function') {
      generarReporteVariaciones();
      ui.alert('✅ Reporte Actualizado', 'El Reporte con Variaciones se recalculó correctamente.', ui.ButtonSet.OK);
    } else {
      ui.alert('⚠️ Advertencia', 'No se encontró una función llamada "generarReporteVariaciones" en tus archivos.', ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('❌ Error', 'No se pudo actualizar el reporte de variaciones: ' + error.toString(), ui.ButtonSet.OK);
  }
}