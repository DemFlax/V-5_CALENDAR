/**
 * INICIALIZACIÓN DEL SISTEMA
 * Configuración inicial completa
 */

/**
 * Inicializa el sistema completo desde cero
 */
function inicializarSistema() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const confirmacion = ui.alert(
    'Inicializar Sistema',
    '¿Desea inicializar el sistema?\n\n' +
    'Esto creará:\n' +
    '- Pestaña REGISTRO\n' +
    '- Hojas mensuales (próximos 3 meses)\n' +
    '- Estructura base del Master\n\n' +
    '⚠️ Si ya existen, serán reemplazadas.',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmacion !== ui.Button.YES) {
    return;
  }
  
  try {
    // 1. Crear/Configurar REGISTRO
    crearPestanaRegistro(ss);
    
    // 2. Crear hojas mensuales del Master
    crearHojasMensualesMaster(ss);
    
    ui.alert(
      '✅ Sistema Inicializado',
      'El sistema ha sido configurado correctamente.\n\n' +
      'Próximos pasos:\n' +
      '1. Configurar Calendar ID desde el menú\n' +
      '2. Crear guías\n' +
      '3. Instalar trigger automático',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error en inicialización: ' + error.toString());
  }
}

/**
 * Crea o reconfigura la pestaña REGISTRO
 */
function crearPestanaRegistro(ss) {
  let sheetRegistro = ss.getSheetByName(CONFIG.SHEET_REGISTRO);
  
  // Si existe, limpiar
  if (sheetRegistro) {
    sheetRegistro.clear();
  } else {
    sheetRegistro = ss.insertSheet(CONFIG.SHEET_REGISTRO);
  }
  
  // Encabezados
  const encabezados = ['TIMESTAMP', 'CODIGO', 'NOMBRE', 'EMAIL', 'FILE_ID', 'URL'];
  sheetRegistro.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  sheetRegistro.getRange(1, 1, 1, encabezados.length).setFontWeight('bold');
  sheetRegistro.getRange(1, 1, 1, encabezados.length).setBackground('#4285F4');
  sheetRegistro.getRange(1, 1, 1, encabezados.length).setFontColor('#FFFFFF');
  
  // Ajustar anchos
  sheetRegistro.setColumnWidth(1, 150); // TIMESTAMP
  sheetRegistro.setColumnWidth(2, 80);  // CODIGO
  sheetRegistro.setColumnWidth(3, 150); // NOMBRE
  sheetRegistro.setColumnWidth(4, 200); // EMAIL
  sheetRegistro.setColumnWidth(5, 250); // FILE_ID
  sheetRegistro.setColumnWidth(6, 300); // URL
  
  // Congelar fila de encabezados
  sheetRegistro.setFrozenRows(1);
  
  Logger.log('Pestaña REGISTRO creada');
}

/**
 * Crea hojas mensuales para el Master
 */
function crearHojasMensualesMaster(ss) {
  const hoy = new Date();
  
  // Crear hojas para los próximos 3 meses
  for (let i = 0; i < 3; i++) {
    const fecha = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
    const mes = fecha.getMonth() + 1;
    const anio = fecha.getFullYear();
    
    crearHojaMensualMaster(ss, mes, anio);
  }
}

/**
 * Crea una hoja mensual individual para el Master
 */
function crearHojaMensualMaster(ss, mes, anio) {
  const nombreHoja = `${mes}_${anio}`;
  
  // Eliminar si existe
  let sheet = ss.getSheetByName(nombreHoja);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  sheet = ss.insertSheet(nombreHoja);
  
  // Encabezados iniciales
  sheet.getRange(1, 1).setValue('FECHA');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(1, 1).setHorizontalAlignment('center');
  sheet.getRange(1, 1).setBackground('#4285F4');
  sheet.getRange(1, 1).setFontColor('#FFFFFF');
  
  // Fila 2 vacía para MAÑANA/TARDE (se llenará al crear guías)
  
  // Generar todas las fechas del mes
  const fechas = generarFechasMes(mes, anio);
  const filaInicio = 3; // Fila 1=FECHA, Fila 2=MAÑANA/TARDE, Fila 3+=datos
  
  for (let i = 0; i < fechas.length; i++) {
    const fila = filaInicio + i;
    sheet.getRange(fila, 1).setValue(fechas[i]);
    sheet.getRange(fila, 1).setNumberFormat('dd/mm/yyyy');
  }
  
  // Congelar columna A y filas 1-2
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(2);
  
  // Ajustar ancho columna A
  sheet.setColumnWidth(1, 120);
  
  Logger.log(`Hoja mensual ${nombreHoja} creada con ${fechas.length} fechas`);
}