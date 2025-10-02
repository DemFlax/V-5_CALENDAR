/**
 * GESTIÓN DEL REGISTRO DE GUÍAS
 * Lectura y escritura de la tabla REGISTRO
 */

/**
 * Obtiene todos los guías activos del registro
 * @returns {Array<ClaseGuia>}
 */
function obtenerGuiasDelRegistro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistro = ss.getSheetByName(CONFIG.SHEET_REGISTRO);
  
  if (!sheetRegistro) {
    throw new Error(`Pestaña ${CONFIG.SHEET_REGISTRO} no encontrada`);
  }
  
  const data = sheetRegistro.getDataRange().getValues();
  const guias = [];
  
  // Saltar fila de encabezados (índice 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const codigo = row[CONFIG.REGISTRO_COL.CODIGO];
    const nombre = row[CONFIG.REGISTRO_COL.NOMBRE];
    const email = row[CONFIG.REGISTRO_COL.EMAIL];
    const fileId = row[CONFIG.REGISTRO_COL.FILE_ID];
    
    // Validar que tenga los datos mínimos
    if (codigo && nombre && fileId) {
      guias.push(new ClaseGuia(codigo, nombre, email, fileId));
    }
  }
  
  return guias;
}

/**
 * Agrega un nuevo guía al registro
 */
function agregarGuiaAlRegistro(codigo, nombre, email, fileId, url) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistro = ss.getSheetByName(CONFIG.SHEET_REGISTRO);
  
  const timestamp = new Date();
  const nuevaFila = [timestamp, codigo, nombre, email, fileId, url];
  
  sheetRegistro.appendRow(nuevaFila);
}

/**
 * Verifica si un código de guía ya existe
 */
function codigoGuiaExiste(codigo) {
  const guias = obtenerGuiasDelRegistro();
  return guias.some(g => g.codigo === codigo);
}

/**
 * Genera el siguiente código de guía disponible
 */
function generarSiguienteCodigoGuia() {
  const guias = obtenerGuiasDelRegistro();
  
  if (guias.length === 0) {
    return 'G01';
  }
  
  // Extraer números de los códigos existentes
  const numeros = guias.map(g => {
    const match = g.codigo.match(/G(\d+)/);
    return match ? parseInt(match[1]) : 0;
  });
  
  const maxNumero = Math.max(...numeros);
  const siguienteNumero = maxNumero + 1;
  
  return 'G' + siguienteNumero.toString().padStart(2, '0');
}