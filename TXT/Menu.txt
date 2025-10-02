/**
 * MENÚ PERSONALIZADO Y TRIGGERS
 * Interfaz de usuario y automatización
 */

/**
 * Crea el menú personalizado al abrir la hoja
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Sincronizador')
    .addItem('🚀 Inicializar Sistema', 'inicializarSistema')
    .addSeparator()
    .addItem('▶️ Ejecutar Sincronización Manual', 'ejecutarSincronizacionManual')
    .addSeparator()
    .addItem('➕ Crear Nuevo Guía', 'crearNuevoGuiaUI')
    .addItem('🔧 Configurar Calendar ID', 'configurarCalendarIdUI')
    .addSeparator()
    .addItem('⏰ Instalar Trigger Automático', 'instalarTriggerAutomatico')
    .addItem('🗑️ Eliminar Triggers', 'eliminarTodosLosTriggers')
    .addToUi();
}

/**
 * Ejecuta sincronización manual desde el menú
 */
function ejecutarSincronizacionManual() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Sincronización Manual',
    '¿Desea ejecutar la sincronización ahora?',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta === ui.Button.YES) {
    try {
      ejecutarSincronizacion();
      ui.alert('Sincronización completada exitosamente');
    } catch (error) {
      ui.alert('Error en sincronización: ' + error.toString());
    }
  }
}

/**
 * Instala el trigger automático de tiempo
 */
function instalarTriggerAutomatico() {
  const ui = SpreadsheetApp.getUi();
  
  // Verificar si ya existe un trigger
  const triggers = ScriptApp.getProjectTriggers();
  const triggerExistente = triggers.find(t => 
    t.getHandlerFunction() === CONFIG.TRIGGER_FUNCTION
  );
  
  if (triggerExistente) {
    const respuesta = ui.alert(
      'Trigger Existente',
      'Ya existe un trigger automático. ¿Desea reemplazarlo?',
      ui.ButtonSet.YES_NO
    );
    
    if (respuesta === ui.Button.NO) {
      return;
    }
    
    ScriptApp.deleteTrigger(triggerExistente);
  }
  
  // Crear nuevo trigger
  ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTION)
    .timeBased()
    .everyMinutes(CONFIG.TRIGGER_INTERVAL_MINUTES)
    .create();
  
  ui.alert(
    '✅ Trigger Instalado',
    `La sincronización se ejecutará automáticamente cada ${CONFIG.TRIGGER_INTERVAL_MINUTES} minuto(s).`,
    ui.ButtonSet.OK
  );
}

/**
 * Elimina todos los triggers del proyecto
 */
function eliminarTodosLosTriggers() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Eliminar Triggers',
    '¿Está seguro que desea eliminar todos los triggers automáticos?',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    ui.alert('Todos los triggers han sido eliminados');
  }
}

/**
 * Configura el Calendar ID del maestro
 */
function configurarCalendarIdUI() {
  const ui = SpreadsheetApp.getUi();
  
  const calendarIdActual = obtenerCalendarIdMaestro() || '(No configurado)';
  
  const respuesta = ui.prompt(
    'Configurar Calendar ID',
    `Calendar ID actual: ${calendarIdActual}\n\nIngrese el nuevo Calendar ID:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuesta.getSelectedButton() === ui.Button.OK) {
    const nuevoId = respuesta.getResponseText().trim();
    
    if (nuevoId) {
      establecerCalendarIdMaestro(nuevoId);
      ui.alert('Calendar ID configurado correctamente');
    } else {
      ui.alert('Debe ingresar un Calendar ID válido');
    }
  }
}