/**
 * PANEL DE CONTROL MASTER
 * Interfaz de gestión para el manager del sistema de tours
 */

class PanelControlMaster {
  
  // Crear menú principal
  static crearMenu() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🗓️ SHERPAS')
      .addItem('👤 Crear Nuevo Guía', 'PanelControlMaster.crearNuevoGuia')
      .addItem('🗑️ Eliminar Guía', 'PanelControlMaster.eliminarGuia')
      .addItem('📋 Lista de Guías', 'PanelControlMaster.mostrarListaGuias')
      .addSeparator()
      .addItem('📊 Generar Master Calendar', 'PanelControlMaster.generarMasterCalendar')
      .addItem('🔄 Sincronizar Guías', 'ServicioSincronizacion.sincronizarTodosLosGuias')
      .addSeparator()
      .addItem('🔗 Instalar Triggers', 'ServicioSincronizacion.instalarTriggers')
      .addItem('📋 Estado del Sistema', 'PanelControlMaster.mostrarEstadoSistema')
      .addSeparator()
      .addItem('⚙️ Configuración', 'PanelControlMaster.mostrarPanelConfiguracion')
      .addItem('📧 Test Email', 'PanelControlMaster.testearEmail')
      .addToUi();
  }

  // Crear nuevo guía
  static crearNuevoGuia() {
    const ui = SpreadsheetApp.getUi();
    
    const codigo = ui.prompt('👤 Crear Nuevo Guía', 'Código (G01, G02...):', ui.ButtonSet.OK_CANCEL);
    if (codigo.getSelectedButton() !== ui.Button.OK || !codigo.getResponseText().trim()) return;
    
    const codigoLimpio = codigo.getResponseText().trim().toUpperCase();
    if (!ConfiguracionSistema.validarCodigoGuia(codigoLimpio)) {
      ui.alert('❌ Error', 'Código inválido. Formato: G01, G02...', ui.ButtonSet.OK);
      return;
    }
    
    const guiasExistentes = ConfiguracionSistema.getGuiasConfigurados();
    if (guiasExistentes.find(g => g.codigo === codigoLimpio)) {
      ui.alert('❌ Error', `Código ${codigoLimpio} ya existe.`, ui.ButtonSet.OK);
      return;
    }
    
    const nombre = ui.prompt('👤 Crear Nuevo Guía', 'Nombre del guía:', ui.ButtonSet.OK_CANCEL);
    if (nombre.getSelectedButton() !== ui.Button.OK || !nombre.getResponseText().trim()) return;
    
    const email = ui.prompt('👤 Crear Nuevo Guía', 'Email del guía:', ui.ButtonSet.OK_CANCEL);
    if (email.getSelectedButton() !== ui.Button.OK || !email.getResponseText().trim()) return;
    
    const emailLimpio = email.getResponseText().trim().toLowerCase();
    if (!ConfiguracionSistema.validarEmail(emailLimpio)) {
      ui.alert('❌ Error', 'Email inválido', ui.ButtonSet.OK);
      return;
    }
    
    try {
      const sheetId = ServicioSincronizacion.crearCalendarioGuia(codigoLimpio, nombre.getResponseText().trim(), emailLimpio);
      ui.alert('✅ Guía Creado', `${codigoLimpio} - ${nombre.getResponseText().trim()} creado exitosamente.`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('❌ Error', `Error: ${error.message}`, ui.ButtonSet.OK);
    }
  }

  // Eliminar guía
  static eliminarGuia() {
    const ui = SpreadsheetApp.getUi();
    const guias = ConfiguracionSistema.getGuiasConfigurados();
    
    if (guias.length === 0) {
      ui.alert('❌ Error', 'No hay guías para eliminar.', ui.ButtonSet.OK);
      return;
    }
    
    let lista = 'Guías:\n';
    guias.forEach(g => lista += `• ${g.codigo} - ${g.nombre}\n`);
    
    const codigo = ui.prompt('🗑️ Eliminar Guía', `${lista}\nCódigo a eliminar:`, ui.ButtonSet.OK_CANCEL);
    if (codigo.getSelectedButton() !== ui.Button.OK) return;
    
    const confirmacion = ui.alert('⚠️ Confirmar', `¿Eliminar ${codigo.getResponseText()}?`, ui.ButtonSet.YES_NO);
    if (confirmacion === ui.Button.YES) {
      ServicioSincronizacion.eliminarCalendarioGuia(codigo.getResponseText().trim().toUpperCase());
    }
  }

  // Lista de guías
  static mostrarListaGuias() {
    const guias = ConfiguracionSistema.getGuiasConfigurados();
    const mensaje = guias.length === 0 
      ? 'No hay guías configurados.' 
      : `Guías (${guias.length}):\n\n${guias.map(g => `• ${g.codigo} - ${g.nombre} (${g.email})`).join('\n')}`;
    
    SpreadsheetApp.getUi().alert('📋 Lista de Guías', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  // Generar master calendar
  static generarMasterCalendar() {
    const guias = ConfiguracionSistema.getGuiasConfigurados();
    if (guias.length === 0) {
      SpreadsheetApp.getUi().alert('❌ Error', 'No hay guías. Crea guías primero.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    try {
      const masterCalendar = new MasterCalendar();
      ConfiguracionSistema.MESES_ACTIVOS.forEach(mes => {
        masterCalendar.crearPestanaMes(mes, guias);
      });
      
      SpreadsheetApp.getUi().alert('✅ Master Generado', `Master calendar creado con ${guias.length} guías.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (error) {
      SpreadsheetApp.getUi().alert('❌ Error', `Error: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }

  // Estado del sistema
  static mostrarEstadoSistema() {
    const guias = ConfiguracionSistema.getGuiasConfigurados();
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onEditMasterCalendar');
    
    const mensaje = `📊 ESTADO:

👥 Guías: ${guias.length}
📅 Pestañas: ${sheets.length}
🔗 Triggers: ${triggers.length}

${guias.length > 0 ? '✅' : '❌'} Guías configurados
${sheets.length > 1 ? '✅' : '❌'} Master generado
${triggers.length > 0 ? '✅' : '❌'} Triggers activos`;
    
    SpreadsheetApp.getUi().alert('📊 Estado', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  // Configuración
  static mostrarPanelConfiguracion() {
    const config = ConfiguracionSistema.obtenerConfiguracionCompleta();
    const mensaje = `⚙️ CONFIGURACIÓN:

Master ID: ${config.masterSheetId}
Carpeta: ${config.folderGuias}
Meses: ${config.mesesActivos.join(', ')}
Guías: ${config.guias.length}`;
    
    SpreadsheetApp.getUi().alert('⚙️ Config', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  // Test email
  static testearEmail() {
    const ui = SpreadsheetApp.getUi();
    const email = ui.prompt('📧 Test Email', 'Email para prueba:', ui.ButtonSet.OK_CANCEL);
    
    if (email.getSelectedButton() !== ui.Button.OK) return;
    
    try {
      const exito = ServicioEmail.enviarEmailTest(email.getResponseText().trim());
      ui.alert(exito ? '✅ Email enviado' : '❌ Error enviando', '', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
    }
  }
}

// Funciones globales requeridas
function onOpen() {
  PanelControlMaster.crearMenu();
}

function onEditMasterCalendar(e) {
  ServicioSincronizacion.onEditMasterCalendar(e);
}