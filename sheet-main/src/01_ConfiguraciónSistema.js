/**
 * CONFIGURACIÓN DEL SISTEMA DE TOURS
 * Configuración centralizada y constantes del sistema
 */

class ConfiguracionSistema {
  
  // Configuración dinámica de guías desde PropertiesService
  static getGuiasConfigurados() {
    const props = PropertiesService.getScriptProperties();
    const guiasStr = props.getProperty('GUIAS_CONFIGURADOS');
    return guiasStr ? JSON.parse(guiasStr) : [];
  }

  // Obtener guía por sheet ID
  static obtenerGuiaPorSheetId(sheetId) {
    return this.getGuiasConfigurados().find(g => g.sheetId === sheetId);
  }

  // Obtener guía por código
  static obtenerGuiaPorCodigo(codigo) {
    return this.getGuiasConfigurados().find(g => g.codigo === codigo);
  }

  // Master sheet ID dinámico
  static get MASTER_SHEET_ID() {
    const props = PropertiesService.getScriptProperties();
    const masterId = props.getProperty('MASTER_SHEET_ID');
    return masterId || SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  // Carpeta para calendarios de guías
  static get FOLDER_GUIAS() {
    return 'Calendarios_Guias';
  }

  // ID de carpeta padre donde crear subcarpetas
  static get FOLDER_PADRE_ID() {
    return '1zUSzaQrXdLRazr9uCH8A5TeNkcgOzwsP'; 
  }

  // Meses activos del sistema
  static get MESES_ACTIVOS() {
    return ['2025-10', '2025-11', '2025-12'];
  }

  // Opciones de desplegables para guías
  static get OPCIONES_GUIA() {
    return ['', 'NO DISPONIBLE', 'REVERTIR'];
  }

  // Opciones para master - mañana
  static get OPCIONES_MASTER_MANANA() {
    return ['', 'LIBERAR', 'ASIGNAR M'];
  }

  // Opciones para master - tarde
  static get OPCIONES_MASTER_TARDE() {
    return ['', 'LIBERAR', 'ASIGNAR T1', 'ASIGNAR T2', 'ASIGNAR T3'];
  }

  // Email del manager
  static get EMAIL_MANAGER() {
    return 'manager@tours.com';
  }

  // Configuración de colores
  static get COLORES() {
    return {
      NO_DISPONIBLE: '#ff0000',
      ASIGNADO: '#00ff00',
      LIBRE: '#ffffff'
    };
  }

  // Días de la semana en español
  static get DIAS_SEMANA() {
    return ['DOM', 'LUN', 'MAR', 'MIE', 'JUE', 'VIE', 'SAB'];
  }

  // Meses en español
  static get MESES_NOMBRES() {
    return {
      '2025-10': 'Octubre 2025',
      '2025-11': 'Noviembre 2025', 
      '2025-12': 'Diciembre 2025'
    };
  }

  // Validar email
  static validarEmail(email) {
    const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return regex.test(email);
  }

  // Validar código de guía
  static validarCodigoGuia(codigo) {
    const regex = /^[A-Z]\d{2}$/;
    return regex.test(codigo);
  }

  // Obtener configuración completa del sistema
  static obtenerConfiguracionCompleta() {
    return {
      guias: this.getGuiasConfigurados(),
      masterSheetId: this.MASTER_SHEET_ID,
      mesesActivos: this.MESES_ACTIVOS,
      folderGuias: this.FOLDER_GUIAS,
      emailManager: this.EMAIL_MANAGER
    };
  }

  // Verificar integridad del sistema
  static verificarIntegridad() {
    const issues = [];
    
    const guias = this.getGuiasConfigurados();
    if (guias.length === 0) {
      issues.push('No hay guías configurados');
    }

    guias.forEach(guia => {
      try {
        SpreadsheetApp.openById(guia.sheetId);
      } catch (e) {
        issues.push(`No se puede acceder al calendario de ${guia.codigo}`);
      }
    });

    return {
      valido: issues.length === 0,
      errores: issues
    };
  }
}