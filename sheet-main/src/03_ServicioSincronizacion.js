/**
 * SERVICIO DE SINCRONIZACIÓN
 * Maneja la sincronización entre calendarios y triggers + Template System
 */

class ServicioSincronizacion {
  
  // Trigger principal para cambios en calendarios de guías
  static onEditCalendarioGuia(e) {
    try {
      const sheet = e.source.getActiveSheet();
      const range = e.range;
      const sheetId = e.source.getId();
      
      // Identificar guía
      const guia = ConfiguracionSistema.obtenerGuiaPorSheetId(sheetId);
      if (!guia) return;
      
      // Procesar cambio
      const cambio = this.analizarCambio(range, sheet);
      if (!cambio) return;
      
      this.procesarCambioGuia(guia, cambio, range.getValue());
      
    } catch (error) {
      console.error('Error en onEditCalendarioGuia:', error);
    }
  }
  
  // Trigger para cambios en master calendar
  static onEditMasterCalendar(e) {
    try {
      const range = e.range;
      const sheet = e.source.getActiveSheet();
      
      const cambio = this.analizarCambioMaster(range, sheet);
      if (!cambio) return;
      
      this.procesarCambioMaster(cambio, range.getValue());
      
    } catch (error) {
      console.error('Error en onEditMasterCalendar:', error);
    }
  }
  
  // Analizar cambio en calendario de guía
  static analizarCambio(range, sheet) {
    const fila = range.getRow();
    const columna = range.getColumn();
    
    // Verificar si es una celda de turno válida
    if (!this.esCeldaTurno(range, sheet)) return null;
    
    const fecha = this.extraerFechaDeCelda(range, sheet);
    const turno = this.extraerTurnoDeCelda(range, sheet);
    
    if (!fecha || !turno) return null;
    
    return {
      fecha: fecha,
      turno: turno,
      fila: fila,
      columna: columna
    };
  }
  
  // Analizar cambio en master calendar
  static analizarCambioMaster(range, sheet) {
    const fila = range.getRow();
    const columna = range.getColumn();
    
    // Obtener fecha de la fila
    const fechaCelda = sheet.getRange(fila, 1);
    const fecha = fechaCelda.getValue();
    
    if (!(fecha instanceof Date)) return null;
    
    // Determinar si es columna de mañana o tarde
    const turno = columna === 2 ? 'MAÑANA' : (columna === 3 ? 'TARDE' : null);
    if (!turno) return null;
    
    // Extraer código de guía del header
    const header = sheet.getRange(1, columna).getValue();
    const match = header.match(/([A-Z]\d+)/);
    const codigoGuia = match ? match[1] : null;
    
    return {
      fecha: fecha,
      turno: turno,
      codigoGuia: codigoGuia,
      fila: fila,
      columna: columna
    };
  }
  
  // Procesar cambio de guía
  static procesarCambioGuia(guia, cambio, valor) {
    try {
      // Actualizar master calendar
      this.actualizarMasterDesdeGuia(guia, cambio, valor);
      
      // Actualizar colores en calendario guía
      this.actualizarColoresGuia(guia, cambio, valor);
      
    } catch (error) {
      console.error('Error procesando cambio de guía:', error);
    }
  }
  
  // Procesar cambio del master
  static procesarCambioMaster(cambio, valor) {
    try {
      const guia = ConfiguracionSistema.obtenerGuiaPorCodigo(cambio.codigoGuia);
      if (!guia) return;
      
      // Validar disponibilidad del guía
      if (!this.validarDisponibilidadGuia(guia, cambio)) {
        this.revertirCambioMaster(cambio);
        SpreadsheetApp.getUi().alert('❌ Error', 'El guía no está disponible en ese turno', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      // Actualizar calendario del guía
      this.actualizarGuiaDesdemaster(guia, cambio, valor);
      
      // Enviar notificación
      if (valor.includes('ASIGNAR')) {
        ServicioEmail.enviarAsignacionTour(guia.email, cambio.fecha, cambio.turno, valor);
      } else if (valor === 'LIBERAR') {
        ServicioEmail.enviarLiberacionTour(guia.email, cambio.fecha, cambio.turno);
      }
      
    } catch (error) {
      console.error('Error procesando cambio de master:', error);
    }
  }
  
  // Verificar si es celda de turno
  static esCeldaTurno(range, sheet) {
    const valor = range.getValue();
    const validationRule = range.getDataValidation();
    
    return validationRule && validationRule.getCriteriaValues();
  }
  
  // Extraer fecha de celda
  static extraerFechaDeCelda(range, sheet) {
    // Buscar número de día en celdas superiores
    const fila = range.getRow();
    const columna = range.getColumn();
    
    // Buscar en filas superiores el número del día
    for (let f = fila - 1; f >= 1; f--) {
      const valor = sheet.getRange(f, columna).getValue();
      if (typeof valor === 'number' && valor >= 1 && valor <= 31) {
        // Construir fecha basada en la pestaña
        const nombrePestana = sheet.getName();
        const [año, mes] = nombrePestana.split('-');
        return new Date(año, mes - 1, valor);
      }
    }
    return null;
  }
  
  // Extraer turno de celda
  static extraerTurnoDeCelda(range, sheet) {
    const fila = range.getRow();
    const columna = range.getColumn();
    
    // Buscar en celdas laterales "MAÑANA" o "TARDE"
    for (let c = columna - 1; c >= 1; c--) {
      const valor = sheet.getRange(fila, c).getValue();
      if (valor === 'MAÑANA' || valor === 'TARDE') {
        return valor;
      }
    }
    return null;
  }
  
  // Actualizar master desde guía
  static actualizarMasterDesdeGuia(guia, cambio, valor) {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet();
    const pestanaMes = this.obtenerPestanaMaster(cambio.fecha);
    const sheet = masterSheet.getSheetByName(pestanaMes);
    
    if (!sheet) return;
    
    // Buscar fila con la fecha
    const filaFecha = this.buscarFilaFecha(sheet, cambio.fecha);
    if (!filaFecha) return;
    
    // Buscar columna del guía y turno
    const columna = this.buscarColumnaGuiaTurno(sheet, guia.codigo, cambio.turno);
    if (!columna) return;
    
    // Actualizar celda en master
    const celda = sheet.getRange(filaFecha, columna);
    if (valor === 'NO DISPONIBLE') {
      celda.setBackground('#ff0000');
    } else if (valor === 'REVERTIR') {
      celda.setBackground('#ffffff');
      celda.setValue('');
    }
  }
  
  // Actualizar guía desde master
  static actualizarGuiaDesdemaster(guia, cambio, valor) {
    const guiaSheet = SpreadsheetApp.openById(guia.sheetId);
    const pestanaMes = this.obtenerPestanaMes(cambio.fecha);
    const sheet = guiaSheet.getSheetByName(pestanaMes);
    
    if (!sheet) return;
    
    // Buscar celda correspondiente
    const celda = this.buscarCeldaGuia(sheet, cambio.fecha, cambio.turno);
    if (!celda) return;
    
    if (valor.includes('ASIGNAR')) {
      celda.setValue(valor);
      celda.setBackground('#00ff00');
    } else if (valor === 'LIBERAR') {
      celda.setValue(cambio.turno);
      celda.setBackground('#ffffff');
    }
  }
  
  // Actualizar colores en calendario guía
  static actualizarColoresGuia(guia, cambio, valor) {
    const guiaSheet = SpreadsheetApp.openById(guia.sheetId);
    const pestanaMes = this.obtenerPestanaMes(cambio.fecha);
    const sheet = guiaSheet.getSheetByName(pestanaMes);
    
    if (!sheet) return;
    
    const celda = this.buscarCeldaGuia(sheet, cambio.fecha, cambio.turno);
    if (!celda) return;
    
    if (valor === 'NO DISPONIBLE') {
      celda.setBackground('#ff0000');
    } else if (valor === 'REVERTIR') {
      celda.setBackground('#ffffff');
      celda.setValue(cambio.turno);
    }
  }
  
  // Validar disponibilidad del guía
  static validarDisponibilidadGuia(guia, cambio) {
    const guiaSheet = SpreadsheetApp.openById(guia.sheetId);
    const pestanaMes = this.obtenerPestanaMes(cambio.fecha);
    const sheet = guiaSheet.getSheetByName(pestanaMes);
    
    if (!sheet) return false;
    
    const celda = this.buscarCeldaGuia(sheet, cambio.fecha, cambio.turno);
    if (!celda) return false;
    
    return celda.getValue() !== 'NO DISPONIBLE';
  }
  
  // Funciones auxiliares
  static obtenerPestanaMaster(fecha) {
    const año = fecha.getFullYear();
    const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
    return `${año}-${mes}`;
  }
  
  static obtenerPestanaMes(fecha) {
    const año = fecha.getFullYear();
    const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
    return `${año}-${mes}`;
  }
  
  static buscarFilaFecha(sheet, fecha) {
    const rango = sheet.getRange('A:A');
    const valores = rango.getValues();
    
    for (let i = 0; i < valores.length; i++) {
      if (valores[i][0] instanceof Date && 
          valores[i][0].getTime() === fecha.getTime()) {
        return i + 1;
      }
    }
    return null;
  }
  
  static buscarColumnaGuiaTurno(sheet, codigoGuia, turno) {
    const primeraFila = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (let i = 0; i < primeraFila.length; i++) {
      const header = primeraFila[i];
      if (header.includes(codigoGuia) && header.includes(turno)) {
        return i + 1;
      }
    }
    return null;
  }
  
  static buscarCeldaGuia(sheet, fecha, turno) {
    // Implementación simplificada - buscar por fecha y turno
    const datos = sheet.getDataRange().getValues();
    
    for (let fila = 0; fila < datos.length; fila++) {
      for (let col = 0; col < datos[fila].length; col++) {
        // Lógica para encontrar la celda correcta basada en fecha y turno
        // Esta implementación necesita ajustarse según el layout exacto
      }
    }
    return null;
  }
  
  static revertirCambioMaster(cambio) {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet();
    const pestanaMes = this.obtenerPestanaMaster(cambio.fecha);
    const sheet = masterSheet.getSheetByName(pestanaMes);
    
    if (!sheet) return;
    
    const filaFecha = this.buscarFilaFecha(sheet, cambio.fecha);
    const columna = this.buscarColumnaGuiaTurno(sheet, cambio.codigoGuia, cambio.turno);
    
    if (filaFecha && columna) {
      const celda = sheet.getRange(filaFecha, columna);
      celda.setValue('');
    }
  }
  
  // ===== TEMPLATE SYSTEM =====
  
  // Crear calendario individual para un guía
  static crearCalendarioGuia(codigo, nombre, email) {
    try {
      // Validar inputs
      if (!codigo || !nombre || !email) {
        throw new Error('Faltan datos obligatorios (código, nombre o email)');
      }
      
      if (!ConfiguracionSistema.validarCodigoGuia(codigo)) {
        throw new Error('Código de guía inválido. Formato: G01, G02, etc.');
      }
      
      if (!ConfiguracionSistema.validarEmail(email)) {
        throw new Error('Email inválido');
      }
      
      Logger.log(`Iniciando creación de calendario para ${codigo}`);
      
      const folderName = ConfiguracionSistema.FOLDER_GUIAS;
      Logger.log(`Buscando/creando carpeta: ${folderName}`);
      let folder = this.buscarOCrearCarpeta(folderName);
      Logger.log(`Carpeta obtenida: ${folder.getName()} (ID: ${folder.getId()})`);
      
      // Crear nombre de archivo seguro
      const nombreLimpio = nombre.replace(/[^a-zA-Z0-9]/g, '').substring(0, 10);
      const nombreArchivo = `Cal_${codigo}_${nombreLimpio}`;
      
      Logger.log(`Creando archivo: ${nombreArchivo}`);
      
      // Crear spreadsheet
      const nuevoSheet = SpreadsheetApp.create(nombreArchivo);
      const archivoSheet = DriveApp.getFileById(nuevoSheet.getId());
      
      Logger.log(`Archivo creado con ID: ${nuevoSheet.getId()}`);
      
      // Mover a carpeta
      Logger.log(`Moviendo archivo a carpeta: ${folder.getName()}`);
      archivoSheet.moveTo(folder);
      Logger.log(`Archivo movido exitosamente`);
      
      // Configurar pestañas de meses usando CalendarioGuia
      Logger.log(`Configurando pestañas de meses`);
      const calendario = new CalendarioGuia(nuevoSheet.getId(), codigo, nombre);
      ConfiguracionSistema.MESES_ACTIVOS.forEach(mes => {
        Logger.log(`Creando pestaña: ${mes}`);
        calendario.crearPestanaMes(mes);
      });
      
      // Eliminar Sheet1 por defecto
      const defaultSheet = nuevoSheet.getSheetByName('Sheet1');
      if (defaultSheet) {
        nuevoSheet.deleteSheet(defaultSheet);
        Logger.log(`Sheet1 por defecto eliminado`);
      }
      
      // Registrar guía en propiedades
      Logger.log(`Registrando guía en sistema`);
      this.registrarGuiaEnSistema(codigo, nombre, email, nuevoSheet.getId());
      
      // NUEVO: Actualizar master calendar automáticamente
      Logger.log(`Actualizando master calendar`);
      this.actualizarMasterConNuevoGuia(codigo, nombre);
      
      // Enviar email de bienvenida
      const urlCalendario = `https://docs.google.com/spreadsheets/d/${nuevoSheet.getId()}`;
      Logger.log(`Enviando email de bienvenida a: ${email}`);
      ServicioEmail.enviarBienvenidaGuia(email, nombre, codigo, urlCalendario);
      
      SpreadsheetApp.getUi().alert(
        'Guía Creado', 
        `Calendario creado para ${nombre} (${codigo})\nColumnas añadidas al master automáticamente\nEmail de bienvenida enviado.`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      Logger.log(`Guía ${codigo} creado exitosamente`);
      return nuevoSheet.getId();
      
    } catch (error) {
      Logger.log(`Error completo: ${error.message} - Stack: ${error.stack}`);
      SpreadsheetApp.getUi().alert('Error', `Error creando guía: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
      throw error;
    }
  }

  // NUEVA FUNCIÓN: Actualizar master calendar con nuevo guía
  static actualizarMasterConNuevoGuia(codigo, nombre) {
    try {
      const masterSheet = SpreadsheetApp.getActiveSpreadsheet();
      const meses = ConfiguracionSistema.MESES_ACTIVOS;
      
      meses.forEach(mes => {
        const sheet = masterSheet.getSheetByName(mes);
        if (!sheet) return;
        
        // Encontrar última columna usada
        const ultimaColumna = sheet.getLastColumn();
        
        // Añadir columnas para nuevo guía
        const colMañana = ultimaColumna + 1;
        const colTarde = ultimaColumna + 2;
        
        // Headers
        sheet.getRange(2, colMañana).setValue(`${codigo}-MAÑANA`);
        sheet.getRange(2, colTarde).setValue(`${codigo}-TARDE`);
        
        // Formato
        sheet.getRange(2, colMañana).setFontWeight('bold').setHorizontalAlignment('center');
        sheet.getRange(2, colTarde).setFontWeight('bold').setHorizontalAlignment('center');
        
        // Crear desplegables para todas las fechas del mes
        const numFilas = sheet.getLastRow();
        for (let fila = 3; fila <= numFilas; fila++) {
          // Columna mañana
          const celdaMañana = sheet.getRange(fila, colMañana);
          const reglaMañana = SpreadsheetApp.newDataValidation()
            .requireValueInList(ConfiguracionSistema.OPCIONES_MASTER_MANANA)
            .setAllowInvalid(false)
            .build();
          celdaMañana.setDataValidation(reglaMañana);
          celdaMañana.setHorizontalAlignment('center');
          
          // Columna tarde
          const celdaTarde = sheet.getRange(fila, colTarde);
          const reglaTarde = SpreadsheetApp.newDataValidation()
            .requireValueInList(ConfiguracionSistema.OPCIONES_MASTER_TARDE)
            .setAllowInvalid(false)
            .build();
          celdaTarde.setDataValidation(reglaTarde);
          celdaTarde.setHorizontalAlignment('center');
        }
        
        // Ajustar ancho de columnas
        sheet.setColumnWidth(colMañana, 120);
        sheet.setColumnWidth(colTarde, 120);
        
        Logger.log(`Columnas añadidas al mes ${mes}: ${codigo}-MAÑANA, ${codigo}-TARDE`);
      });
      
      Logger.log(`Master calendar actualizado para guía ${codigo}`);
      
    } catch (error) {
      Logger.log(`Error actualizando master: ${error.message}`);
      throw error;
    }
  }

  // Obtener o crear template si no existe
  static obtenerOCrearTemplate() {
    const props = PropertiesService.getScriptProperties();
    let templateId = props.getProperty('TEMPLATE_GUIA_ID');
    
    // Verificar si template existe
    if (templateId) {
      try {
        DriveApp.getFileById(templateId);
        return templateId;
      } catch (e) {
        templateId = null;
      }
    }
    
    // Crear nuevo template
    templateId = this.crearTemplateGuia();
    props.setProperty('TEMPLATE_GUIA_ID', templateId);
    return templateId;
  }

  // Crear template guía con código completo
  static crearTemplateGuia() {
    try {
      const templateSheet = SpreadsheetApp.create('TEMPLATE_Guía');
      const templateId = templateSheet.getId();
      
      this.configurarPestanasMeses(templateSheet);
      this.copiarCodigoATemplate(templateId);
      
      const templateFile = DriveApp.getFileById(templateId);
      const folderSistema = this.buscarOCrearCarpeta('Sistema_Tours');
      templateFile.moveTo(folderSistema);
      
      Logger.log(`Template creado: ${templateId}`);
      return templateId;
      
    } catch (error) {
      Logger.log(`Error creando template: ${error.message}`);
      throw error;
    }
  }

  // Copiar código Apps Script al template (simulado)
  static copiarCodigoATemplate(sheetId) {
    const sheet = SpreadsheetApp.openById(sheetId);
    const infoSheet = sheet.insertSheet('INSTRUCCIONES');
    
    const instrucciones = [
      ['INSTRUCCIONES PARA COMPLETAR TEMPLATE'],
      [''],
      ['1. Ir a Extensiones > Apps Script'],
      ['2. Copiar manualmente estos archivos:'],
      ['   - ConfiguracionSistema.gs'],
      ['   - ModeloCalendario.gs'], 
      ['   - ServicioSincronizacion.gs'],
      ['   - ServicioEmail.gs'],
      ['   - PanelControlGuia.gs'],
      [''],
      ['3. Instalar trigger: PanelControlGuia.instalarTriggerGuia()'],
      [''],
      ['ESTA HOJA SE ELIMINARÁ AL USAR EL TEMPLATE']
    ];
    
    infoSheet.getRange(1, 1, instrucciones.length, 1).setValues(instrucciones);
    infoSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
  }

  // Configurar sheet guía desde template
  static configurarSheetGuiaDesdeTemplate(sheetId, codigo, nombre, email) {
    try {
      const sheet = SpreadsheetApp.openById(sheetId);
      
      // Eliminar hoja de instrucciones si existe
      const instruccionesSheet = sheet.getSheetByName('INSTRUCCIONES');
      if (instruccionesSheet) {
        sheet.deleteSheet(instruccionesSheet);
      }
      
      Logger.log(`Sheet configurado para ${codigo}: ${sheetId}`);
      
    } catch (error) {
      Logger.log(`Error configurando sheet: ${error.message}`);
      throw error;
    }
  }
  
  // Buscar o crear carpeta dentro de la carpeta padre
  static buscarOCrearCarpeta(nombreCarpeta) {
    try {
      // Obtener carpeta padre
      const folderPadreId = ConfiguracionSistema.FOLDER_PADRE_ID;
      const folderPadre = DriveApp.getFolderById(folderPadreId);
      Logger.log(`Carpeta padre obtenida: ${folderPadre.getName()}`);
      
      // Buscar subcarpeta dentro del padre
      const subcarpetas = folderPadre.getFoldersByName(nombreCarpeta);
      
      if (subcarpetas.hasNext()) {
        Logger.log(`Subcarpeta ${nombreCarpeta} encontrada`);
        return subcarpetas.next();
      } else {
        Logger.log(`Creando subcarpeta: ${nombreCarpeta} dentro de ${folderPadre.getName()}`);
        const nuevaCarpeta = folderPadre.createFolder(nombreCarpeta);
        Logger.log(`Subcarpeta creada con ID: ${nuevaCarpeta.getId()}`);
        return nuevaCarpeta;
      }
    } catch (error) {
      Logger.log(`Error con carpeta ${nombreCarpeta}: ${error.message}`);
      // Si falla, usar carpeta raíz como fallback
      Logger.log(`Usando carpeta raíz como fallback`);
      return DriveApp.getRootFolder();
    }
  }
  
  // Configurar pestañas de meses en spreadsheet
  static configurarPestanasMeses(spreadsheet) {
    const meses = ConfiguracionSistema.MESES_ACTIVOS;
    
    // Eliminar Sheet1 por defecto
    const defaultSheet = spreadsheet.getSheetByName('Sheet1');
    if (defaultSheet) {
      spreadsheet.deleteSheet(defaultSheet);
    }
    
    meses.forEach(mes => {
      const calendario = new CalendarioGuia(spreadsheet.getId(), 'TEMP', 'TEMP');
      calendario.crearPestanaMes(mes);
    });
  }
  
  // Registrar guía en sistema
  static registrarGuiaEnSistema(codigo, nombre, email, sheetId) {
    const props = PropertiesService.getScriptProperties();
    const guiasExistentes = ConfiguracionSistema.getGuiasConfigurados();
    
    const nuevoGuia = {
      nombre: nombre,
      email: email,
      codigo: codigo,
      sheetId: sheetId
    };
    
    guiasExistentes.push(nuevoGuia);
    props.setProperty('GUIAS_CONFIGURADOS', JSON.stringify(guiasExistentes));
    
    Logger.log(`Guía registrado: ${codigo} - ${nombre}`);
  }
  
  // Eliminar calendario de guía y registro
  static eliminarCalendarioGuia(codigo) {
    try {
      const guias = ConfiguracionSistema.getGuiasConfigurados();
      const guiaIndex = guias.findIndex(g => g.codigo === codigo);
      
      if (guiaIndex === -1) {
        throw new Error(`Guía ${codigo} no encontrado`);
      }
      
      const guia = guias[guiaIndex];
      
      // Eliminar columnas del master calendar
      this.eliminarColumnasGuiaDelMaster(codigo);
      
      // Eliminar archivo de Drive
      try {
        const archivo = DriveApp.getFileById(guia.sheetId);
        archivo.setTrashed(true);
      } catch (e) {
        Logger.log(`Archivo ya eliminado o no encontrado: ${guia.sheetId}`);
      }
      
      // Eliminar de configuración
      guias.splice(guiaIndex, 1);
      const props = PropertiesService.getScriptProperties();
      props.setProperty('GUIAS_CONFIGURADOS', JSON.stringify(guias));
      
      SpreadsheetApp.getUi().alert(
        '✅ Guía Eliminado',
        `${guia.nombre} (${codigo}) eliminado exitosamente del sistema y master calendar`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      SpreadsheetApp.getUi().alert(
        '❌ Error',
        `Error eliminando guía: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }

  // Eliminar columnas de guía del master calendar
  static eliminarColumnasGuiaDelMaster(codigoGuia) {
    try {
      const masterSheet = SpreadsheetApp.getActiveSpreadsheet();
      const meses = ConfiguracionSistema.MESES_ACTIVOS;
      
      meses.forEach(mes => {
        const sheet = masterSheet.getSheetByName(mes);
        if (!sheet) return;
        
        // Obtener headers de la FILA 2 (no fila 1)
        const ultimaColumna = sheet.getLastColumn();
        const headers = sheet.getRange(2, 1, 1, ultimaColumna).getValues()[0];
        
        // Buscar columnas que contengan el código del guía (de atrás hacia adelante)
        for (let col = headers.length - 1; col >= 0; col--) {
          const header = headers[col];
          if (header && header.toString().includes(codigoGuia)) {
            Logger.log(`Eliminando columna ${col + 1}: ${header}`);
            sheet.deleteColumn(col + 1);
          }
        }
        
        Logger.log(`Columnas de ${codigoGuia} eliminadas de ${mes}`);
      });
      
    } catch (error) {
      Logger.log(`Error eliminando columnas del master: ${error.message}`);
    }
  }
  
  // Instalar triggers del sistema - SINTAXIS CORREGIDA
  static instalarTriggers() {
    try {
      // Limpiar triggers existentes
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
      
      // Crear trigger instalable con sintaxis correcta
      ScriptApp.newTrigger('onEditMasterCalendar')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
      
      SpreadsheetApp.getUi().alert('Triggers instalados', 'Trigger creado correctamente.', SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }

  // ===== SINCRONIZACIÓN MANUAL =====
  
  // Sincronizar todos los calendarios de guías hacia master
  static sincronizarTodosLosGuias() {
    try {
      const guias = ConfiguracionSistema.getGuiasConfigurados();
      if (guias.length === 0) {
        SpreadsheetApp.getUi().alert('❌ Error', 'No hay guías configurados', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      let cambiosDetectados = [];
      
      guias.forEach(guia => {
        try {
          const guiaSheet = SpreadsheetApp.openById(guia.sheetId);
          const meses = ConfiguracionSistema.MESES_ACTIVOS;
          
          meses.forEach(mes => {
            const pestana = guiaSheet.getSheetByName(mes);
            if (!pestana) return;
            
            // Leer todos los NO DISPONIBLE
            const datos = pestana.getDataRange().getValues();
            const colores = pestana.getDataRange().getBackgrounds();
            
            for (let i = 0; i < datos.length; i++) {
              for (let j = 0; j < datos[i].length; j++) {
                if (datos[i][j] === 'NO DISPONIBLE' || colores[i][j] === '#ff0000') {
                  // Extraer fecha y turno de esta posición
                  const cambio = this.extraerInfoCelda(pestana, i+1, j+1);
                  if (cambio) {
                    cambiosDetectados.push({
                      guia: guia.codigo,
                      fecha: cambio.fecha,
                      turno: cambio.turno,
                      estado: 'NO DISPONIBLE'
                    });
                  }
                }
              }
            }
          });
        } catch (error) {
          Logger.log(`Error leyendo guía ${guia.codigo}: ${error.message}`);
        }
      });
      
      // Aplicar cambios al master
      this.aplicarCambiosAlMaster(cambiosDetectados);
      
      SpreadsheetApp.getUi().alert(
        '✅ Sincronización Completa', 
        `${cambiosDetectados.length} cambios aplicados al master`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      SpreadsheetApp.getUi().alert('❌ Error', `Error: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }

  // Extraer información de celda específica
  static extraerInfoCelda(sheet, fila, columna) {
    try {
      // Buscar fecha del día en la misma columna hacia arriba
      let fecha = null;
      for (let f = fila - 1; f >= 1; f--) {
        const valor = sheet.getRange(f, columna).getValue();
        if (typeof valor === 'number' && valor >= 1 && valor <= 31) {
          const nombrePestana = sheet.getName();
          const [año, mes] = nombrePestana.split('-');
          fecha = new Date(año, mes - 1, valor);
          break;
        }
      }
      
      // Determinar turno basado en el valor de la fila anterior
      let turno = null;
      const valorFila = sheet.getRange(fila, 1).getValue();
      if (valorFila === 'MAÑANA' || valorFila === 'TARDE') {
        turno = valorFila;
      } else {
        // Buscar en celdas adyacentes
        for (let c = columna - 1; c >= 1; c--) {
          const valor = sheet.getRange(fila, c).getValue();
          if (valor === 'MAÑANA' || valor === 'TARDE') {
            turno = valor;
            break;
          }
        }
      }
      
      if (fecha && turno) {
        return { fecha, turno };
      }
      
    } catch (error) {
      Logger.log(`Error extrayendo info de celda: ${error.message}`);
    }
    
    return null;
  }

  // Aplicar cambios detectados al master calendar
  static aplicarCambiosAlMaster(cambios) {
    const masterSheet = SpreadsheetApp.getActiveSpreadsheet();
    
    cambios.forEach(cambio => {
      try {
        const pestanaMes = this.obtenerPestanaMaster(cambio.fecha);
        const sheet = masterSheet.getSheetByName(pestanaMes);
        
        if (!sheet) return;
        
        // Buscar fila de la fecha
        const filaFecha = this.buscarFilaFecha(sheet, cambio.fecha);
        if (!filaFecha) return;
        
        // Buscar columna del guía y turno
        const columna = this.buscarColumnaGuiaTurno(sheet, cambio.guia, cambio.turno);
        if (!columna) return;
        
        // Aplicar cambio
        const celda = sheet.getRange(filaFecha, columna);
        if (cambio.estado === 'NO DISPONIBLE') {
          celda.setBackground('#ff0000');
        }
        
      } catch (error) {
        Logger.log(`Error aplicando cambio: ${error.message}`);
      }
    });
  }
}