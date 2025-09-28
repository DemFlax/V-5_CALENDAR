/**
 * MODELOS DE DATOS DEL SISTEMA
 * Clases para manejar calendarios de guías y master
 */

class CalendarioGuia {
  constructor(sheetId, codigo, nombre) {
    this.sheetId = sheetId;
    this.codigo = codigo;
    this.nombre = nombre;
    this.spreadsheet = SpreadsheetApp.openById(sheetId);
  }
  
  // Crear pestaña de mes con formato calendario
  crearPestanaMes(mes) {
    const [año, mesNum] = mes.split('-');
    const primerDia = new Date(año, mesNum - 1, 1);
    const ultimoDia = new Date(año, mesNum, 0);
    const diasEnMes = ultimoDia.getDate();
    
    // Crear o obtener hoja
    let sheet = this.spreadsheet.getSheetByName(mes);
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(mes);
    }
    
    // Limpiar contenido existente
    sheet.clear();
    
    // Configurar título
    const nombreMes = ConfiguracionSistema.MESES_NOMBRES[mes] || mes;
    sheet.getRange('A1').setValue(`CALENDARIO ${this.codigo} - ${nombreMes}`);
    sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
    
    // Crear headers de días de la semana
    const diasSemana = ConfiguracionSistema.DIAS_SEMANA;
    for (let i = 0; i < diasSemana.length; i++) {
      sheet.getRange(2, i + 1).setValue(diasSemana[i]);
      sheet.getRange(2, i + 1).setFontWeight('bold').setHorizontalAlignment('center');
    }
    
    // Calcular posición del primer día
    const primerDiaSemana = primerDia.getDay();
    
    let fila = 3;
    let columna = primerDiaSemana + 1;
    
    // Llenar calendario
    for (let dia = 1; dia <= diasEnMes; dia++) {
      // Escribir número del día
      sheet.getRange(fila, columna).setValue(dia);
      sheet.getRange(fila, columna).setFontWeight('bold').setHorizontalAlignment('center');
      
      // Crear celdas de turnos SIN etiquetas en columna A
      this.crearCeldasTurno(sheet, fila + 1, columna, dia);
      this.crearCeldasTurno(sheet, fila + 2, columna, dia);
      
      // Siguiente posición
      columna++;
      if (columna > 7) {
        columna = 1;
        fila += 4;
      }
    }
    
    // Añadir etiquetas MAÑANA/TARDE sin validación
    this.crearEtiquetasTurnos(sheet);
    
    // Formato general
    this.aplicarFormatoCalendario(sheet);
    
    return sheet;
  }

  // Crear etiquetas de turnos sin validación
  crearEtiquetasTurnos(sheet) {
    // Solo texto informativo en primeras celdas
    sheet.getRange('A4').setValue('MAÑANA');
    sheet.getRange('A4').setFontWeight('bold').setBackground('#e6f3ff');
    
    sheet.getRange('A5').setValue('TARDE');
    sheet.getRange('A5').setFontWeight('bold').setBackground('#fff2e6');
  }
  
  // Crear celdas de turno con desplegables
  crearCeldasTurno(sheet, fila, columna, dia) {
    const celda = sheet.getRange(fila, columna);
    
    // NO establecer valor inicial - dejar vacío para evitar error de validación
    
    // Crear validación de datos (desplegable)
    const opciones = ConfiguracionSistema.OPCIONES_GUIA;
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(opciones)
      .setAllowInvalid(false)
      .build();
    
    celda.setDataValidation(regla);
    celda.setHorizontalAlignment('center');
    
    return celda;
  }
  
  // Aplicar formato visual al calendario
  aplicarFormatoCalendario(sheet) {
    // Ajustar anchos de columna
    for (let col = 1; col <= 7; col++) {
      sheet.setColumnWidth(col, 120);
    }
    
    // Colores alternos para las semanas
    const rango = sheet.getDataRange();
    const numFilas = rango.getNumRows();
    
    for (let fila = 3; fila <= numFilas; fila += 4) {
      if (Math.floor((fila - 3) / 4) % 2 === 0) {
        sheet.getRange(fila, 1, 3, 7).setBackground('#f0f0f0');
      }
    }
    
    // Bordes
    rango.setBorder(true, true, true, true, true, true);
  }
}

class MasterCalendar {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // Crear pestaña de mes para master calendar
  crearPestanaMes(mes, guias) {
    const [año, mesNum] = mes.split('-');
    const primerDia = new Date(año, mesNum - 1, 1);
    const ultimoDia = new Date(año, mesNum, 0);
    const diasEnMes = ultimoDia.getDate();
    
    // Crear o obtener hoja
    let sheet = this.spreadsheet.getSheetByName(mes);
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(mes);
    }
    
    // Limpiar contenido existente
    sheet.clear();
    
    // Configurar título
    const nombreMes = ConfiguracionSistema.MESES_NOMBRES[mes] || mes;
    sheet.getRange('A1').setValue(`MASTER CALENDAR - ${nombreMes}`);
    sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
    
    // Headers
    this.crearHeadersMaster(sheet, guias);
    
    // Crear filas de fechas
    this.crearFilasFechas(sheet, año, mesNum, diasEnMes, guias);
    
    // Aplicar formato
    this.aplicarFormatoMaster(sheet, guias);
    
    return sheet;
  }
  
  // Crear headers del master calendar
  crearHeadersMaster(sheet, guias) {
    let columna = 1;
    
    // Columna de fecha
    sheet.getRange(2, columna).setValue('FECHA');
    sheet.getRange(2, columna).setFontWeight('bold').setHorizontalAlignment('center');
    columna++;
    
    // Columnas para cada guía (mañana y tarde)
    guias.forEach(guia => {
      // Columna mañana
      sheet.getRange(2, columna).setValue(`${guia.codigo}-MAÑANA`);
      sheet.getRange(2, columna).setFontWeight('bold').setHorizontalAlignment('center');
      columna++;
      
      // Columna tarde
      sheet.getRange(2, columna).setValue(`${guia.codigo}-TARDE`);
      sheet.getRange(2, columna).setFontWeight('bold').setHorizontalAlignment('center');
      columna++;
    });
  }
  
  // Crear filas de fechas con desplegables
  crearFilasFechas(sheet, año, mesNum, diasEnMes, guias) {
    for (let dia = 1; dia <= diasEnMes; dia++) {
      const fecha = new Date(año, mesNum - 1, dia);
      const fila = dia + 2; // +2 porque empezamos en fila 3
      
      // Columna de fecha
      sheet.getRange(fila, 1).setValue(fecha);
      sheet.getRange(fila, 1).setNumberFormat('dd/mm/yyyy');
      sheet.getRange(fila, 1).setHorizontalAlignment('center');
      
      let columna = 2;
      
      // Crear desplegables para cada guía
      guias.forEach(guia => {
        // Desplegable mañana
        this.crearDesplegableMaster(sheet, fila, columna, 'MAÑANA');
        columna++;
        
        // Desplegable tarde
        this.crearDesplegableMaster(sheet, fila, columna, 'TARDE');
        columna++;
      });
    }
  }
  
  // Crear desplegable para master calendar
  crearDesplegableMaster(sheet, fila, columna, turno) {
    const celda = sheet.getRange(fila, columna);
    
    // Opciones según el turno
    const opciones = turno === 'MAÑANA' 
      ? ConfiguracionSistema.OPCIONES_MASTER_MANANA
      : ConfiguracionSistema.OPCIONES_MASTER_TARDE;
    
    // Crear validación de datos
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(opciones)
      .setAllowInvalid(false)
      .build();
    
    celda.setDataValidation(regla);
    celda.setHorizontalAlignment('center');
    
    return celda;
  }
  
  // Aplicar formato al master calendar
  aplicarFormatoMaster(sheet, guias) {
    // Ajustar anchos de columna
    sheet.setColumnWidth(1, 100); // Columna fecha
    
    for (let i = 2; i <= guias.length * 2 + 1; i++) {
      sheet.setColumnWidth(i, 120); // Columnas de guías
    }
    
    // Colores alternos para las fechas
    const numFilas = sheet.getLastRow();
    for (let fila = 3; fila <= numFilas; fila++) {
      if (fila % 2 === 0) {
        sheet.getRange(fila, 1, 1, guias.length * 2 + 1).setBackground('#f8f9fa');
      }
    }
    
    // Bordes
    const rango = sheet.getRange(1, 1, numFilas, guias.length * 2 + 1);
    rango.setBorder(true, true, true, true, true, true);
    
    // Congelar primera fila y primera columna
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(1);
  }
}

// Clase auxiliar para manejo de fechas
class UtilFechas {
  // Obtener nombre del día en español
  static obtenerNombreDia(fecha) {
    const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    return dias[fecha.getDay()];
  }
  
  // Obtener fechas de un mes
  static obtenerFechasMes(año, mes) {
    const fechas = [];
    const ultimoDia = new Date(año, mes, 0).getDate();
    
    for (let dia = 1; dia <= ultimoDia; dia++) {
      fechas.push(new Date(año, mes - 1, dia));
    }
    
    return fechas;
  }
  
  // Formatear fecha para display
  static formatearFecha(fecha) {
    const dia = fecha.getDate().toString().padStart(2, '0');
    const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
    const año = fecha.getFullYear();
    return `${dia}/${mes}/${año}`;
  }
}