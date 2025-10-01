/**
 * MOTOR PRINCIPAL DE SINCRONIZACIÓN
 * Lógica central del sistema
 */

/**
 * Función principal de sincronización - llamada por trigger
 */
function ejecutarSincronizacion() {
  try {
    Logger.log('=== INICIO SINCRONIZACIÓN ===');
    const inicio = new Date();
    
    // 1. Obtener guías del registro
    const guias = obtenerGuiasDelRegistro();
    Logger.log(`Guías encontrados: ${guias.length}`);
    
    if (guias.length === 0) {
      Logger.log('No hay guías registrados. Fin de sincronización.');
      return;
    }
    
    // LOG: Info de guías
    for (const guia of guias) {
      Logger.log(`Guía: ${guia.codigo} - ${guia.nombre} - FileID: ${guia.fileId}`);
    }
    
    // 2. Leer estado actual de la Hoja Maestra
    Logger.log('--- Leyendo Hoja Maestra ---');
    leerEstadoHojaMaestra(guias);
    
    // LOG: Turnos leídos del Master
    for (const guia of guias) {
      const turnos = guia.obtenerTodosTurnos();
      Logger.log(`${guia.codigo}: ${turnos.length} turnos leídos del Master`);
      for (let i = 0; i < Math.min(3, turnos.length); i++) {
        const t = turnos[i];
        Logger.log(`  Master: ${formatearFecha(t.fecha)} ${t.tipoTurno} = "${t.estadoMaster}"`);
      }
    }
    
    // 3. Leer estado de cada calendario de guía
    Logger.log('--- Leyendo Calendarios de Guías ---');
    for (const guia of guias) {
      leerEstadoCalendarioGuia(guia);
      
      // LOG: Estados leídos del guía
      const turnos = guia.obtenerTodosTurnos();
      Logger.log(`${guia.codigo}: ${turnos.length} turnos después de leer calendario`);
      for (let i = 0; i < Math.min(3, turnos.length); i++) {
        const t = turnos[i];
        Logger.log(`  Guía: ${formatearFecha(t.fecha)} ${t.tipoTurno} = "${t.estadoGuia}" Lock="${t.lockStatusGuia}"`);
      }
    }
    
    // 4. Resolver estados aplicando reglas de negocio
    Logger.log('--- Resolviendo Estados ---');
    let contadorActualizacionesMaster = 0;
    let contadorActualizacionesGuia = 0;
    let contadorNotificaciones = 0;
    
    for (const guia of guias) {
      for (const turno of guia.obtenerTodosTurnos()) {
        turno.resolverEstado();
        
        if (turno.requiereActualizacionMaster) contadorActualizacionesMaster++;
        if (turno.requiereActualizacionGuia) contadorActualizacionesGuia++;
        if (turno.requiereNotificacion) contadorNotificaciones++;
        
        // LOG: Primeros turnos resueltos
        if (contadorActualizacionesMaster <= 3 || contadorActualizacionesGuia <= 3) {
          Logger.log(`${guia.codigo} ${formatearFecha(turno.fecha)} ${turno.tipoTurno}:`);
          Logger.log(`  EstadoFinal="${turno.estadoFinal}" Lock="${turno.lockStatusFinal}"`);
          Logger.log(`  ActMaster=${turno.requiereActualizacionMaster} ActGuia=${turno.requiereActualizacionGuia} Notif=${turno.requiereNotificacion}`);
        }
      }
    }
    
    Logger.log(`RESUMEN: ActMaster=${contadorActualizacionesMaster} ActGuia=${contadorActualizacionesGuia} Notif=${contadorNotificaciones}`);
    
    // 5. Escribir estados resueltos
    Logger.log('--- Escribiendo en Hoja Maestra ---');
    escribirEstadosEnHojaMaestra(guias);
    
    Logger.log('--- Escribiendo en Calendarios de Guías ---');
    for (const guia of guias) {
      escribirEstadosEnCalendarioGuia(guia);
    }
    
    // 6. Procesar notificaciones y calendario
    Logger.log('--- Procesando Notificaciones y Calendar ---');
    procesarNotificacionesYCalendario(guias);
    
    const fin = new Date();
    Logger.log(`=== FIN SINCRONIZACIÓN (${fin - inicio}ms) ===`);
    
  } catch (error) {
    Logger.log(`ERROR EN SINCRONIZACIÓN: ${error.toString()}`);
    Logger.log(error.stack);
    throw error;
  }
}

/**
 * Lee el estado actual de la Hoja Maestra
 */
function leerEstadoHojaMaestra(guias) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Buscar hojas mensuales (formato: MM_YYYY o YYYY_MM)
  for (const sheet of sheets) {
    const nombre = sheet.getName();
    if (!esHojaMensual(nombre)) continue;
    
    const data = sheet.getDataRange().getValues();
    const { mes, anio } = extraerMesAnioDeNombre(nombre);
    
    // Fila 0: Encabezados con códigos de guía
    // Fila 1: MAÑANA/TARDE
    // Fila 2+: Fechas y estados
    
    if (data.length < 3) continue;
    
    const encabezados = data[0];
    const subEncabezados = data[1];
    
    // Mapear columnas a guías
    const mapaColumnas = construirMapaColumnasGuias(encabezados, subEncabezados, guias);
    
    // Leer datos de fechas
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const fecha = row[0]; // Columna A
      
      if (!fecha || !(fecha instanceof Date)) continue;
      
      // Para cada guía, leer sus turnos
      for (const [guia, columnas] of mapaColumnas.entries()) {
        const estadoManana = row[columnas.manana] || '';
        const estadoTarde = row[columnas.tarde] || '';
        
        // Crear o actualizar turnos
        let turnoManana = guia.obtenerTurno(fecha, 'MANANA');
        if (!turnoManana) {
          turnoManana = new ClaseTurno(fecha, 'MANANA', guia);
          guia.agregarTurno(fecha, 'MANANA', turnoManana);
        }
        turnoManana.estadoMaster = estadoManana;
        
        // Determinar tipo de turno de tarde (T1, T2, T3 o genérico TARDE)
        const tipoTarde = determinarTipoTurnoTarde(estadoTarde);
        let turnoTarde = guia.obtenerTurno(fecha, tipoTarde);
        if (!turnoTarde) {
          turnoTarde = new ClaseTurno(fecha, tipoTarde, guia);
          guia.agregarTurno(fecha, tipoTarde, turnoTarde);
        }
        turnoTarde.estadoMaster = estadoTarde;
      }
    }
  }
}

/**
 * Lee el estado de un calendario de guía
 */
function leerEstadoCalendarioGuia(guia) {
  try {
    const ssGuia = SpreadsheetApp.openById(guia.fileId);
    const sheets = ssGuia.getSheets();
    
    for (const sheet of sheets) {
      const nombre = sheet.getName();
      if (!esHojaMensual(nombre)) continue;
      
      const { mes, anio } = extraerMesAnioDeNombre(nombre);
      const data = sheet.getDataRange().getValues();
      
      // Iterar por semanas y días
      leerDiasCalendarioGuia(data, mes, anio, guia);
    }
    
  } catch (error) {
    Logger.log(`Error leyendo calendario de ${guia.codigo}: ${error.toString()}`);
  }
}

/**
 * Lee los días de un calendario de guía específico
 */
function leerDiasCalendarioGuia(data, mes, anio, guia) {
  const COL_LOCK = CONFIG.GUIA_CAL.COL_LOCK_STATUS;
  const COL_TIMESTAMP = CONFIG.GUIA_CAL.COL_TIMESTAMP;
  
  // Recorrer TODAS las filas buscando números de día
  for (let row = 0; row < data.length - 2; row++) {
    for (let col = 0; col < 7; col++) {
      const celda = data[row][col];
      
      // Verificar si la celda contiene un número de día
      if (typeof celda === 'number' && celda >= 1 && celda <= 31) {
        const numeroDia = celda;
        
        try {
          const fecha = new Date(anio, mes - 1, numeroDia);
          
          // Leer MAÑANA (fila actual +1)
          const estadoManana = data[row + 1][col] || '';
          const lockManana = data[row + 1][COL_LOCK] || '';
          const timestampManana = data[row + 1][COL_TIMESTAMP] || null;
          
          let turnoManana = guia.obtenerTurno(fecha, 'MANANA');
          if (!turnoManana) {
            turnoManana = new ClaseTurno(fecha, 'MANANA', guia);
            guia.agregarTurno(fecha, 'MANANA', turnoManana);
          }
          turnoManana.estadoGuia = estadoManana;
          turnoManana.lockStatusGuia = lockManana;
          turnoManana.timestampGuia = timestampManana;
          
          // Leer TARDE (fila actual +2)
          const estadoTarde = data[row + 2][col] || '';
          const lockTarde = data[row + 2][COL_LOCK] || '';
          const timestampTarde = data[row + 2][COL_TIMESTAMP] || null;
          
          const tipoTarde = determinarTipoTurnoTarde(estadoTarde);
          let turnoTarde = guia.obtenerTurno(fecha, tipoTarde);
          if (!turnoTarde) {
            turnoTarde = new ClaseTurno(fecha, tipoTarde, guia);
            guia.agregarTurno(fecha, tipoTarde, turnoTarde);
          }
          turnoTarde.estadoGuia = estadoTarde;
          turnoTarde.lockStatusGuia = lockTarde;
          turnoTarde.timestampGuia = timestampTarde;
          
        } catch (error) {
          // Fecha inválida, saltar
          Logger.log(`Fecha inválida: ${numeroDia}/${mes}/${anio}`);
        }
      }
    }
  }
}

/**
 * Escribe los estados resueltos en la Hoja Maestra
 */
function escribirEstadosEnHojaMaestra(guias) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const VIS = CONFIG.ESTADOS_VISIBLES;
  
  for (const sheet of sheets) {
    const nombreHoja = sheet.getName();
    if (!esHojaMensual(nombreHoja)) continue;
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 3) continue;
    
    const encabezados = data[0];
    const subEncabezados = data[1];
    
    const mapaColumnas = construirMapaColumnasGuias(encabezados, subEncabezados, guias);
    
    // Preparar actualizaciones
    const actualizaciones = [];
    const formateos = [];
    
    for (let i = 2; i < data.length; i++) {
      const fecha = data[i][0];
      if (!(fecha instanceof Date)) continue;
      
      for (const [guia, columnas] of mapaColumnas.entries()) {
        const turnoManana = guia.obtenerTurno(fecha, 'MANANA');
        const turnoTarde = buscarTurnoTarde(guia, fecha);
        
        if (turnoManana && turnoManana.requiereActualizacionMaster) {
          const estadoFinal = turnoManana.estadoFinal;
          
          // SOLO escribir estados NO vacíos
          // Estados válidos: "NO DISPONIBLE", "ASIGNADO M"
          // NO escribir: "", "MAÑANA", "TARDE"
          if (estadoFinal === VIS.NO_DISPONIBLE || estadoFinal === VIS.ASIGNADO_M) {
            actualizaciones.push({
              fila: i + 1,
              columna: columnas.manana + 1,
              valor: estadoFinal
            });
            
            const color = obtenerColorParaEstado(estadoFinal);
            formateos.push({
              fila: i + 1,
              columna: columnas.manana + 1,
              color: color
            });
          }
        }
        
        if (turnoTarde && turnoTarde.requiereActualizacionMaster) {
          const estadoFinal = turnoTarde.estadoFinal;
          
          // SOLO escribir estados NO vacíos
          // Estados válidos: "NO DISPONIBLE", "ASIGNADO T1/T2/T3"
          // NO escribir: "", "MAÑANA", "TARDE"
          if (estadoFinal === VIS.NO_DISPONIBLE || 
              estadoFinal === VIS.ASIGNADO_T1 || 
              estadoFinal === VIS.ASIGNADO_T2 || 
              estadoFinal === VIS.ASIGNADO_T3) {
            actualizaciones.push({
              fila: i + 1,
              columna: columnas.tarde + 1,
              valor: estadoFinal
            });
            
            const color = obtenerColorParaEstado(estadoFinal);
            formateos.push({
              fila: i + 1,
              columna: columnas.tarde + 1,
              color: color
            });
          }
        }
      }
    }
    
    // Aplicar actualizaciones en lote
    aplicarActualizacionesEnLote(sheet, actualizaciones);
    aplicarFormateosEnLote(sheet, formateos);
  }
}

/**
 * Escribe los estados resueltos en el calendario de un guía
 */
function escribirEstadosEnCalendarioGuia(guia) {
  try {
    const ssGuia = SpreadsheetApp.openById(guia.fileId);
    const sheets = ssGuia.getSheets();
    
    for (const sheet of sheets) {
      const nombre = sheet.getName();
      if (!esHojaMensual(nombre)) continue;
      
      const { mes, anio } = extraerMesAnioDeNombre(nombre);
      const data = sheet.getDataRange().getValues();
      
      const actualizaciones = [];
      const COL_LOCK = CONFIG.GUIA_CAL.COL_LOCK_STATUS;
      const COL_TIMESTAMP = CONFIG.GUIA_CAL.COL_TIMESTAMP;
      const FILAS_POR_DIA = CONFIG.GUIA_CAL.FILAS_POR_DIA;
      
      // Recorrer días
      for (let row = 0; row < data.length; row += FILAS_POR_DIA) {
        for (let col = 0; col < 7; col++) {
          const numeroDia = data[row][col];
          
          if (typeof numeroDia === 'number' && numeroDia >= 1 && numeroDia <= 31) {
            const fecha = new Date(anio, mes - 1, numeroDia);
            
            // Validar que tenemos suficientes filas
            if (row + 2 >= data.length) continue;
            
            // Actualizar MAÑANA
            const turnoManana = guia.obtenerTurno(fecha, 'MANANA');
            if (turnoManana && turnoManana.requiereActualizacionGuia) {
              actualizaciones.push({
                fila: row + 2,
                columna: col + 1,
                valor: turnoManana.estadoFinal
              });
              actualizaciones.push({
                fila: row + 2,
                columna: COL_LOCK + 1,
                valor: turnoManana.lockStatusFinal
              });
              actualizaciones.push({
                fila: row + 2,
                columna: COL_TIMESTAMP + 1,
                valor: new Date()
              });
            }
            
            // Actualizar TARDE
            const turnoTarde = buscarTurnoTarde(guia, fecha);
            if (turnoTarde && turnoTarde.requiereActualizacionGuia) {
              actualizaciones.push({
                fila: row + 3,
                columna: col + 1,
                valor: turnoTarde.estadoFinal
              });
              actualizaciones.push({
                fila: row + 3,
                columna: COL_LOCK + 1,
                valor: turnoTarde.lockStatusFinal
              });
              actualizaciones.push({
                fila: row + 3,
                columna: COL_TIMESTAMP + 1,
                valor: new Date()
              });
            }
          }
        }
      }
      
      // Aplicar actualizaciones
      aplicarActualizacionesEnLote(sheet, actualizaciones);
    }
    
  } catch (error) {
    Logger.log(`Error escribiendo calendario de ${guia.codigo}: ${error.toString()}`);
  }
}

/**
 * Procesa notificaciones y actualiza Google Calendar
 */
function procesarNotificacionesYCalendario(guias) {
  const calendarId = obtenerCalendarIdMaestro();
  if (!calendarId) {
    Logger.log('ADVERTENCIA: Calendar ID no configurado');
    return;
  }
  
  for (const guia of guias) {
    for (const turno of guia.obtenerTodosTurnos()) {
      if (turno.requiereNotificacion) {
        // Agregar guía como invitado al evento
        agregarGuiaAEvento(calendarId, turno.fecha, turno.tipoTurno, guia.email);
        
        // Enviar notificación por email
        enviarNotificacionAsignacion(guia, turno);
      }
    }
  }
}