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
    const timestampEjecucion = new Date(); // Timestamp de esta ejecución (Master)
    
    // 1. Obtener guías del registro
    const guias = obtenerGuiasDelRegistro();
    Logger.log(`Guías encontrados: ${guias.length}`);
    
    if (guias.length === 0) {
      Logger.log('No hay guías registrados. Fin de sincronización.');
      return;
    }
    
    // 2. Leer estado actual de la Hoja Maestra (con timestamp)
    leerEstadoHojaMaestra(guias, timestampEjecucion);
    
    // 3. Leer estado de cada calendario de guía
    for (const guia of guias) {
      leerEstadoCalendarioGuia(guia);
    }
    
    // 4. Resolver estados aplicando reglas de negocio
    for (const guia of guias) {
      for (const turno of guia.obtenerTodosTurnos()) {
        turno.resolverEstado();
      }
    }
    
    // 5. Escribir estados resueltos
    escribirEstadosEnHojaMaestra(guias);
    
    for (const guia of guias) {
      escribirEstadosEnCalendarioGuia(guia);
    }
    
    // 6. Procesar notificaciones y calendario
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
function leerEstadoHojaMaestra(guias, timestampEjecucion) {
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
        
        // Crear o actualizar turnos MAÑANA
        let turnoManana = guia.obtenerTurno(fecha, 'MANANA');
        if (!turnoManana) {
          turnoManana = new ClaseTurno(fecha, 'MANANA', guia);
          guia.agregarTurno(fecha, 'MANANA', turnoManana);
        }
        turnoManana.estadoMaster = estadoManana;
        turnoManana.timestampMaster = timestampEjecucion; // Timestamp del Master
        
        // Crear o actualizar turnos TARDE
        const tipoTarde = determinarTipoTurnoTarde(estadoTarde);
        let turnoTarde = guia.obtenerTurno(fecha, tipoTarde);
        if (!turnoTarde) {
          turnoTarde = new ClaseTurno(fecha, tipoTarde, guia);
          guia.agregarTurno(fecha, tipoTarde, turnoTarde);
        }
        turnoTarde.estadoMaster = estadoTarde;
        turnoTarde.timestampMaster = timestampEjecucion; // Timestamp del Master
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
      
      // Iterar por días del calendario
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
  const FILAS_POR_DIA = 3;
  
  // Iterar por BLOQUES de 3 filas (número día, MAÑANA, TARDE)
  // Empezar desde fila 1 (después de encabezados) y avanzar de 3 en 3
  for (let rowBloque = 1; rowBloque < data.length; rowBloque += FILAS_POR_DIA) {
    
    // Dentro de cada bloque, revisar las 7 columnas (días de la semana)
    for (let col = 0; col < 7; col++) {
      const celda = data[rowBloque][col];
      
      // Verificar si la celda contiene un número de día
      if (typeof celda === 'number' && celda >= 1 && celda <= 31) {
        const numeroDia = celda;
        
        try {
          const fecha = new Date(anio, mes - 1, numeroDia);
          
          // Leer MAÑANA (fila del bloque + 1)
          const filaManana = rowBloque + 1;
          if (filaManana >= data.length) continue;
          
          const estadoManana = data[filaManana][col] || '';
          const lockManana = data[filaManana][COL_LOCK] || '';
          const timestampManana = data[filaManana][COL_TIMESTAMP] || null;
          
          let turnoManana = guia.obtenerTurno(fecha, 'MANANA');
          if (!turnoManana) {
            turnoManana = new ClaseTurno(fecha, 'MANANA', guia);
            guia.agregarTurno(fecha, 'MANANA', turnoManana);
          }
          turnoManana.estadoGuia = estadoManana;
          turnoManana.lockStatusGuia = lockManana;
          turnoManana.timestampGuia = timestampManana;
          
          // Leer TARDE (fila del bloque + 2)
          const filaTarde = rowBloque + 2;
          if (filaTarde >= data.length) continue;
          
          const estadoTarde = data[filaTarde][col] || '';
          const lockTarde = data[filaTarde][COL_LOCK] || '';
          const timestampTarde = data[filaTarde][COL_TIMESTAMP] || null;
          
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
      const formateos = [];
      const desplegablesRecrear = []; // NUEVO
      const COL_LOCK = CONFIG.GUIA_CAL.COL_LOCK_STATUS;
      const COL_TIMESTAMP = CONFIG.GUIA_CAL.COL_TIMESTAMP;
      const FILAS_POR_DIA = 3;
      
      // Iterar por bloques de 3 filas
      for (let rowBloque = 1; rowBloque < data.length; rowBloque += FILAS_POR_DIA) {
        for (let col = 0; col < 7; col++) {
          const celda = data[rowBloque][col];
          
          if (typeof celda === 'number' && celda >= 1 && celda <= 31) {
            const numeroDia = celda;
            const fecha = new Date(anio, mes - 1, numeroDia);
            
            // Actualizar MAÑANA
            const turnoManana = guia.obtenerTurno(fecha, 'MANANA');
            if (turnoManana && turnoManana.requiereActualizacionGuia) {
              const filaManana = rowBloque + 1;
              
              actualizaciones.push({
                fila: filaManana + 1,
                columna: col + 1,
                valor: turnoManana.estadoFinal
              });
              actualizaciones.push({
                fila: filaManana + 1,
                columna: COL_LOCK + 1,
                valor: turnoManana.lockStatusFinal
              });
              actualizaciones.push({
                fila: filaManana + 1,
                columna: COL_TIMESTAMP + 1,
                valor: new Date()
              });
              
              const color = obtenerColorParaEstado(turnoManana.estadoFinal);
              formateos.push({
                fila: filaManana + 1,
                columna: col + 1,
                color: color
              });
              
              // Si el estado final es inicial (MAÑANA/TARDE), recrear desplegable
              if (turnoManana.estadoFinal === CONFIG.ESTADOS_VISIBLES.MANANA_INICIAL) {
                desplegablesRecrear.push({
                  fila: filaManana + 1,
                  columna: col + 1
                });
              }
            }
            
            // Actualizar TARDE
            const turnoTarde = buscarTurnoTarde(guia, fecha);
            if (turnoTarde && turnoTarde.requiereActualizacionGuia) {
              const filaTarde = rowBloque + 2;
              
              actualizaciones.push({
                fila: filaTarde + 1,
                columna: col + 1,
                valor: turnoTarde.estadoFinal
              });
              actualizaciones.push({
                fila: filaTarde + 1,
                columna: COL_LOCK + 1,
                valor: turnoTarde.lockStatusFinal
              });
              actualizaciones.push({
                fila: filaTarde + 1,
                columna: COL_TIMESTAMP + 1,
                valor: new Date()
              });
              
              const color = obtenerColorParaEstado(turnoTarde.estadoFinal);
              formateos.push({
                fila: filaTarde + 1,
                columna: col + 1,
                color: color
              });
              
              // Si el estado final es inicial (TARDE), recrear desplegable
              if (turnoTarde.estadoFinal === CONFIG.ESTADOS_VISIBLES.TARDE_INICIAL) {
                desplegablesRecrear.push({
                  fila: filaTarde + 1,
                  columna: col + 1
                });
              }
            }
          }
        }
      }
      
      // Aplicar actualizaciones y formatos
      aplicarActualizacionesEnLote(sheet, actualizaciones);
      aplicarFormateosEnLote(sheet, formateos);
      
      // Recrear desplegables donde sea necesario
      for (const desp of desplegablesRecrear) {
        crearDesplegableGuia(sheet, desp.fila, desp.columna);
      }
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