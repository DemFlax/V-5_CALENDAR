/**
 * MOTOR PRINCIPAL DE SINCRONIZACIÓN
 * VERSIÓN CON LOGGING EXHAUSTIVO PARA DIAGNÓSTICO
 */

/**
 * Función principal de sincronización - llamada por trigger
 */
function ejecutarSincronizacion() {
  try {
    Logger.log('═══════════════════════════════════════════════════════════');
    Logger.log('═══ INICIO SINCRONIZACIÓN ═══');
    Logger.log('═══════════════════════════════════════════════════════════');
    const inicio = new Date();
    const timestampEjecucion = new Date();
    
    // 1. Obtener guías del registro
    const guias = obtenerGuiasDelRegistro();
    Logger.log(`\n[FASE 1] Guías encontrados: ${guias.length}`);
    for (const guia of guias) {
      Logger.log(`  - ${guia.codigo}: ${guia.nombre}`);
    }
    
    if (guias.length === 0) {
      Logger.log('❌ No hay guías registrados. Fin de sincronización.');
      return;
    }
    
    // 2. Leer estado actual de la Hoja Maestra (con timestamp)
    Logger.log('\n[FASE 2] ═══ LEYENDO HOJA MAESTRA ═══');
    leerEstadoHojaMaestra(guias, timestampEjecucion);
    
    // DIAGNÓSTICO: Contar turnos después de leer Master
    Logger.log('\n[DIAGNÓSTICO POST-MASTER]');
    for (const guia of guias) {
      const turnos = guia.obtenerTodosTurnos();
      Logger.log(`  ${guia.codigo}: ${turnos.length} turnos creados desde Master`);
      
      // Mostrar primeros 5 turnos
      for (let i = 0; i < Math.min(5, turnos.length); i++) {
        const t = turnos[i];
        const fechaStr = `${t.fecha.getDate().toString().padStart(2,'0')}/${(t.fecha.getMonth()+1).toString().padStart(2,'0')}`;
        Logger.log(`    - ${fechaStr} ${t.tipoTurno}: Master="${t.estadoMaster}"`);
      }
    }
    
    // 3. Leer estado de cada calendario de guía
    Logger.log('\n[FASE 3] ═══ LEYENDO CALENDARIOS DE GUÍAS ═══');
    for (const guia of guias) {
      leerEstadoCalendarioGuia(guia);
    }
    
    // DIAGNÓSTICO: Contar turnos después de leer Guías
    Logger.log('\n[DIAGNÓSTICO POST-GUÍAS]');
    for (const guia of guias) {
      const turnos = guia.obtenerTodosTurnos();
      Logger.log(`  ${guia.codigo}: ${turnos.length} turnos totales`);
      
      // Detectar duplicados por fecha
      const fechasVistas = new Map();
      for (const t of turnos) {
        const key = `${t.fecha.getTime()}_${t.tipoTurno}`;
        if (fechasVistas.has(key)) {
          Logger.log(`    ⚠️ DUPLICADO: ${t.fecha.toDateString()} ${t.tipoTurno}`);
        }
        fechasVistas.set(key, true);
      }
      
      // Mostrar turnos con estado del Guía
      let turnosConEstadoGuia = 0;
      for (const t of turnos) {
        if (t.estadoGuia) {
          turnosConEstadoGuia++;
          const fechaStr = `${t.fecha.getDate().toString().padStart(2,'0')}/${(t.fecha.getMonth()+1).toString().padStart(2,'0')}`;
          Logger.log(`    - ${fechaStr} ${t.tipoTurno}: Guía="${t.estadoGuia}" Lock="${t.lockStatusGuia}"`);
        }
      }
      Logger.log(`  Total con estado del guía: ${turnosConEstadoGuia}`);
    }
    
    // 4. Resolver estados aplicando reglas de negocio
    Logger.log('\n[FASE 4] ═══ RESOLVIENDO ESTADOS ═══');
    let totalResueltos = 0;
    let totalConflictos = 0;
    for (const guia of guias) {
      for (const turno of guia.obtenerTodosTurnos()) {
        turno.resolverEstado();
        totalResueltos++;
        
        if (turno.requiereActualizacionMaster || turno.requiereActualizacionGuia) {
          const fechaStr = `${turno.fecha.getDate().toString().padStart(2,'0')}/${(turno.fecha.getMonth()+1).toString().padStart(2,'0')}`;
          Logger.log(`  ${guia.codigo} ${fechaStr} ${turno.tipoTurno}:`);
          Logger.log(`    Master="${turno.estadoMaster}" → Guía="${turno.estadoGuia}" → Final="${turno.estadoFinal}"`);
          Logger.log(`    Lock="${turno.lockStatusGuia}" → Final="${turno.lockStatusFinal}"`);
          Logger.log(`    Actualizar Master: ${turno.requiereActualizacionMaster}, Guía: ${turno.requiereActualizacionGuia}`);
          totalConflictos++;
        }
      }
    }
    Logger.log(`Total turnos procesados: ${totalResueltos}`);
    Logger.log(`Total con cambios: ${totalConflictos}`);
    
    // 5. Escribir estados resueltos
    Logger.log('\n[FASE 5] ═══ ESCRIBIENDO ESTADOS ═══');
    escribirEstadosEnHojaMaestra(guias);
    
    for (const guia of guias) {
      escribirEstadosEnCalendarioGuia(guia);
    }
    
    // 6. Procesar notificaciones y calendario
    Logger.log('\n[FASE 6] ═══ NOTIFICACIONES Y CALENDARIO ═══');
    procesarNotificacionesYCalendario(guias);
    
    const fin = new Date();
    Logger.log('\n═══════════════════════════════════════════════════════════');
    Logger.log(`═══ FIN SINCRONIZACIÓN (${fin - inicio}ms) ═══`);
    Logger.log('═══════════════════════════════════════════════════════════');
    
  } catch (error) {
    Logger.log(`\n❌❌❌ ERROR EN SINCRONIZACIÓN ❌❌❌`);
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`STACK: ${error.stack}`);
    throw error;
  }
}

/**
 * Lee el estado actual de la Hoja Maestra
 */
function leerEstadoHojaMaestra(guias, timestampEjecucion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  for (const sheet of sheets) {
    const nombre = sheet.getName();
    if (!esHojaMensual(nombre)) continue;
    
    Logger.log(`\n  ► Leyendo hoja Master: ${nombre}`);
    
    const data = sheet.getDataRange().getValues();
    const { mes, anio } = extraerMesAnioDeNombre(nombre);
    
    if (data.length < 3) {
      Logger.log(`    ⚠️ Hoja vacía o sin datos`);
      continue;
    }
    
    const encabezados = data[0];
    const subEncabezados = data[1];
    
    // Mapear columnas a guías
    const mapaColumnas = construirMapaColumnasGuias(encabezados, subEncabezados, guias);
    Logger.log(`    Columnas mapeadas para ${mapaColumnas.size} guías`);
    
    // Leer datos de fechas
    let turnosCreados = 0;
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const fecha = row[0];
      
      if (!fecha || !(fecha instanceof Date)) continue;
      
      const fechaStr = `${fecha.getDate().toString().padStart(2,'0')}/${(fecha.getMonth()+1).toString().padStart(2,'0')}/${fecha.getFullYear()}`;
      
      // Para cada guía, leer sus turnos
      for (const [guia, columnas] of mapaColumnas.entries()) {
        const estadoManana = row[columnas.manana] || '';
        const estadoTarde = row[columnas.tarde] || '';
        
        // Crear o actualizar turnos MAÑANA
        let turnoManana = guia.obtenerTurno(fecha, 'MANANA');
        if (!turnoManana) {
          turnoManana = new ClaseTurno(fecha, 'MANANA', guia);
          guia.agregarTurno(fecha, 'MANANA', turnoManana);
          turnosCreados++;
        }
        turnoManana.estadoMaster = estadoManana;
        turnoManana.timestampMaster = timestampEjecucion;
        
        // Crear o actualizar turnos TARDE
        const tipoTarde = determinarTipoTurnoTarde(estadoTarde);
        let turnoTarde = guia.obtenerTurno(fecha, tipoTarde);
        if (!turnoTarde) {
          turnoTarde = new ClaseTurno(fecha, tipoTarde, guia);
          guia.agregarTurno(fecha, tipoTarde, turnoTarde);
          turnosCreados++;
        }
        turnoTarde.estadoMaster = estadoTarde;
        turnoTarde.timestampMaster = timestampEjecucion;
      }
    }
    Logger.log(`    Turnos creados en esta hoja: ${turnosCreados}`);
  }
}

/**
 * Lee el estado de un calendario de guía
 */
function leerEstadoCalendarioGuia(guia) {
  try {
    Logger.log(`\n  ► Leyendo calendario de ${guia.codigo}`);
    const ssGuia = SpreadsheetApp.openById(guia.fileId);
    const sheets = ssGuia.getSheets();
    
    let hojasLeidas = 0;
    for (const sheet of sheets) {
      const nombre = sheet.getName();
      if (!esHojaMensual(nombre)) continue;
      
      Logger.log(`    Hoja: ${nombre}`);
      const { mes, anio } = extraerMesAnioDeNombre(nombre);
      const data = sheet.getDataRange().getValues();
      
      leerDiasCalendarioGuia(data, mes, anio, guia);
      hojasLeidas++;
    }
    
    if (hojasLeidas === 0) {
      Logger.log(`    ⚠️ No se encontraron hojas mensuales`);
    }
    
  } catch (error) {
    Logger.log(`    ❌ Error: ${error.toString()}`);
  }
}

/**
 * Lee los días de un calendario de guía específico
 */
function leerDiasCalendarioGuia(data, mes, anio, guia) {
  const COL_LOCK = CONFIG.GUIA_CAL.COL_LOCK_STATUS;
  const COL_TIMESTAMP = CONFIG.GUIA_CAL.COL_TIMESTAMP;
  const FILAS_POR_DIA = 3;
  
  let diasLeidos = 0;
  let turnosActualizados = 0;
  
  Logger.log(`      Estructura: ${data.length} filas × ${data[0] ? data[0].length : 0} columnas`);
  Logger.log(`      Leyendo bloques desde fila 1, avanzando de ${FILAS_POR_DIA} en ${FILAS_POR_DIA}...`);
  
  // Iterar por BLOQUES de 3 filas
  for (let rowBloque = 1; rowBloque < data.length; rowBloque += FILAS_POR_DIA) {
    
    // Revisar las 7 columnas (días de la semana)
    for (let col = 0; col < 7; col++) {
      const celda = data[rowBloque][col];
      
      // Verificar si la celda contiene un número de día
      if (typeof celda === 'number' && celda >= 1 && celda <= 31) {
        const numeroDia = celda;
        diasLeidos++;
        
        try {
          // CRÍTICO: Construcción de la fecha
          const fecha = new Date(anio, mes - 1, numeroDia);
          const fechaStr = `${numeroDia.toString().padStart(2,'0')}/${mes.toString().padStart(2,'0')}/${anio}`;
          const fechaTimestamp = fecha.getTime();
          
          Logger.log(`      Día ${numeroDia} (Col ${col}, Fila base ${rowBloque}):`);
          Logger.log(`        Fecha calculada: ${fechaStr} (timestamp: ${fechaTimestamp})`);
          
          // Leer MAÑANA
          const filaManana = rowBloque + 1;
          if (filaManana < data.length) {
            const estadoManana = data[filaManana][col] || '';
            const lockManana = data[filaManana][COL_LOCK] || '';
            const timestampManana = data[filaManana][COL_TIMESTAMP] || null;
            
            Logger.log(`        MAÑANA (fila ${filaManana}): estado="${estadoManana}", lock="${lockManana}"`);
            
            let turnoManana = guia.obtenerTurno(fecha, 'MANANA');
            if (!turnoManana) {
              Logger.log(`        ⚠️ NO existe turno MAÑANA para ${fechaStr} - Creando nuevo`);
              turnoManana = new ClaseTurno(fecha, 'MANANA', guia);
              guia.agregarTurno(fecha, 'MANANA', turnoManana);
            } else {
              Logger.log(`        ✓ Turno MAÑANA ya existe - Actualizando`);
            }
            turnoManana.estadoGuia = estadoManana;
            turnoManana.lockStatusGuia = lockManana;
            turnoManana.timestampGuia = timestampManana;
            turnosActualizados++;
          }
          
          // Leer TARDE
          const filaTarde = rowBloque + 2;
          if (filaTarde < data.length) {
            const estadoTarde = data[filaTarde][col] || '';
            const lockTarde = data[filaTarde][COL_LOCK] || '';
            const timestampTarde = data[filaTarde][COL_TIMESTAMP] || null;
            
            Logger.log(`        TARDE (fila ${filaTarde}): estado="${estadoTarde}", lock="${lockTarde}"`);
            
            const tipoTarde = determinarTipoTurnoTarde(estadoTarde);
            let turnoTarde = guia.obtenerTurno(fecha, tipoTarde);
            if (!turnoTarde) {
              Logger.log(`        ⚠️ NO existe turno TARDE (${tipoTarde}) para ${fechaStr} - Creando nuevo`);
              turnoTarde = new ClaseTurno(fecha, tipoTarde, guia);
              guia.agregarTurno(fecha, tipoTarde, turnoTarde);
            } else {
              Logger.log(`        ✓ Turno TARDE ya existe - Actualizando`);
            }
            turnoTarde.estadoGuia = estadoTarde;
            turnoTarde.lockStatusGuia = lockTarde;
            turnoTarde.timestampGuia = timestampTarde;
            turnosActualizados++;
          }
          
        } catch (error) {
          Logger.log(`        ❌ Error procesando día ${numeroDia}: ${error.toString()}`);
        }
      }
    }
  }
  
  Logger.log(`      Total días leídos: ${diasLeidos}`);
  Logger.log(`      Total turnos actualizados: ${turnosActualizados}`);
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
    
    Logger.log(`\n  ► Escribiendo en hoja Master: ${nombreHoja}`);
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 3) continue;
    
    const encabezados = data[0];
    const subEncabezados = data[1];
    
    const mapaColumnas = construirMapaColumnasGuias(encabezados, subEncabezados, guias);
    
    const actualizaciones = [];
    const formateos = [];
    const desplegablesRecrear = [];
    
    for (let i = 2; i < data.length; i++) {
      const fecha = data[i][0];
      if (!(fecha instanceof Date)) continue;
      
      const fechaStr = `${fecha.getDate()}/${fecha.getMonth()+1}/${fecha.getFullYear()}`;
      
      for (const [guia, columnas] of mapaColumnas.entries()) {
        const turnoManana = guia.obtenerTurno(fecha, 'MANANA');
        const turnoTarde = buscarTurnoTarde(guia, fecha);
        
        if (turnoManana && turnoManana.requiereActualizacionMaster) {
          const estadoFinal = turnoManana.estadoFinal;
          const filaHoja = i + 1;
          const colHoja = columnas.manana + 1;
          
          Logger.log(`    [${fechaStr}] ${guia.codigo} MAÑANA → Fila ${filaHoja}, Col ${colHoja}: "${estadoFinal}"`);
          
          actualizaciones.push({ fila: filaHoja, columna: colHoja, valor: estadoFinal });
          formateos.push({ fila: filaHoja, columna: colHoja, color: obtenerColorParaEstado(estadoFinal) });
          
          if (estadoFinal === '' || estadoFinal === VIS.MANANA_INICIAL) {
            desplegablesRecrear.push({ fila: filaHoja, columna: colHoja, tipo: 'MANANA' });
          }
        }
        
        if (turnoTarde && turnoTarde.requiereActualizacionMaster) {
          const estadoFinal = turnoTarde.estadoFinal;
          const filaHoja = i + 1;
          const colHoja = columnas.tarde + 1;
          
          Logger.log(`    [${fechaStr}] ${guia.codigo} TARDE → Fila ${filaHoja}, Col ${colHoja}: "${estadoFinal}"`);
          
          actualizaciones.push({ fila: filaHoja, columna: colHoja, valor: estadoFinal });
          formateos.push({ fila: filaHoja, columna: colHoja, color: obtenerColorParaEstado(estadoFinal) });
          
          if (estadoFinal === '' || estadoFinal === VIS.TARDE_INICIAL) {
            desplegablesRecrear.push({ fila: filaHoja, columna: colHoja, tipo: 'TARDE' });
          }
        }
      }
    }
    
    Logger.log(`    Total actualizaciones: ${actualizaciones.length}`);
    Logger.log(`    Total desplegables a recrear: ${desplegablesRecrear.length}`);
    
    aplicarActualizacionesEnLote(sheet, actualizaciones);
    aplicarFormateosEnLote(sheet, formateos);
    
    for (const desp of desplegablesRecrear) {
      crearDesplegableMaster(sheet, desp.fila, desp.columna, desp.tipo);
    }
  }
}

/**
 * Escribe los estados resueltos en el calendario de un guía
 */
function escribirEstadosEnCalendarioGuia(guia) {
  try {
    Logger.log(`\n  ► Escribiendo en calendario de ${guia.codigo}`);
    const ssGuia = SpreadsheetApp.openById(guia.fileId);
    const sheets = ssGuia.getSheets();
    
    for (const sheet of sheets) {
      const nombre = sheet.getName();
      if (!esHojaMensual(nombre)) continue;
      
      const { mes, anio } = extraerMesAnioDeNombre(nombre);
      const data = sheet.getDataRange().getValues();
      
      const actualizaciones = [];
      const formateos = [];
      const desplegablesRecrear = [];
      const COL_LOCK = CONFIG.GUIA_CAL.COL_LOCK_STATUS;
      const COL_TIMESTAMP = CONFIG.GUIA_CAL.COL_TIMESTAMP;
      const FILAS_POR_DIA = 3;
      
      for (let rowBloque = 1; rowBloque < data.length; rowBloque += FILAS_POR_DIA) {
        for (let col = 0; col < 7; col++) {
          const celda = data[rowBloque][col];
          
          if (typeof celda === 'number' && celda >= 1 && celda <= 31) {
            const numeroDia = celda;
            const fecha = new Date(anio, mes - 1, numeroDia);
            
            const turnoManana = guia.obtenerTurno(fecha, 'MANANA');
            if (turnoManana && turnoManana.requiereActualizacionGuia) {
              const filaManana = rowBloque + 1;
              
              actualizaciones.push({ fila: filaManana + 1, columna: col + 1, valor: turnoManana.estadoFinal });
              actualizaciones.push({ fila: filaManana + 1, columna: COL_LOCK + 1, valor: turnoManana.lockStatusFinal });
              actualizaciones.push({ fila: filaManana + 1, columna: COL_TIMESTAMP + 1, valor: new Date() });
              
              formateos.push({ fila: filaManana + 1, columna: col + 1, color: obtenerColorParaEstado(turnoManana.estadoFinal) });
              
              if (turnoManana.estadoFinal === CONFIG.ESTADOS_VISIBLES.MANANA_INICIAL) {
                desplegablesRecrear.push({ fila: filaManana + 1, columna: col + 1 });
              }
            }
            
            const turnoTarde = buscarTurnoTarde(guia, fecha);
            if (turnoTarde && turnoTarde.requiereActualizacionGuia) {
              const filaTarde = rowBloque + 2;
              
              actualizaciones.push({ fila: filaTarde + 1, columna: col + 1, valor: turnoTarde.estadoFinal });
              actualizaciones.push({ fila: filaTarde + 1, columna: COL_LOCK + 1, valor: turnoTarde.lockStatusFinal });
              actualizaciones.push({ fila: filaTarde + 1, columna: COL_TIMESTAMP + 1, valor: new Date() });
              
              formateos.push({ fila: filaTarde + 1, columna: col + 1, color: obtenerColorParaEstado(turnoTarde.estadoFinal) });
              
              if (turnoTarde.estadoFinal === CONFIG.ESTADOS_VISIBLES.TARDE_INICIAL) {
                desplegablesRecrear.push({ fila: filaTarde + 1, columna: col + 1 });
              }
            }
          }
        }
      }
      
      Logger.log(`    Hoja ${nombre}: ${actualizaciones.length} actualizaciones`);
      
      aplicarActualizacionesEnLote(sheet, actualizaciones);
      aplicarFormateosEnLote(sheet, formateos);
      
      for (const desp of desplegablesRecrear) {
        crearDesplegableGuia(sheet, desp.fila, desp.columna);
      }
    }
    
  } catch (error) {
    Logger.log(`    ❌ Error: ${error.toString()}`);
  }
}

/**
 * Procesa notificaciones y actualiza Google Calendar
 */
function procesarNotificacionesYCalendario(guias) {
  const calendarId = obtenerCalendarIdMaestro();
  if (!calendarId) {
    Logger.log('  ⚠️ Calendar ID no configurado');
    return;
  }
  
  let notificacionesEnviadas = 0;
  for (const guia of guias) {
    for (const turno of guia.obtenerTodosTurnos()) {
      if (turno.requiereNotificacion) {
        agregarGuiaAEvento(calendarId, turno.fecha, turno.tipoTurno, guia.email);
        enviarNotificacionAsignacion(guia, turno);
        notificacionesEnviadas++;
      }
    }
  }
  
  Logger.log(`  Total notificaciones enviadas: ${notificacionesEnviadas}`);
}

/**
 * NUEVA: Crea desplegable del Master en una celda específica
 */
function crearDesplegableMaster(sheet, fila, columna, tipo) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  
  let opciones;
  if (tipo === 'MANANA') {
    opciones = [VIS.ASIGNAR_MANANA, VIS.LIBERAR_MASTER];
  } else {
    opciones = [VIS.ASIGNAR_T1, VIS.ASIGNAR_T2, VIS.ASIGNAR_T3, VIS.LIBERAR_MASTER];
  }
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}