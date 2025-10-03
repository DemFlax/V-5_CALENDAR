/**
 * CREACIÓN DE NUEVOS GUÍAS
 * Generación de calendario y actualización automática del Master
 */

/**
 * UI para crear nuevo guía desde el menú
 */
function crearNuevoGuiaUI() {
  const ui = SpreadsheetApp.getUi();
  
  // Verificar que REGISTRO exista
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CONFIG.SHEET_REGISTRO)) {
    ui.alert(
      '❌ Sistema No Inicializado',
      'Primero debe inicializar el sistema.\n\n' +
      'Vaya al menú: 🔄 Sincronizador > 🚀 Inicializar Sistema',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Solicitar nombre
  const respuestaNombre = ui.prompt(
    'Crear Nuevo Guía - Paso 1/2',
    'Ingrese el nombre del guía:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuestaNombre.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const nombre = respuestaNombre.getResponseText().trim();
  if (!nombre) {
    ui.alert('Debe ingresar un nombre válido');
    return;
  }
  
  // Solicitar email
  const respuestaEmail = ui.prompt(
    'Crear Nuevo Guía - Paso 2/2',
    'Ingrese el email del guía:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (respuestaEmail.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const email = respuestaEmail.getResponseText().trim();
  if (!esEmailValido(email)) {
    ui.alert('Debe ingresar un email válido');
    return;
  }
  
  // Generar código automático
  const codigo = generarSiguienteCodigoGuia();
  
  // Confirmar
  const confirmacion = ui.alert(
    'Confirmar Creación',
    `Código: ${codigo}\nNombre: ${nombre}\nEmail: ${email}\n\n¿Crear este guía?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmacion !== ui.Button.YES) {
    return;
  }
  
  try {
    ui.alert('Creando guía y actualizando hojas...');
    const resultado = crearNuevoGuia(codigo, nombre, email);
    
    if (resultado.exito) {
      ui.alert(
        '✅ Guía Creado Exitosamente',
        `Código: ${codigo}\n` +
        `Calendario: ${resultado.url}\n` +
        `Columnas agregadas al Master\n\n` +
        `Email enviado a: ${email}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error: ' + resultado.error);
    }
    
  } catch (error) {
    ui.alert('Error creando guía: ' + error.toString());
  }
}

/**
 * Crea un nuevo guía completo (calendario + registro + Master + notificación)
 */
function crearNuevoGuia(codigo, nombre, email) {
  try {
    // 1. Crear calendario del guía
    const nombreCalendario = `Calendario - ${nombre} (${codigo})`;
    const ssGuia = SpreadsheetApp.create(nombreCalendario);
    const fileId = ssGuia.getId();
    const url = ssGuia.getUrl();
    
    configurarCalendarioGuia(ssGuia);
    
    // 2. Mover a la carpeta de guías
    const file = DriveApp.getFileById(fileId);
    const carpetaDestino = DriveApp.getFolderById(CONFIG.CARPETA_GUIAS_ID);
    
    // Agregar a carpeta destino
    carpetaDestino.addFile(file);
    
    // Remover de la raíz (Mi unidad)
    const padres = file.getParents();
    while (padres.hasNext()) {
      const padre = padres.next();
      if (padre.getId() !== CONFIG.CARPETA_GUIAS_ID) {
        padre.removeFile(file);
      }
    }
    
    // 3. Dar permisos de edición al guía
    file.addEditor(email);
    
    // 4. Registrar en REGISTRO
    agregarGuiaAlRegistro(codigo, nombre, email, fileId, url);
    
    // 5. Agregar columnas del guía al Master
    agregarColumnasGuiaEnMaster(codigo, nombre);
    
    // 6. Enviar email
    enviarEmailCreacionGuia(nombre, email, url);
    
    Logger.log(`Guía ${codigo} creado y movido a carpeta ${CONFIG.CARPETA_GUIAS_ID}`);
    
    return { exito: true, fileId: fileId, url: url };
    
  } catch (error) {
    Logger.log(`Error en crearNuevoGuia: ${error.toString()}`);
    return { exito: false, error: error.toString() };
  }
}

/**
 * Agrega las columnas MAÑANA/TARDE del nuevo guía en todas las hojas mensuales del Master
 */
function agregarColumnasGuiaEnMaster(codigo, nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  for (const sheet of sheets) {
    const nombreHoja = sheet.getName();
    if (!esHojaMensual(nombreHoja)) continue;
    
    // Encontrar la próxima columna disponible
    const ultimaColumna = sheet.getLastColumn();
    const colManana = ultimaColumna + 1;
    const colTarde = ultimaColumna + 2;
    
    // Fila 1: Código y Nombre (merge de 2 columnas)
    const rangoEncabezado = sheet.getRange(1, colManana, 1, 2);
    rangoEncabezado.merge();
    rangoEncabezado.setValue(`${codigo} — ${nombre}`);
    rangoEncabezado.setFontWeight('bold');
    rangoEncabezado.setHorizontalAlignment('center');
    rangoEncabezado.setBackground('#4285F4');
    rangoEncabezado.setFontColor('#FFFFFF');
    
    // Fila 2: MAÑANA y TARDE
    sheet.getRange(2, colManana).setValue('MAÑANA');
    sheet.getRange(2, colTarde).setValue('TARDE');
    sheet.getRange(2, colManana, 1, 2).setFontWeight('bold');
    sheet.getRange(2, colManana, 1, 2).setHorizontalAlignment('center');
    sheet.getRange(2, colManana, 1, 2).setBackground('#E8F0FE');
    
    // Ajustar anchos
    sheet.setColumnWidth(colManana, 150);
    sheet.setColumnWidth(colTarde, 150);
    
    // Agregar desplegables en todas las filas de datos
    const ultimaFila = sheet.getLastRow();
    if (ultimaFila >= 3) {
      crearDesplegablesMasterEnRango(sheet, 3, ultimaFila, colManana, colTarde);
    }
  }
  
  Logger.log(`Columnas agregadas para ${codigo} en todas las hojas mensuales`);
}

/**
 * Crea desplegables del Master en un rango específico
 */
function crearDesplegablesMasterEnRango(sheet, filaInicio, filaFin, colManana, colTarde) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  
  // Opciones para MAÑANA
  const opcionesManana = [
    VIS.ASIGNAR_MANANA,
    VIS.LIBERAR_MASTER
  ];
  
  const reglaManana = SpreadsheetApp.newDataValidation()
    .requireValueInList(opcionesManana, true)
    .setAllowInvalid(false)
    .build();
  
  // Opciones para TARDE
  const opcionesTarde = [
    VIS.ASIGNAR_T1,
    VIS.ASIGNAR_T2,
    VIS.ASIGNAR_T3,
    VIS.LIBERAR_MASTER
  ];
  
  const reglaTarde = SpreadsheetApp.newDataValidation()
    .requireValueInList(opcionesTarde, true)
    .setAllowInvalid(false)
    .build();
  
  // Aplicar a todas las filas
  sheet.getRange(filaInicio, colManana, filaFin - filaInicio + 1, 1)
    .setDataValidation(reglaManana);
  
  sheet.getRange(filaInicio, colTarde, filaFin - filaInicio + 1, 1)
    .setDataValidation(reglaTarde);
}

/**
 * Configura el calendario del guía con hojas mensuales
 */
function configurarCalendarioGuia(ssGuia) {
  const hojaDefault = ssGuia.getSheets()[0];
  
  const hoy = new Date();
  
  for (let i = 0; i < 3; i++) {
    const fecha = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
    const mes = fecha.getMonth() + 1;
    const anio = fecha.getFullYear();
    
    crearHojaMensualGuia(ssGuia, mes, anio);
  }
  
  if (ssGuia.getSheets().length > 1) {
    ssGuia.deleteSheet(hojaDefault);
  }
}

/**
 * Crea una hoja mensual en el calendario del guía
 */
function crearHojaMensualGuia(ssGuia, mes, anio) {
  const nombreHoja = `${mes}_${anio}`;
  const sheet = ssGuia.insertSheet(nombreHoja);
  
  const primerDia = new Date(anio, mes - 1, 1);
  const ultimoDia = new Date(anio, mes, 0).getDate();
  const primerDiaSemana = primerDia.getDay();
  
  // Encabezados
  const diasSemana = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
  sheet.getRange(1, 1, 1, 7).setValues([diasSemana]);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  sheet.getRange(1, 1, 1, 7).setHorizontalAlignment('center');
  sheet.getRange(1, 1, 1, 7).setBackground('#4285F4');
  sheet.getRange(1, 1, 1, 7).setFontColor('#FFFFFF');
  
  // Ajustar primer día (Lunes=0)
  let ajustePrimerDia = primerDiaSemana - 1;
  if (ajustePrimerDia < 0) ajustePrimerDia = 6;
  
  let fila = 2;
  let col = ajustePrimerDia;
  
  // Llenar días
  for (let dia = 1; dia <= ultimoDia; dia++) {
    // Número del día
    sheet.getRange(fila, col + 1).setValue(dia);
    sheet.getRange(fila, col + 1).setFontWeight('bold');
    sheet.getRange(fila, col + 1).setHorizontalAlignment('center');
    
    // MAÑANA
    sheet.getRange(fila + 1, col + 1).setValue('MAÑANA');
    crearDesplegableGuia(sheet, fila + 1, col + 1);
    
    // TARDE
    sheet.getRange(fila + 2, col + 1).setValue('TARDE');
    crearDesplegableGuia(sheet, fila + 2, col + 1);
    
    // Columnas ocultas
    sheet.getRange(fila + 1, 8).setValue('');
    sheet.getRange(fila + 1, 9).setValue('');
    sheet.getRange(fila + 2, 8).setValue('');
    sheet.getRange(fila + 2, 9).setValue('');
    
    col++;
    if (col === 7) {
      col = 0;
      fila += 3;
    }
  }
  
  // Ocultar columnas H e I
  sheet.hideColumns(8, 2);
  
  // Formatear
  for (let i = 1; i <= 7; i++) {
    sheet.setColumnWidth(i, 100);
  }
  
  // Proteger columnas ocultas
  const rangoProtegido = sheet.getRange(1, 8, sheet.getMaxRows(), 2);
  const proteccion = rangoProtegido.protect();
  proteccion.setDescription('Columnas de sistema - No editar');
}

/**
 * Crea desplegable de validación para celda de guía
 */
function crearDesplegableGuia(sheet, fila, columna) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const opciones = [VIS.NO_DISPONIBLE, VIS.LIBERAR];
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}

/**
 * Envía email de notificación al guía recién creado
 */
function enviarEmailCreacionGuia(nombre, email, urlCalendario) {
  const asunto = `Bienvenido al Sistema de Tours - Tu Calendario Personal`;
  
  const cuerpo = `Hola ${nombre},\n\n` +
                 `Se ha creado tu calendario personal para la gestión de tours.\n\n` +
                 `Puedes acceder a tu calendario en el siguiente enlace:\n` +
                 `${urlCalendario}\n\n` +
                 `Instrucciones:\n` +
                 `- Marca "NO DISPONIBLE" en los días/turnos que no puedas trabajar\n` +
                 `- Usa "LIBERAR" para volver a estar disponible\n` +
                 `- Cuando se te asigne un tour, verás "ASIGNADO M/T1/T2/T3"\n` +
                 `- NO modifiques las columnas ocultas del sistema\n\n` +
                 `Cualquier duda, contacta con el administrador.\n\n` +
                 `Saludos,\n` +
                 `Sistema de Gestión de Tours`;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      body: cuerpo
    });
    Logger.log(`Email de bienvenida enviado a ${email}`);
  } catch (error) {
    Logger.log(`Error enviando email de bienvenida: ${error.toString()}`);
  }
}
/**
 * Crea desplegable MAÑANA del Master en una celda específica
 */
function crearDesplegableMasterManana(sheet, fila, columna) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const opciones = [VIS.ASIGNAR_MANANA, VIS.LIBERAR_MASTER];
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}

/**
 * Crea desplegable TARDE del Master en una celda específica
 */
function crearDesplegableMasterTarde(sheet, fila, columna) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const opciones = [VIS.ASIGNAR_T1, VIS.ASIGNAR_T2, VIS.ASIGNAR_T3, VIS.LIBERAR_MASTER];
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}/**
 * FUNCIONES ADICIONALES PARA DESPLEGABLES
 * Agregar estas funciones al archivo CreateGuide.js existente
 */

/**
 * Crea desplegable MAÑANA del Master en una celda específica
 */
function crearDesplegableMasterManana(sheet, fila, columna) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const opciones = [VIS.ASIGNAR_MANANA, VIS.LIBERAR_MASTER];
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}

/**
 * Crea desplegable TARDE del Master en una celda específica
 */
function crearDesplegableMasterTarde(sheet, fila, columna) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const opciones = [VIS.ASIGNAR_T1, VIS.ASIGNAR_T2, VIS.ASIGNAR_T3, VIS.LIBERAR_MASTER];
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}

/**
 * Crea desplegable de validación para celda de guía
 */
function crearDesplegableGuia(sheet, fila, columna) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const opciones = [VIS.NO_DISPONIBLE, VIS.LIBERAR];
  
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(fila, columna).setDataValidation(regla);
}