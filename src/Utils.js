/**
 * FUNCIONES UTILITARIAS
 * Helpers y funciones auxiliares
 */

/**
 * Determina si un nombre de hoja es mensual (formato MM_YYYY o YYYY_MM)
 */
function esHojaMensual(nombreHoja) {
  // Patrones: 11_2025, 2025_11, 10_2025, etc.
  const patron1 = /^(\d{1,2})_(\d{4})$/;  // MM_YYYY
  const patron2 = /^(\d{4})_(\d{1,2})$/;  // YYYY_MM
  
  return patron1.test(nombreHoja) || patron2.test(nombreHoja);
}

/**
 * Extrae mes y año de un nombre de hoja
 */
function extraerMesAnioDeNombre(nombreHoja) {
  const patron1 = /^(\d{1,2})_(\d{4})$/;  // MM_YYYY
  const patron2 = /^(\d{4})_(\d{1,2})$/;  // YYYY_MM
  
  let match = nombreHoja.match(patron1);
  if (match) {
    return { mes: parseInt(match[1]), anio: parseInt(match[2]) };
  }
  
  match = nombreHoja.match(patron2);
  if (match) {
    return { mes: parseInt(match[2]), anio: parseInt(match[1]) };
  }
  
  return { mes: 1, anio: 2025 }; // Default
}

/**
 * Construye un mapa de columnas para cada guía en la Hoja Maestra
 */
function construirMapaColumnasGuias(encabezados, subEncabezados, guias) {
  const mapa = new Map();
  
  for (const guia of guias) {
    // Buscar el código del guía en los encabezados
    for (let col = 0; col < encabezados.length; col++) {
      const encabezado = encabezados[col];
      
      if (typeof encabezado === 'string' && encabezado.includes(guia.codigo)) {
        // Encontrado - las siguientes dos columnas son MAÑANA y TARDE
        const colManana = col;
        const colTarde = col + 1;
        
        // Verificar que los subencabezados coincidan
        if (subEncabezados[colManana] === 'MAÑANA' && 
            subEncabezados[colTarde] === 'TARDE') {
          mapa.set(guia, { manana: colManana, tarde: colTarde });
          break;
        }
      }
    }
  }
  
  return mapa;
}

/**
 * Determina el tipo de turno de tarde basado en el estado
 */
function determinarTipoTurnoTarde(estado) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  
  if (estado === VIS.ASIGNAR_T1 || estado === VIS.ASIGNADO_T1) {
    return 'T1';
  } else if (estado === VIS.ASIGNAR_T2 || estado === VIS.ASIGNADO_T2) {
    return 'T2';
  } else if (estado === VIS.ASIGNAR_T3 || estado === VIS.ASIGNADO_T3) {
    return 'T3';
  }
  
  return 'T1'; // Default
}

/**
 * Busca el turno de tarde de un guía para una fecha (cualquier T1/T2/T3)
 */
function buscarTurnoTarde(guia, fecha) {
  for (const tipo of ['T1', 'T2', 'T3']) {
    const turno = guia.obtenerTurno(fecha, tipo);
    if (turno) return turno;
  }
  return null;
}

/**
 * Obtiene el color de fondo para un estado
 */
function obtenerColorParaEstado(estado) {
  const VIS = CONFIG.ESTADOS_VISIBLES;
  const COL = CONFIG.COLORES;
  
  if (estado === VIS.NO_DISPONIBLE) {
    return COL.NO_DISPONIBLE;
  } else if (estado.startsWith('ASIGNADO')) {
    return COL.ASIGNADO;
  }
  
  return COL.DISPONIBLE;
}

/**
 * Aplica actualizaciones en lote a una hoja
 */
function aplicarActualizacionesEnLote(sheet, actualizaciones) {
  if (actualizaciones.length === 0) return;
  
  for (const act of actualizaciones) {
    sheet.getRange(act.fila, act.columna).setValue(act.valor);
  }
}

/**
 * Aplica formateos de color en lote
 */
function aplicarFormateosEnLote(sheet, formateos) {
  if (formateos.length === 0) return;
  
  for (const fmt of formateos) {
    sheet.getRange(fmt.fila, fmt.columna).setBackground(fmt.color);
  }
}

/**
 * Valida formato de email
 */
function esEmailValido(email) {
  const patron = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return patron.test(email);
}

/**
 * Genera fechas de un mes completo
 */
function generarFechasMes(mes, anio) {
  const fechas = [];
  const diasEnMes = new Date(anio, mes, 0).getDate();
  
  for (let dia = 1; dia <= diasEnMes; dia++) {
    fechas.push(new Date(anio, mes - 1, dia));
  }
  
  return fechas;
}

/**
 * Formatea una fecha como DD/MM/YYYY
 */
function formatearFecha(fecha) {
  const dia = fecha.getDate().toString().padStart(2, '0');
  const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
  const anio = fecha.getFullYear();
  return `${dia}/${mes}/${anio}`;
}