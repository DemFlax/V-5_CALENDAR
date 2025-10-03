/**
 * CONFIGURACIÓN GLOBAL DEL SISTEMA
 * Constantes y configuración centralizada
 */

const CONFIG = {
  SHEET_REGISTRO: 'REGISTRO',
  
  CARPETA_GUIAS_ID: '1qZjy2Syg5Ag-zOti9N9gj5IFHr9vP3J4',
  
  REGISTRO_COL: {
    TIMESTAMP: 0,
    CODIGO: 1,
    NOMBRE: 2,
    EMAIL: 3,
    FILE_ID: 4,
    URL: 5
  },
  
  GUIA_CAL: {
    DIAS_SEMANA: 7,
    FILAS_POR_DIA: 3,
    COL_LOCK_STATUS: 7,
    COL_TIMESTAMP: 8,
    ROW_OFFSET_MANANA: 1,
    ROW_OFFSET_TARDE: 2
  },
  
  LOCK_STATUS: {
    GUIA_NO_DISPONIBLE: 'G-ND',
    MASTER_ASIGNADO_MANANA: 'M-AM',
    MASTER_ASIGNADO_T1: 'M-AT1',
    MASTER_ASIGNADO_T2: 'M-AT2',
    MASTER_ASIGNADO_T3: 'M-AT3',
    LIBERADO_GUIA: 'L-G',
    LIBERADO_MASTER: 'L-M',
    VACIO: ''
  },
  
  ESTADOS_VISIBLES: {
    MANANA_INICIAL: 'MAÑANA',
    TARDE_INICIAL: 'TARDE',
    NO_DISPONIBLE: 'NO DISPONIBLE',
    LIBERAR: 'LIBERAR',
    ASIGNADO_M: 'ASIGNADO M',
    ASIGNADO_T1: 'ASIGNADO T1',
    ASIGNADO_T2: 'ASIGNADO T2',
    ASIGNADO_T3: 'ASIGNADO T3',
    ASIGNAR_MANANA: 'ASIGNAR MAÑANA',
    ASIGNAR_T1: 'ASIGNAR T1',
    ASIGNAR_T2: 'ASIGNAR T2',
    ASIGNAR_T3: 'ASIGNAR T3',
    LIBERAR_MASTER: 'LIBERAR'
  },
  
  COLORES: {
    DISPONIBLE: '#FFFFFF',
    NO_DISPONIBLE: '#FF0000',
    ASIGNADO: '#00FF00'
  },
  
  HORARIOS: {
    MANANA: { hora: 12, minuto: 15 },
    T1: { hora: 17, minuto: 15 },
    T2: { hora: 18, minuto: 15 },
    T3: { hora: 19, minuto: 15 }
  },
  
  PROP_CALENDAR_ID: 'MASTER_CALENDAR_ID',
  
  TRIGGER_FUNCTION: 'ejecutarSincronizacion',
  TRIGGER_INTERVAL_MINUTES: 1
};

/**
 * Obtiene el ID del calendario maestro desde las propiedades del script
 */
function obtenerCalendarIdMaestro() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(CONFIG.PROP_CALENDAR_ID);
}

/**
 * Establece el ID del calendario maestro
 */
function establecerCalendarIdMaestro(calendarId) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.PROP_CALENDAR_ID, calendarId);
}