/**
 * CLASES DE DOMINIO
 * ClaseGuia y ClaseTurno para lógica de negocio en memoria
 */

class ClaseGuia {
  constructor(codigo, nombre, email, fileId) {
    this.codigo = codigo;
    this.nombre = nombre;
    this.email = email;
    this.fileId = fileId;
    this.turnos = new Map(); // Mapa de fecha+turno -> ClaseTurno
  }
  
  /**
   * Agrega un turno a este guía
   */
  agregarTurno(fecha, tipoTurno, turno) {
    const key = `${fecha.getTime()}_${tipoTurno}`;
    this.turnos.set(key, turno);
  }
  
  /**
   * Obtiene un turno específico
   */
  obtenerTurno(fecha, tipoTurno) {
    const key = `${fecha.getTime()}_${tipoTurno}`;
    return this.turnos.get(key);
  }
  
  /**
   * Obtiene todos los turnos
   */
  obtenerTodosTurnos() {
    return Array.from(this.turnos.values());
  }
}

class ClaseTurno {
  constructor(fecha, tipoTurno, guia) {
    this.fecha = fecha;
    this.tipoTurno = tipoTurno; // 'MANANA', 'T1', 'T2', 'T3'
    this.guia = guia;
    
    // Estados desde ambas hojas
    this.estadoMaster = '';
    this.estadoGuia = '';
    this.lockStatusGuia = '';
    this.timestampGuia = null;
    this.timestampMaster = null;
    
    // Estado resuelto después de aplicar reglas
    this.estadoFinal = '';
    this.lockStatusFinal = '';
    this.requiereActualizacionMaster = false;
    this.requiereActualizacionGuia = false;
    this.requiereNotificacion = false;
  }
  
  /**
   * Aplica la lógica de resolución de conflictos
   * CORREGIDO: Respeta jerarquía Master > Guía en bloqueos activos
   */
  resolverEstado() {
    const LOCK = CONFIG.LOCK_STATUS;
    const VIS = CONFIG.ESTADOS_VISIBLES;
    
    // CASO 1: Guía intenta marcar NO DISPONIBLE
    if (this.estadoGuia === VIS.NO_DISPONIBLE) {
      // CORRECCIÓN CRÍTICA: Si hay un bloqueo Master activo (M-A*), 
      // el Guía NO puede marcar NO DISPONIBLE
      if (this.lockStatusGuia && this.lockStatusGuia.startsWith('M-A')) {
        // Mantener la asignación del Master, ignorar intento del Guía
        this.estadoFinal = this._getEstadoVisibleParaLock(this.lockStatusGuia);
        this.lockStatusFinal = this.lockStatusGuia;
        // Actualizar guía para revertir su intento
        this.requiereActualizacionGuia = true;
        Logger.log(`⚠️ Guía ${this.guia.codigo} intentó marcar NO DISPONIBLE sobre asignación Master - IGNORADO`);
        return;
      }
      
      // Si no hay bloqueo Master, procesar el NO DISPONIBLE del Guía
      if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        // Ya está procesado
        this.estadoFinal = VIS.NO_DISPONIBLE;
        this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
        this.requiereActualizacionMaster = (this.estadoMaster !== VIS.NO_DISPONIBLE);
      } else {
        // Nueva marca de NO DISPONIBLE
        this.estadoFinal = VIS.NO_DISPONIBLE;
        this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
        Logger.log(`✓ Guía ${this.guia.codigo} marcó NO DISPONIBLE - ${this.fecha.toDateString()}`);
      }
      return;
    }
    
    // CASO 2: Master asigna un turno
    const asignacionesMaster = [VIS.ASIGNAR_MANANA, VIS.ASIGNAR_T1, VIS.ASIGNAR_T2, VIS.ASIGNAR_T3];
    if (asignacionesMaster.includes(this.estadoMaster)) {
      const lockCorrespondiente = this._getLockParaAsignacion(this.estadoMaster);
      
      // Si el turno ya está asignado con este lock, no hacer nada
      if (this.lockStatusGuia === lockCorrespondiente) {
        this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
        this.lockStatusFinal = lockCorrespondiente;
      } else if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        // CONFLICTO: Master intenta asignar pero guía marcó NO DISPONIBLE
        // Aplicar First Wins: comparar timestamps
        
        if (this.timestampMaster && this.timestampGuia) {
          if (this.timestampMaster < this.timestampGuia) {
            // GANA MASTER (actuó primero)
            this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
            this.lockStatusFinal = lockCorrespondiente;
            this.requiereActualizacionMaster = true;
            this.requiereActualizacionGuia = true;
            this.requiereNotificacion = true;
            Logger.log(`⚡ CONFLICTO: Master gana (First Wins) - ${this.fecha.toDateString()}`);
          } else {
            // GANA GUÍA (actuó primero)
            this.estadoFinal = VIS.NO_DISPONIBLE;
            this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
            this.requiereActualizacionMaster = true;
            Logger.log(`⚡ CONFLICTO: Guía gana (First Wins) - ${this.fecha.toDateString()}`);
          }
        } else {
          // Sin timestamps válidos, asumir que Master gana (ejecución actual)
          this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
          this.lockStatusFinal = lockCorrespondiente;
          this.requiereActualizacionMaster = true;
          this.requiereActualizacionGuia = true;
          this.requiereNotificacion = true;
          Logger.log(`⚡ CONFLICTO: Master gana (sin timestamps) - ${this.fecha.toDateString()}`);
        }
      } else {
        // Asignación nueva o reasignación
        this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
        this.lockStatusFinal = lockCorrespondiente;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
        this.requiereNotificacion = true;
        Logger.log(`✓ Master asignó turno a ${this.guia.codigo} - ${this.fecha.toDateString()}`);
      }
      return;
    }
    
    // CASO 3: Master libera un turno
    if (this.estadoMaster === VIS.LIBERAR_MASTER) {
      // CORRECCIÓN: Verificar qué tipo de bloqueo hay
      if (this.lockStatusGuia && this.lockStatusGuia.startsWith('M-A')) {
        // Liberar asignación del Master
        this.estadoFinal = this._getEstadoInicial();
        this.lockStatusFinal = LOCK.LIBERADO_MASTER;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
        Logger.log(`✓ Master liberó turno de ${this.guia.codigo} - ${this.fecha.toDateString()}`);
      } else if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        // NO se puede liberar un NO DISPONIBLE del Guía desde el Master
        this.estadoFinal = VIS.NO_DISPONIBLE;
        this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
        Logger.log(`⚠️ Master intentó liberar NO DISPONIBLE del Guía - IGNORADO`);
        // No actualizar nada, mantener el bloqueo del Guía
      } else {
        // Ya estaba libre
        this.estadoFinal = this.estadoGuia || this._getEstadoInicial();
        this.lockStatusFinal = this.lockStatusGuia || LOCK.VACIO;
      }
      return;
    }
    
    // CASO 4: Guía libera un turno
    if (this.estadoGuia === VIS.LIBERAR) {
      if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        // Liberar NO DISPONIBLE del Guía
        this.estadoFinal = this._getEstadoInicial();
        this.lockStatusFinal = LOCK.LIBERADO_GUIA;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
        Logger.log(`✓ Guía ${this.guia.codigo} liberó NO DISPONIBLE - ${this.fecha.toDateString()}`);
      } else if (this.lockStatusGuia && this.lockStatusGuia.startsWith('M-A')) {
        // NO puede liberar una asignación del Master
        this.estadoFinal = this._getEstadoVisibleParaLock(this.lockStatusGuia);
        this.lockStatusFinal = this.lockStatusGuia;
        this.requiereActualizacionGuia = true;
        Logger.log(`⚠️ Guía intentó liberar asignación del Master - IGNORADO`);
      } else {
        // Ya estaba libre
        this.estadoFinal = this._getEstadoInicial();
        this.lockStatusFinal = LOCK.VACIO;
        this.requiereActualizacionGuia = true;
      }
      return;
    }
    
    // CASO 5: Estados estables - propagar desde guía a master si es necesario
    if (this.lockStatusGuia) {
      this.estadoFinal = this._getEstadoVisibleParaLock(this.lockStatusGuia);
      this.lockStatusFinal = this.lockStatusGuia;
      
      // Verificar si master necesita actualización
      const estadoEsperadoMaster = this._getEstadoMasterParaLock(this.lockStatusGuia);
      if (this.estadoMaster !== estadoEsperadoMaster) {
        this.requiereActualizacionMaster = true;
      }
    } else {
      // Estado inicial
      this.estadoFinal = this._getEstadoInicial();
      this.lockStatusFinal = LOCK.VACIO;
    }
  }
  
  /**
   * Helpers privados
   */
  _getLockParaAsignacion(estadoMaster) {
    const VIS = CONFIG.ESTADOS_VISIBLES;
    const LOCK = CONFIG.LOCK_STATUS;
    
    switch(estadoMaster) {
      case VIS.ASIGNAR_MANANA: return LOCK.MASTER_ASIGNADO_MANANA;
      case VIS.ASIGNAR_T1: return LOCK.MASTER_ASIGNADO_T1;
      case VIS.ASIGNAR_T2: return LOCK.MASTER_ASIGNADO_T2;
      case VIS.ASIGNAR_T3: return LOCK.MASTER_ASIGNADO_T3;
      default: return LOCK.VACIO;
    }
  }
  
  _getEstadoVisibleParaLock(lockStatus) {
    const VIS = CONFIG.ESTADOS_VISIBLES;
    const LOCK = CONFIG.LOCK_STATUS;
    
    switch(lockStatus) {
      case LOCK.GUIA_NO_DISPONIBLE: return VIS.NO_DISPONIBLE;
      case LOCK.MASTER_ASIGNADO_MANANA: return VIS.ASIGNADO_M;
      case LOCK.MASTER_ASIGNADO_T1: return VIS.ASIGNADO_T1;
      case LOCK.MASTER_ASIGNADO_T2: return VIS.ASIGNADO_T2;
      case LOCK.MASTER_ASIGNADO_T3: return VIS.ASIGNADO_T3;
      default: return this._getEstadoInicial();
    }
  }
  
  _getEstadoMasterParaLock(lockStatus) {
    const VIS = CONFIG.ESTADOS_VISIBLES;
    const LOCK = CONFIG.LOCK_STATUS;
    
    switch(lockStatus) {
      case LOCK.GUIA_NO_DISPONIBLE: return VIS.NO_DISPONIBLE;
      case LOCK.MASTER_ASIGNADO_MANANA: return VIS.ASIGNADO_M;
      case LOCK.MASTER_ASIGNADO_T1: return VIS.ASIGNADO_T1;
      case LOCK.MASTER_ASIGNADO_T2: return VIS.ASIGNADO_T2;
      case LOCK.MASTER_ASIGNADO_T3: return VIS.ASIGNADO_T3;
      default: return '';
    }
  }
  
  _getEstadoInicial() {
    return this.tipoTurno === 'MANANA' ? 
      CONFIG.ESTADOS_VISIBLES.MANANA_INICIAL : 
      CONFIG.ESTADOS_VISIBLES.TARDE_INICIAL;
  }
}