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
    this.turnos = new Map();
  }
  
  agregarTurno(fecha, tipoTurno, turno) {
    const key = `${fecha.getTime()}_${tipoTurno}`;
    this.turnos.set(key, turno);
  }
  
  obtenerTurno(fecha, tipoTurno) {
    const key = `${fecha.getTime()}_${tipoTurno}`;
    return this.turnos.get(key);
  }
  
  obtenerTodosTurnos() {
    return Array.from(this.turnos.values());
  }
}

class ClaseTurno {
  constructor(fecha, tipoTurno, guia) {
    this.fecha = fecha;
    this.tipoTurno = tipoTurno;
    this.guia = guia;
    
    this.estadoMaster = '';
    this.estadoGuia = '';
    this.lockStatusGuia = '';
    this.timestampGuia = null;
    this.timestampMaster = null;
    
    this.estadoFinal = '';
    this.lockStatusFinal = '';
    this.requiereActualizacionMaster = false;
    this.requiereActualizacionGuia = false;
    this.requiereNotificacion = false;
  }
  
  resolverEstado() {
    const LOCK = CONFIG.LOCK_STATUS;
    const VIS = CONFIG.ESTADOS_VISIBLES;
    
    if (this.estadoGuia === VIS.NO_DISPONIBLE) {
      if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        this.estadoFinal = VIS.NO_DISPONIBLE;
        this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
        this.requiereActualizacionMaster = (this.estadoMaster !== VIS.NO_DISPONIBLE);
      } else {
        this.estadoFinal = VIS.NO_DISPONIBLE;
        this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
      }
      return;
    }
    
    const asignacionesMaster = [VIS.ASIGNAR_MANANA, VIS.ASIGNAR_T1, VIS.ASIGNAR_T2, VIS.ASIGNAR_T3];
    if (asignacionesMaster.includes(this.estadoMaster)) {
      const lockCorrespondiente = this._getLockParaAsignacion(this.estadoMaster);
      
      if (this.lockStatusGuia === lockCorrespondiente) {
        this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
        this.lockStatusFinal = lockCorrespondiente;
        return;
      } 
      
      if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        if (this.timestampMaster && this.timestampGuia) {
          if (this.timestampMaster < this.timestampGuia) {
            this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
            this.lockStatusFinal = lockCorrespondiente;
            this.requiereActualizacionMaster = true;
            this.requiereActualizacionGuia = true;
            this.requiereNotificacion = true;
          } else {
            this.estadoFinal = VIS.NO_DISPONIBLE;
            this.lockStatusFinal = LOCK.GUIA_NO_DISPONIBLE;
            this.requiereActualizacionMaster = true;
          }
        } else {
          this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
          this.lockStatusFinal = lockCorrespondiente;
          this.requiereActualizacionMaster = true;
          this.requiereActualizacionGuia = true;
          this.requiereNotificacion = true;
        }
        return;
      }
      
      this.estadoFinal = this._getEstadoVisibleParaLock(lockCorrespondiente);
      this.lockStatusFinal = lockCorrespondiente;
      this.requiereActualizacionMaster = true;
      this.requiereActualizacionGuia = true;
      this.requiereNotificacion = true;
      return;
    }
    
    if (this.estadoMaster === VIS.LIBERAR_MASTER) {
      if (this.lockStatusGuia.startsWith('M-A')) {
        this.estadoFinal = this._getEstadoInicial();
        this.lockStatusFinal = LOCK.LIBERADO_MASTER;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
      } else {
        this.estadoFinal = this.estadoGuia || this._getEstadoInicial();
        this.lockStatusFinal = this.lockStatusGuia || LOCK.VACIO;
      }
      return;
    }
    
    if (this.estadoGuia === VIS.LIBERAR) {
      if (this.lockStatusGuia === LOCK.GUIA_NO_DISPONIBLE) {
        this.estadoFinal = this._getEstadoInicial();
        this.lockStatusFinal = LOCK.LIBERADO_GUIA;
        this.requiereActualizacionMaster = true;
        this.requiereActualizacionGuia = true;
      } else if (this.lockStatusGuia.startsWith('M-A')) {
        this.estadoFinal = this._getEstadoVisibleParaLock(this.lockStatusGuia);
        this.lockStatusFinal = this.lockStatusGuia;
      } else {
        this.estadoFinal = this._getEstadoInicial();
        this.lockStatusFinal = LOCK.VACIO;
        this.requiereActualizacionGuia = true;
      }
      return;
    }
    
    if (this.lockStatusGuia) {
      this.estadoFinal = this._getEstadoVisibleParaLock(this.lockStatusGuia);
      this.lockStatusFinal = this.lockStatusGuia;
      
      const estadoEsperadoMaster = this._getEstadoMasterParaLock(this.lockStatusGuia);
      if (this.estadoMaster !== estadoEsperadoMaster) {
        this.requiereActualizacionMaster = true;
      }
    } else {
      this.estadoFinal = this._getEstadoInicial();
      this.lockStatusFinal = LOCK.VACIO;
    }
  }
  
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