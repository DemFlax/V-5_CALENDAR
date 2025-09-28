/**
 * SERVICIO DE EMAIL
 * Gestiona notificaciones automáticas del sistema
 */

class ServicioEmail {
  
  // Enviar notificación de asignación de tour
  static enviarAsignacionTour(emailGuia, fecha, turno, tipoTour) {
    try {
      const fechaFormateada = this.formatearFecha(fecha);
      const asunto = `🟢 Tour Asignado - ${fechaFormateada} ${turno}`;
      
      const cuerpo = `
Hola,

Se te ha asignado un nuevo tour:

📅 Fecha: ${fechaFormateada}
🕐 Turno: ${turno}
🎯 Tipo: ${tipoTour}

El tour aparecerá automáticamente en tu calendario en color verde.

¡Gracias por tu trabajo!

Equipo de Gestión Tours Madrid
      `;
      
      MailApp.sendEmail(emailGuia, asunto, cuerpo);
      Logger.log(`Email de asignación enviado a ${emailGuia}`);
      
    } catch (error) {
      Logger.log(`Error enviando email de asignación: ${error.message}`);
    }
  }
  
  // Enviar notificación de liberación de tour
  static enviarLiberacionTour(emailGuia, fecha, turno) {
    try {
      const fechaFormateada = this.formatearFecha(fecha);
      const asunto = `🔵 Tour Liberado - ${fechaFormateada} ${turno}`;
      
      const cuerpo = `
Hola,

Se ha liberado un tour que tenías asignado:

📅 Fecha: ${fechaFormateada}
🕐 Turno: ${turno}

Tu calendario se ha actualizado automáticamente y ahora muestras disponible para ese turno.

Equipo de Gestión Tours Madrid
      `;
      
      MailApp.sendEmail(emailGuia, asunto, cuerpo);
      Logger.log(`Email de liberación enviado a ${emailGuia}`);
      
    } catch (error) {
      Logger.log(`Error enviando email de liberación: ${error.message}`);
    }
  }
  
  // Enviar notificación de cambio de disponibilidad
  static enviarCambioDisponibilidad(emailGuia, fecha, turno, estado) {
    try {
      const fechaFormateada = this.formatearFecha(fecha);
      const emoji = estado === 'NO DISPONIBLE' ? '🔴' : '🟢';
      const accion = estado === 'NO DISPONIBLE' ? 'bloqueado' : 'liberado';
      
      const asunto = `${emoji} Disponibilidad ${accion} - ${fechaFormateada} ${turno}`;
      
      const cuerpo = `
Hola,

Has ${accion} tu disponibilidad:

📅 Fecha: ${fechaFormateada}
🕐 Turno: ${turno}
📊 Estado: ${estado}

El master calendar se ha actualizado automáticamente.

Equipo de Gestión Tours Madrid
      `;
      
      MailApp.sendEmail(emailGuia, asunto, cuerpo);
      Logger.log(`Email de cambio de disponibilidad enviado a ${emailGuia}`);
      
    } catch (error) {
      Logger.log(`Error enviando email de disponibilidad: ${error.message}`);
    }
  }

  // Enviar bienvenida a nuevo guía
  static enviarBienvenidaGuia(emailGuia, nombreGuia, codigo, urlCalendario) {
    try {
      const asunto = `🎉 Bienvenido al Equipo - Tours Madrid`;
      
      const cuerpo = `
Hola ${nombreGuia},

¡Bienvenido al equipo de Tours Madrid!

Tu información de acceso:
👤 Código: ${codigo}
📧 Email: ${emailGuia}
📅 Tu calendario: ${urlCalendario}

Instrucciones:
- Marca "NO DISPONIBLE" en los turnos que no puedas trabajar
- Usa "REVERTIR" para volver a estar disponible  
- Las asignaciones aparecerán automáticamente en verde

¡Esperamos trabajar contigo pronto!

Equipo de Gestión Tours Madrid
      `;
      
      MailApp.sendEmail(emailGuia, asunto, cuerpo);
      Logger.log(`Email de bienvenida enviado a ${emailGuia}`);
      
    } catch (error) {
      Logger.log(`Error enviando email de bienvenida: ${error.message}`);
    }
  }

  // Notificar al manager sobre cambios
  static notificarManager(asunto, mensaje) {
    try {
      const emailManager = ConfiguracionSistema.EMAIL_MANAGER;
      const cuerpo = `
📊 NOTIFICACIÓN DEL SISTEMA

${mensaje}

Hora: ${new Date().toLocaleString('es-ES')}

Sistema Tours Madrid
      `;
      
      MailApp.sendEmail(emailManager, `[SISTEMA] ${asunto}`, cuerpo);
      Logger.log(`Notificación enviada al manager: ${asunto}`);
      
    } catch (error) {
      Logger.log(`Error notificando al manager: ${error.message}`);
    }
  }

  // Enviar resumen diario al manager
  static enviarResumenDiario() {
    try {
      const guias = ConfiguracionSistema.getGuiasConfigurados();
      const fechaHoy = new Date();
      
      let resumen = `📊 RESUMEN DIARIO - ${this.formatearFecha(fechaHoy)}\n\n`;
      resumen += `👥 Guías activos: ${guias.length}\n`;
      resumen += `📅 Calendarios sincronizados: ${guias.length}\n\n`;
      
      resumen += `🔴 Disponibilidad de guías:\n`;
      // Aquí se podría agregar lógica para revisar disponibilidad del día
      
      guias.forEach(guia => {
        resumen += `• ${guia.codigo} (${guia.nombre}): Activo\n`;
      });
      
      this.notificarManager('Resumen Diario', resumen);
      
    } catch (error) {
      Logger.log(`Error enviando resumen diario: ${error.message}`);
    }
  }

  // Enviar alerta de conflicto
  static enviarAlertaConflicto(detallesConflicto) {
    try {
      const asunto = '⚠️ Conflicto Detectado en Calendario';
      const mensaje = `
Se ha detectado un conflicto en el sistema:

${detallesConflicto}

Revisa inmediatamente los calendarios para resolver el problema.
      `;
      
      this.notificarManager(asunto, mensaje);
      
    } catch (error) {
      Logger.log(`Error enviando alerta de conflicto: ${error.message}`);
    }
  }

  // Enviar notificación de sincronización
  static enviarNotificacionSincronizacion(numeroGuias, cambiosDetectados) {
    try {
      const mensaje = `
Sincronización completada:

👥 Guías sincronizados: ${numeroGuias}
🔄 Cambios aplicados: ${cambiosDetectados}
⏰ Hora: ${new Date().toLocaleString('es-ES')}

Todos los calendarios están actualizados.
      `;
      
      this.notificarManager('Sincronización Completada', mensaje);
      
    } catch (error) {
      Logger.log(`Error enviando notificación de sincronización: ${error.message}`);
    }
  }

  // Función auxiliar para formatear fechas
  static formatearFecha(fecha) {
    if (!(fecha instanceof Date)) {
      fecha = new Date(fecha);
    }
    
    const opciones = { 
      weekday: 'long', 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    };
    
    return fecha.toLocaleDateString('es-ES', opciones);
  }

  // Función auxiliar para validar email
  static validarEmail(email) {
    return ConfiguracionSistema.validarEmail(email);
  }

  // Test de email (para verificar configuración)
  static enviarEmailTest(emailDestino) {
    try {
      const asunto = '✅ Test de Email - Sistema Tours';
      const cuerpo = `
Este es un email de prueba del sistema de Tours Madrid.

Si recibes este mensaje, la configuración de email está funcionando correctamente.

Hora del test: ${new Date().toLocaleString('es-ES')}
      `;
      
      MailApp.sendEmail(emailDestino, asunto, cuerpo);
      return true;
      
    } catch (error) {
      Logger.log(`Error en test de email: ${error.message}`);
      return false;
    }
  }

  // Configurar plantillas de email personalizadas
  static configurarPlantillas() {
    return {
      asignacion: {
        asunto: '🟢 Tour Asignado - {fecha} {turno}',
        cuerpo: `
Hola,

Se te ha asignado un nuevo tour:

📅 Fecha: {fecha}
🕐 Turno: {turno}
🎯 Tipo: {tipo}

¡Gracias por tu trabajo!

Equipo Tours Madrid`
      },
      liberacion: {
        asunto: '🔵 Tour Liberado - {fecha} {turno}',
        cuerpo: `
Hola,

Se ha liberado un tour:

📅 Fecha: {fecha}
🕐 Turno: {turno}

Equipo Tours Madrid`
      }
    };
  }
}