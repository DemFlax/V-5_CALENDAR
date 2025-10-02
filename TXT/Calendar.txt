/**
 * INTEGRACIÓN CON GOOGLE CALENDAR
 * Gestión de eventos y invitados
 */

/**
 * Agrega un guía como invitado a un evento existente
 */
function agregarGuiaAEvento(calendarId, fecha, tipoTurno, emailGuia) {
  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      Logger.log(`Error: Calendario ${calendarId} no encontrado`);
      return;
    }
    
    // Construir fecha/hora del evento
    const horario = obtenerHorarioParaTurno(tipoTurno);
    const fechaEvento = new Date(
      fecha.getFullYear(),
      fecha.getMonth(),
      fecha.getDate(),
      horario.hora,
      horario.minuto
    );
    
    // Buscar evento por fecha/hora exacta
    const fechaFin = new Date(fechaEvento.getTime() + 60000); // +1 minuto para búsqueda
    const eventos = calendar.getEvents(fechaEvento, fechaFin);
    
    if (eventos.length === 0) {
      Logger.log(`No se encontró evento para ${fecha.toDateString()} a las ${horario.hora}:${horario.minuto}`);
      return;
    }
    
    // Tomar el primer evento encontrado (asumimos uno por horario)
    const evento = eventos[0];
    
    // Verificar si ya es invitado
    const invitados = evento.getGuestList();
    const yaEsInvitado = invitados.some(inv => inv.getEmail() === emailGuia);
    
    if (!yaEsInvitado) {
      evento.addGuest(emailGuia);
      Logger.log(`Guía ${emailGuia} agregado a evento del ${fecha.toDateString()}`);
    } else {
      Logger.log(`Guía ${emailGuia} ya es invitado del evento`);
    }
    
  } catch (error) {
    Logger.log(`Error agregando guía a evento: ${error.toString()}`);
  }
}

/**
 * Obtiene el horario correspondiente a un tipo de turno
 */
function obtenerHorarioParaTurno(tipoTurno) {
  switch(tipoTurno) {
    case 'MANANA':
      return CONFIG.HORARIOS.MANANA;
    case 'T1':
      return CONFIG.HORARIOS.T1;
    case 'T2':
      return CONFIG.HORARIOS.T2;
    case 'T3':
      return CONFIG.HORARIOS.T3;
    default:
      return CONFIG.HORARIOS.T1; // Default a T1
  }
}

/**
 * Envía notificación de asignación por email
 */
function enviarNotificacionAsignacion(guia, turno) {
  try {
    const fecha = Utilities.formatDate(turno.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const turnoTexto = obtenerTextoTurno(turno.tipoTurno);
    
    const asunto = `Asignación de Tour - ${fecha}`;
    const cuerpo = `Hola ${guia.nombre},\n\n` +
                   `Se te ha asignado el siguiente tour:\n\n` +
                   `Fecha: ${fecha}\n` +
                   `Turno: ${turnoTexto}\n\n` +
                   `Recibirás una invitación de calendario por separado.\n\n` +
                   `Saludos,\n` +
                   `Sistema de Gestión de Tours`;
    
    MailApp.sendEmail({
      to: guia.email,
      subject: asunto,
      body: cuerpo
    });
    
    Logger.log(`Notificación enviada a ${guia.email}`);
    
  } catch (error) {
    Logger.log(`Error enviando notificación: ${error.toString()}`);
  }
}

/**
 * Obtiene texto descriptivo del turno
 */
function obtenerTextoTurno(tipoTurno) {
  const horario = obtenerHorarioParaTurno(tipoTurno);
  const hora12 = horario.hora > 12 ? horario.hora - 12 : horario.hora;
  const ampm = horario.hora >= 12 ? 'PM' : 'AM';
  const minutos = horario.minuto.toString().padStart(2, '0');
  
  switch(tipoTurno) {
    case 'MANANA':
      return `Mañana (${hora12}:${minutos} ${ampm})`;
    case 'T1':
      return `Tarde T1 (${hora12}:${minutos} ${ampm})`;
    case 'T2':
      return `Tarde T2 (${hora12}:${minutos} ${ampm})`;
    case 'T3':
      return `Tarde T3 (${hora12}:${minutos} ${ampm})`;
    default:
      return `Turno (${hora12}:${minutos} ${ampm})`;
  }
}