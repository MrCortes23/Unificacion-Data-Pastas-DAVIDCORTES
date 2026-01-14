
//FUNCIONES DE ENVIOS DE CORREOS

/**
 * Envía correo de notificación al líder responsable - VERSIÓN CORREGIDA
 */
function sendEmailToLeader(leaderInfo, formData, reportId) {
  try {
    console.log(`📧 Intentando enviar correo para reporte ${reportId}`);
    console.log(`👤 Información del líder:`, leaderInfo);

    // Validación más robusta del email
    if (!leaderInfo || !leaderInfo.email) {
      console.warn('⚠️ No hay información del líder o email está vacío');
      return false;
    }

    const email = leaderInfo.email.trim();

    // Validación básica de formato de email
    if (!email || email === '' || !email.includes('@')) {
      console.warn(`⚠️ Email inválido: "${email}"`);
      return false;
    }

    console.log(`✅ Email válido detectado: ${email}`);

    const subject = `🚨 Nuevo Reporte N2 Asignado - ${reportId}`;

    // Formatear fecha de solución con manejo de errores
    let fechaSolucionFormateada = 'No especificada';
    try {
      const fechaSolucion = new Date(formData.fechaSolucion);
      if (!isNaN(fechaSolucion.getTime())) {
        fechaSolucionFormateada = Utilities.formatDate(fechaSolucion, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
    } catch (dateError) {
      console.warn('⚠️ Error formateando fecha:', dateError);
    }

    // Formatear fecha de reporte
    let fechaReporteFormateada = 'No especificada';
    try {
      const fechaReporte = new Date(formData.fecha);
      if (!isNaN(fechaReporte.getTime())) {
        fechaReporteFormateada = Utilities.formatDate(fechaReporte, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      }
    } catch (dateError) {
      console.warn('⚠️ Error formateando fecha de reporte:', dateError);
    }

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
        <div style="background-color: #d9534f; color: white; padding: 15px; border-radius: 8px 8px 0 0; text-align: center;">
          <h2 style="margin: 0;">Notificación de Reporte N2</h2>
        </div>
        
        <div style="padding: 20px; background-color: #f8f9fa; border-radius: 0 0 8px 8px;">
          <p>Hola <strong>${leaderInfo.nombre || 'Líder Responsable'}</strong>,</p>
          <p>Se le ha asignado un nuevo reporte N2 que requiere su atención.</p>
          
          <div style="background-color: white; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #d9534f;">
            <h3 style="margin-top: 0; color: #d9534f;">Detalles del Reporte</h3>
            
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">ID del Reporte:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>${reportId}</strong></td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Proceso:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${formData.proceso || 'No especificado'}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Zona/Proceso:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${formData.zonaProceso || 'No especificado'}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Anormalidad:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${formData.anormalidad || 'No especificado'}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Proceso Responsable:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${formData.procesoResponsable || 'No especificado'}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Fecha Límite Solución:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;"><strong>${fechaSolucionFormateada}</strong></td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Reportado por:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${formData.nombreCedula || 'No especificado'}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Fecha de Reporte:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${fechaReporteFormateada}</td>
              </tr>
            </table>
          </div>
          
          <div style="background-color: #fff3cd; padding: 15px; border-radius: 5px; border-left: 4px solid #ffc107; margin: 15px 0;">
            <p style="margin: 0;"><strong>Acción requerida:</strong> Por favor revisar este reporte y tomar las acciones correspondientes en el sistema.</p>
          </div>
          
          <p>Puede acceder al sistema para ver más detalles y actualizar el estado del reporte.</p>
          
          <div style="text-align: center; margin: 20px 0;">
            <a href="${ScriptApp.getService().getUrl()}" 
               style="background-color: #d9534f; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; display: inline-block;">
              Acceder al Sistema
            </a>
          </div>
        </div>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; text-align: center;">
          <p style="margin: 0; font-size: 12px; color: #6c757d;">
            Este es un mensaje automático generado por el Sistema de Reportes N2.<br>
            Por favor no responder directamente a este correo.
          </p>
        </div>
      </div>
    `;

    const plainBody = `
NOTIFICACIÓN DE REPORTE N2

Hola ${leaderInfo.nombre || 'Líder Responsable'},

Se le ha asignado un nuevo reporte N2 que requiere su atención.

DETALLES DEL REPORTE:
- ID del Reporte: ${reportId}
- Proceso: ${formData.proceso || 'No especificado'}
- Zona/Proceso: ${formData.zonaProceso || 'No especificado'}
- Anormalidad: ${formData.anormalidad || 'No especificado'}
- Proceso Responsable: ${formData.procesoResponsable || 'No especificado'}
- Fecha Límite Solución: ${fechaSolucionFormateada}
- Reportado por: ${formData.nombreCedula || 'No especificado'}
- Fecha de Reporte: ${fechaReporteFormateada}

ACCIÓN REQUERIDA: Por favor revisar este reporte y tomar las acciones correspondientes en el sistema.

Puede acceder al sistema en: ${ScriptApp.getService().getUrl()}

Este es un mensaje automático. Por favor no responder directamente a este correo.
    `;

    console.log(`✉️ Enviando correo a: ${email}`);

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });

    console.log(`✅ Correo enviado exitosamente a: ${email}`);
    return true;

  } catch (emailError) {
    console.error(`❌ Error enviando correo: ${emailError.message}`);
    console.error(`Stack trace: ${emailError.stack}`);
    return false;
  }
}

/**
 * Envía correo de confirmación al creador de la tarjeta
 */
function sendEmailToCreador(creadorEmail, data, fotosLinks) {
  try {
    const subject = ` Tarjeta de Anormalidad Creada - ${data.prioridad}`;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
        <div style="background-color: #28a745; color: white; padding: 15px; border-radius: 8px 8px 0 0; text-align: center;">
          <h2 style="margin: 0;">Confirmación de Tarjeta de Anormalidad</h2>
        </div>
        
        <div style="padding: 20px; background-color: #f8f9fa; border-radius: 0 0 8px 8px;">
          <p>Hola <strong>${data.nombreCedula}</strong>,</p>
          <p>Su tarjeta de anormalidad ha sido registrada exitosamente en el sistema.</p>
          
          <div style="background-color: white; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid #28a745;">
            <h3 style="margin-top: 0; color: #28a745;">Detalles de la Tarjeta</h3>
            
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">Zona de Riesgo:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.zonaRiesgo}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Ubicación:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.ubicacion}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Prioridad:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">
                  <span style="color: ${data.prioridad === 'Alta' ? '#dc3545' :
        data.prioridad === 'Media' ? '#fd7e14' : '#28a745'
      }; font-weight: bold;">${data.prioridad}</span>
                </td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Descripción:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.descripcionProblema}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Tipo de Riesgo:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.tipoRiesgo}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Responsable Asignado:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.responsableSolucion}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Fotos Adjuntas:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${fotosLinks.length} imagen(es)</td>
              </tr>
            </table>
          </div>
          
          <div style="background-color: #d1ecf1; padding: 15px; border-radius: 5px; border-left: 4px solid #17a2b8; margin: 15px 0;">
            <p style="margin: 0;"><strong>Estado:</strong> La tarjeta ha sido asignada a <strong>${data.responsableSolucion}</strong> para su revisión y solución.</p>
          </div>
          
          <p>Puede dar seguimiento a esta tarjeta accediendo al sistema.</p>
          
          <div style="text-align: center; margin: 20px 0;">
            <a href="${ScriptApp.getService().getUrl()}" 
               style="background-color: #28a745; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; display: inline-block;">
              Ver en el Sistema
            </a>
          </div>
        </div>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; text-align: center;">
          <p style="margin: 0; font-size: 12px; color: #6c757d;">
            Este es un mensaje automático del Sistema de Tarjetas de Anormalidad.
          </p>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: creadorEmail,
      subject: subject,
      htmlBody: htmlBody
    });

    console.log(`✅ Correo de confirmación enviado al creador: ${creadorEmail}`);

  } catch (emailError) {
    console.error(`❌ Error enviando correo al creador: ${emailError}`);
  }
}

/**
 * Envía correo de notificación al responsable asignado
 */
function sendEmailToResponsable(responsableEmail, data, fotosLinks, creadorEmail) {
  try {
    const subject = `🚨 Nueva Tarjeta de Anormalidad Asignada - ${data.prioridad}`;

    // Determinar color según prioridad
    const colorPrioridad = data.prioridad === 'Alta' ? '#dc3545' :
      data.prioridad === 'Media' ? '#fd7e14' : '#ffc107';

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
        <div style="background-color: ${colorPrioridad}; color: white; padding: 15px; border-radius: 8px 8px 0 0; text-align: center;">
          <h2 style="margin: 0;">Tarjeta de Anormalidad Asignada</h2>
        </div>
        
        <div style="padding: 20px; background-color: #f8f9fa; border-radius: 0 0 8px 8px;">
          <p>Estimado <strong>${data.responsableSolucion}</strong>,</p>
          <p>Se le ha asignado una nueva tarjeta de anormalidad que requiere su atención.</p>
          
          <div style="background-color: white; padding: 15px; border-radius: 5px; margin: 15px 0; border-left: 4px solid ${colorPrioridad};">
            <h3 style="margin-top: 0; color: ${colorPrioridad};">Detalles de la Tarjeta</h3>
            
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold; width: 40%;">Prioridad:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">
                  <span style="color: ${colorPrioridad}; font-weight: bold;">${data.prioridad}</span>
                </td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Zona de Riesgo:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.zonaRiesgo}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Ubicación:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.ubicacion}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Descripción del Problema:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.descripcionProblema}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Tipo de Riesgo:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.tipoRiesgo}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Reportado por:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.nombreCedula}</td>
              </tr>
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Fotos Adjuntas:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${fotosLinks.length} imagen(es)</td>
              </tr>
              ${data.generadaPor ? `
              <tr>
                <td style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Generada por:</td>
                <td style="padding: 8px; border-bottom: 1px solid #eee;">${data.generadaPor}</td>
              </tr>
              ` : ''}
            </table>
          </div>
          
          <div style="background-color: #fff3cd; padding: 15px; border-radius: 5px; border-left: 4px solid #ffc107; margin: 15px 0;">
            <p style="margin: 0;"><strong>Acción requerida:</strong> Por favor revisar esta anormalidad reportada y tomar las acciones correspondientes.</p>
          </div>
          
          ${fotosLinks.length > 0 ? `
          <div style="margin: 15px 0;">
            <h4>📸 Fotos adjuntas:</h4>
            <div style="display: flex; gap: 10px; flex-wrap: wrap;">
              ${fotosLinks.map(link => `
                <a href="${link}" target="_blank" style="display: inline-block;">
                  <img src="${link}" style="width: 100px; height: 100px; object-fit: cover; border-radius: 5px; border: 1px solid #ddd;">
                </a>
              `).join('')}
            </div>
          </div>
          ` : ''}
          
          <div style="text-align: center; margin: 20px 0;">
            <a href="${ScriptApp.getService().getUrl()}" 
               style="background-color: ${colorPrioridad}; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; display: inline-block;">
              Acceder al Sistema
            </a>
          </div>
        </div>
        
        <div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; text-align: center;">
          <p style="margin: 0; font-size: 12px; color: #6c757d;">
            Este es un mensaje automático del Sistema de Tarjetas de Anormalidad.
          </p>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: responsableEmail,
      subject: subject,
      htmlBody: htmlBody
    });

    console.log(`✅ Correo de notificación enviado al responsable: ${responsableEmail}`);

  } catch (emailError) {
    console.error(`❌ Error enviando correo al responsable: ${emailError}`);
  }
}

// Función para programar el envío de correos después de 10 segundos
function programarEnvioCorreos(fila, data, fotosLinks) {
  try {
    // Guardar los datos en Properties
    PropertiesService.getScriptProperties()
      .setProperty('EMAIL_DATA_' + fila, JSON.stringify({
        fila: fila,
        data: data,
        fotosLinks: fotosLinks
      }));
    
    // Crear trigger para ejecutar después de 10 segundos
    ScriptApp.newTrigger('enviarCorreoConRetraso')
      .timeBased()
      .after(5000) // 10 segundos
      .create();
    
    console.log(`Correo programado para fila ${fila} (en 10 segundos)`);
    
  } catch (error) {
    console.error('Error al programar correo:', error);
  }
}

function enviarCorreoConRetraso() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);
    
    // Buscar todas las tareas pendientes
    const allProps = properties.getProperties();
    
    for (const key in allProps) {
      if (key.startsWith('EMAIL_DATA_')) {
        try {
          const task = JSON.parse(allProps[key]);
          const fila = task.fila;
          
          // Leer correo desde columna U (21) - que es "Correos"
          const correoU = sheet.getRange(fila, 21).getValue();
          
          console.log(`📧 Correos leídos de columna U (fila ${fila}): ${correoU}`);
          
          // Procesar múltiples correos separados por comas
          if (correoU && correoU.trim() !== '') {
            const correosArray = correoU.split(',').map(email => email.trim()).filter(email => email.includes('@'));
            
            if (correosArray.length > 0) {
              console.log(`📨 Enviando a ${correosArray.length} destinatarios en un solo correo:`, correosArray);
              
              // Crear una cadena con todos los correos para el campo "to"
              const todosLosCorreos = correosArray.join(', ');
              
              // Enviar UN SOLO CORREO a todos los destinatarios
              try {
                sendEmailToResponsable(todosLosCorreos, task.data, task.fotosLinks, '');
                console.log(`✅ Correo enviado a todos los destinatarios: ${todosLosCorreos}`);
              } catch (emailError) {
                console.error(`❌ Error enviando correo grupal:`, emailError);
              }
            } else {
              console.log(`⚠️ No se encontraron correos válidos en columna U para fila ${fila}`);
            }
          } else {
            console.log(`⚠️ Columna U vacía para fila ${fila}`);
          }
          
          // Eliminar la tarea
          properties.deleteProperty(key);
          
        } catch (err) {
          console.error(`Error con tarea ${key}:`, err);
        }
      }
    }
    
    // Limpiar triggers
    limpiarTriggers();
    
  } catch (error) {
    console.error('Error en enviarCorreoConRetraso:', error);
  }
}

// Función para limpiar triggers
function limpiarTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'enviarCorreoConRetraso') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Obtiene el email del creador basado en su nombre/cedula
 */
function getEmailByNombre(nombreCedula) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      console.warn('No se encontró la hoja de líderes');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    // Buscar por nombre o cédula en la columna A
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nombreSheet = String(row[0]).trim(); // Columna A
      const cedulaSheet = row[1] ? String(row[1]).trim() : ''; // Columna B
      const emailSheet = row[4] ? String(row[4]).trim() : ''; // Columna E

      // Buscar coincidencia en nombre o cédula
      if (nombreSheet.includes(nombreCedula) || cedulaSheet.includes(nombreCedula) || nombreCedula.includes(nombreSheet)) {
        return emailSheet;
      }
    }

    console.warn('No se encontró email para:', nombreCedula);
    return null;

  } catch (error) {
    console.error('Error al obtener email del creador:', error);
    return null;
  }
}