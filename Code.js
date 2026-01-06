// Configuración de las hojas de cálculo
const SPREADSHEET_ID = '1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII';
const FOLDER_ID_TARJETAS = '1lT2y-uofsPlKQd2Qy5NKJTYh8-rVNBWJ';
const FOLDER_ID_N2 = '1nJtTIkWvWixB_X2EGkcscigqVHVlJ5NO';
const SHEETS = {
  LIDERES: 'Lideres',
  REPORTES_N2: 'Reportes_N2',
  REPORTES_TARJETAS: 'Reportes_Tarjetas',
  REPORTES_CICLOS: 'Reportes_Ciclos',
  REPORTES_MAQUINAS: 'Reportes_Maquinas',
  COMENTARIOS_N2: "Reportes_N2_COMENTARIOS",
  COMENTARIOS_MAQUINAS: "Reportes_Maquinas_COMENTARIOS",
  CICLOS_HISTORIAL: "Ciclos_Historial"
};

const RESPONSABLES_EMAILS = {
  "Jefe del Area Seleccionada": "(POR MAPEAR)",
  "Jefe Aseguramiento de Calidad": "pragestionhumana@pastascomarrico.com",
  "Coordinador de Gestión Ambiental": "pragestionhumana@pastascomarrico.com",
  "Coordinador de Proyectos": "pragestionhumana@pastascomarrico.com",
  "Obras Civiles": "pragestionhumana@pastascomarrico.com",
  "Reparaciones Metalmecanicas IMB": "pragestionhumana@pastascomarrico.com",
  "Técnico Eléctrico": "pragestionhumana@pastascomarrico.com",
  "Técnico Mecánico": "pragestionhumana@pastascomarrico.com",
  "Servicios Técnicos": "pragestionhumana@pastascomarrico.com",

  // Por defecto
  "Por Asignar": "pragestionhumana@pastascomarrico.com"
};

/**
 * Función principal para servir la aplicación web
 */
function doGet() {
  const title = 'PASTAS';
  const faviconUrl = 'https://alimentosdoria.com/wp-content/uploads/2023/01/logo-doria.png';

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(title)
    .setFaviconUrl(faviconUrl)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función para incluir archivos HTML (CSS y JavaScript)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtiene la lista de líderes desde la hoja "Lideres"
 * Formato esperado: Columna A = Nombre, Columna B = Cédula
 */
function getLeaders() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      throw new Error(`La hoja "${SHEETS.LIDERES}" no existe`);
    }

    const data = sheet.getDataRange().getValues();
    const leaders = [];

    // Saltar la primera fila si contiene encabezados
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Solo columna A
        leaders.push({
          info: row[0].toString().trim()
        });
      }
    }

    return leaders;
  } catch (error) {
    console.error('Error al obtener líderes:', error);
    throw new Error('No se pudieron cargar los líderes: ' + error.message);
  }
}

/**
 * Guarda un reporte N2 en la hoja correspondiente
 */
function submitN2Report(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.REPORTES_N2);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.REPORTES_N2);
      const headers = [
        'Fecha de Registro', 'Líder Responsable', 'Proceso', 'ZonaProceso',
        'Anormalidad', 'Proceso Responsable', 'Fecha Prevista Solución',
        'Estado', 'ID Reporte', 'Nombre y Cédula', 'Fotos'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    let fotosLinks = [];

    const getDirectDriveLink = (fileUrl) => {
      const match = fileUrl.match(/\/d\/(.*?)\//);
      if (!match) return fileUrl;
      const fileId = match[1];
      return `https://drive.google.com/uc?export=view&id=${fileId}`;
    };

    // 📸 Guardar fotos en Drive - IDÉNTICO A TARJETAS
    if (formData.fotos && formData.fotos.length > 0) {
      const folder = DriveApp.getFolderById(FOLDER_ID_N2);

      fotosLinks = formData.fotos.map((base64, i) => {
        try {
          const contentType = base64.split(';')[0].split(':')[1];
          const bytes = Utilities.base64Decode(base64.split(',')[1]);
          const blob = Utilities.newBlob(bytes, contentType, `foto_n2_${Date.now()}_${i + 1}.jpg`);
          const file = folder.createFile(blob);

          // ✅ HACER PÚBLICO EXPLÍCITAMENTE
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

          // ✅ CONVERSIÓN IDÉNTICA A TARJETAS
          return getDirectDriveLink(file.getUrl());

        } catch (error) {
          console.error(`Error guardando foto ${i + 1}:`, error);
          return null;
        }
      }).filter(link => link !== null);
    }

    // Generar ID único
    const reportId = 'N2-' + new Date().getTime();

    // Preparar datos
    const fechaRegistro = parseLocalDate(formData.fecha);
    const fechaSolucion = parseLocalDate(formData.fechaSolucion);

    const rowData = [
      fechaRegistro,
      formData.liderResponsable,
      formData.proceso,
      formData.zonaProceso,
      formData.anormalidad,
      formData.procesoResponsable,
      fechaSolucion,
      'Pendiente',
      reportId,
      formData.nombreCedula,
      JSON.stringify(fotosLinks) // ✅ Guardar solo URLs como tarjetas
    ];

    // Insertar en hoja
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    // Formatear
    sheet.getRange(nextRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(nextRow, 7).setNumberFormat('dd/mm/yyyy');

    console.log('📧 Intentando enviar correo de notificación...');
    const leaderInfo = getLeaderInfoFromString(formData.liderResponsable);

    if (leaderInfo && leaderInfo.email) {
      const emailEnviado = sendEmailToLeader(leaderInfo, formData, reportId);
      if (emailEnviado) {
        console.log('✅ Notificación por correo enviada exitosamente');
      } else {
        console.log('⚠️ No se pudo enviar la notificación por correo');
      }
    } else {
      console.warn('⚠️ No se pudo obtener información del líder para enviar correo');
    }

    return {
      success: true,
      reportId,
      message: 'Reporte N2 guardado exitosamente',
      fotos: fotosLinks
    };

  } catch (error) {
    console.error('Error al guardar reporte N2:', error);
    throw new Error('No se pudo guardar el reporte: ' + error.message);
  }
}

/**
 * Obtiene la información del líder desde el string (formato: "Nombre - Cédula")
 */
function getLeaderInfoFromString(leaderString) {
  try {
    console.log('🔍 Buscando información del líder:', leaderString);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      console.error('❌ No se encontró la hoja de líderes');
      return null;
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      console.warn('⚠️ Hoja de líderes vacía o solo tiene encabezados');
      return null;
    }

    // Parsear el string del líder para obtener la cédula
    let cedulaBuscada = '';
    let nombreBuscado = '';

    // Intentar diferentes formatos
    if (leaderString.includes(' - ')) {
      const parts = leaderString.split(' - ');
      if (parts.length >= 2) {
        nombreBuscado = parts[0].trim();
        cedulaBuscada = parts[1].trim();
      }
    } else {
      // Si no viene en formato esperado, usar el string completo como cédula
      cedulaBuscada = leaderString.trim();
    }

    console.log(`📝 Búsqueda - Cédula: "${cedulaBuscada}", Nombre: "${nombreBuscado}"`);

    // Buscar en la hoja de líderes 
    // Asumiendo: columna A = nombre, columna B = cédula, columna E = email
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Verificar que la fila tenga datos
      if (!row[0] && !row[1]) continue;

      const nombreSheet = String(row[0] || '').trim();
      const cedulaSheet = String(row[1] || '').trim();
      const emailSheet = row[4] ? String(row[4]).trim() : '';

      console.log(`📋 Fila ${i}: "${nombreSheet}" | "${cedulaSheet}" | "${emailSheet}"`);

      // Buscar por cédula (más confiable)
      if (cedulaSheet && cedulaSheet === cedulaBuscada) {
        console.log(`✅ Líder encontrado por cédula: ${nombreSheet}`);
        return {
          nombre: nombreSheet,
          cedula: cedulaSheet,
          email: emailSheet
        };
      }

      // Buscar por nombre si la cédula no coincide
      if (nombreBuscado && nombreSheet && nombreSheet.includes(nombreBuscado)) {
        console.log(`✅ Líder encontrado por nombre: ${nombreSheet}`);
        return {
          nombre: nombreSheet,
          cedula: cedulaSheet,
          email: emailSheet
        };
      }
    }

    console.warn('❌ Líder no encontrado en la hoja para:', leaderString);

    // Log de las primeras filas para debugging
    console.log('📊 Primeras filas de líderes:');
    for (let i = 1; i < Math.min(5, data.length); i++) {
      const row = data[i];
      console.log(`Fila ${i}: ${String(row[0])} | ${String(row[1])} | ${String(row[4])}`);
    }

    return null;

  } catch (error) {
    console.error('💥 Error crítico al obtener información del líder:', error);
    return null;
  }
}

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
            <p style="margin: 0;"><strong>📋 Acción requerida:</strong> Por favor revisar este reporte y tomar las acciones correspondientes en el sistema.</p>
          </div>
          
          <p>Puede acceder al sistema para ver más detalles y actualizar el estado del reporte.</p>
          
          <div style="text-align: center; margin: 20px 0;">
            <a href="${ScriptApp.getService().getUrl()}" 
               style="background-color: #d9534f; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; display: inline-block;">
              📊 Acceder al Sistema
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

// Conversor de fechas sin desfase UTC
function parseLocalDate(dateString) {
  if (!dateString) return new Date();

  const [datePart, timePart] = dateString.trim().split(' ');
  const [year, month, day] = datePart.split('-').map(Number);

  let hour = 0, minute = 0;
  if (timePart) {
    [hour, minute] = timePart.split(':').map(Number);
  }

  return new Date(year, month - 1, day, hour, minute);
}

/**
 * Obtiene todos los reportes N2 desde la hoja de cálculo
 */
function getN2Reports() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const reports = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      // ✅ CONVERTIR FOTOS DE DRIVE A BASE64 (IGUAL QUE TARJETAS)
      let fotosBase64 = [];
      if (row[10]) {
        try {
          const urls = JSON.parse(row[10]);
          fotosBase64 = urls.map(url => {
            const idMatch = url.match(/id=([a-zA-Z0-9_-]+)/);
            if (!idMatch) return '';

            try {
              const file = DriveApp.getFileById(idMatch[1]);
              const blob = file.getBlob();
              return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
            } catch (e) {
              console.log('Error convirtiendo URL a Base64:', e);
              return ''; // Devolver string vacío si falla
            }
          }).filter(base64 => base64 !== ''); // Filtrar strings vacíos

        } catch (e) {
          console.log('Error parseando fotos para fila ' + i + ': ' + e);
          fotosBase64 = [];
        }
      }

      reports.push({
        fechaRegistro: row[0]
          ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : "",
        liderResponsable: row[1] || "",
        proceso: row[2] || "",
        zonaProceso: row[3] || "",
        anormalidad: row[4] || "",
        procesoResponsable: row[5] || "",
        fechaSolucion: row[6]
          ? Utilities.formatDate(new Date(row[6]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : "",
        estado: row[7] || "Pendiente",
        id: row[8] || "",
        nombreCedula: row[9] || "",
        fotos: fotosBase64 // ✅ Esto ahora será array de Base64
      });
    }

    // Ordenar por fecha descendente
    reports.sort((a, b) => new Date(b.fechaRegistro) - new Date(a.fechaRegistro));

    console.log("✅ Reportes N2 obtenidos: " + reports.length);
    if (reports.length > 0 && reports[0].fotos.length > 0) {
      console.log("✅ Primera foto convertida a Base64:", reports[0].fotos[0].substring(0, 50) + "...");
    }

    return reports;

  } catch (error) {
    Logger.log("❌ Error al obtener reportes N2: " + error);
    return [];
  }
}

/**
 * Actualiza el estado de un reporte N2
 */
function updateReportStatus(reportId, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);

    if (!sheet) {
      throw new Error('La hoja de reportes N2 no existe');
    }

    const data = sheet.getDataRange().getValues();

    // Buscar el reporte por ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][8] === reportId) {
        sheet.getRange(i + 1, 8).setValue(newStatus); // Columna G contiene el estado
        return {
          success: true,
          message: 'Estado actualizado exitosamente'
        };
      }
    }

    throw new Error('Reporte no encontrado');

  } catch (error) {
    console.error('Error al actualizar estado:', error);
    throw new Error('No se pudo actualizar el estado: ' + error.message);
  }
}

//LOGIN

function validarCedula(cedula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LIDERES");
  const data = sheet.getDataRange().getValues();

  // Recorre desde la fila 2 (fila 1 son encabezados)
  for (let i = 1; i < data.length; i++) {
    const cedulaSheet = data[i][1];
    const rolSheet = data[i][2];
    const procesoSheet = data[i][3];
    const correoSheet = data[i][4];
    const empresaSheet = data[i][5];

    if (String(cedulaSheet) === String(cedula)) {
      return { success: true, rol: rolSheet, proceso_user: procesoSheet, correo: correoSheet, empresa: empresaSheet };
    }
  }

  return { success: false };
}

function getNombreByCedula(cedula) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      throw new Error(`La hoja "${SHEETS.LIDERES}" no existe`);
    }

    const data = sheet.getDataRange().getValues();

    // Buscar la cédula en la columna B (índice 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const cedulaSheet = String(row[1]).trim();

      if (cedulaSheet === String(cedula).trim()) {
        // Retornar el nombre de la columna A (índice 0)
        return {
          success: true,
          nombre: row[0] ? row[0].toString().trim() : 'Usuario'
        };
      }
    }

    return { success: false, nombre: '' };
  } catch (error) {
    console.error('Error al obtener nombre:', error);
    return { success: false, nombre: '' };
  }
}

function submitTarjetaReport(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);

    if (!sheet) throw new Error(`La hoja "${SHEETS.REPORTES_TARJETAS}" no existe`);

    let fotosLinks = [];

    // Función para convertir link de Drive a link directo
    const getDirectDriveLink = (fileUrl) => {
      const match = fileUrl.match(/\/d\/(.*?)\//);
      if (!match) return fileUrl;
      const fileId = match[1];
      return `https://drive.google.com/uc?export=view&id=${fileId}`;
    };

    // 📸 Guardar fotos en Drive si existen
    if (data.fotos && data.fotos.length > 0) {
      const folder = DriveApp.getFolderById(FOLDER_ID_TARJETAS);
      fotosLinks = data.fotos.map((base64, i) => {
        const contentType = base64.split(';')[0].split(':')[1];
        const bytes = Utilities.base64Decode(base64.split(',')[1]);
        const blob = Utilities.newBlob(bytes, contentType, `foto_${Date.now()}_${i + 1}.jpg`);
        const file = folder.createFile(blob);
        // Convertimos a link directo
        return getDirectDriveLink(file.getUrl());
      });
    }

    const totalTarjetas = sheet.getLastRow() - 1;
    const tarjetaId = `TAR-${String(totalTarjetas + 1).padStart(4, '0')}`;


    // CORRECCIÓN: Array con todas las columnas en el orden correcto
    const newRow = [
      data.zonaRiesgo || '',
      data.nombreCedula || '',
      data.ubicacion || '',
      data.prioridad || '',
      data.descripcionProblema || '',
      data.tipoRiesgo || '',
      data.problemaAsociado || '',
      data.sistemaGestion || '',
      data.responsableSolucion || '',
      data.generadaPor || '',
      data.fechaCreacionTarjeta || '',
      data.estado || 'Abierta',
      JSON.stringify(fotosLinks),
      '',
      '',
      data.requiereSAP || 'No',
      tarjetaId
    ];

    sheet.appendRow(newRow);

    const creadorEmail = getEmailByNombre(data.nombreCedula);
    const responsableEmail = RESPONSABLES_EMAILS[data.responsableSolucion] || RESPONSABLES_EMAILS["Por Asignar"];

    // Enviar correos
    if (creadorEmail) {
      sendEmailToCreador(creadorEmail, data, fotosLinks);
    }

    if (responsableEmail) {
      sendEmailToResponsable(responsableEmail, data, fotosLinks, creadorEmail);
    }

    return {
      success: true,
      tarjetaId: tarjetaId,
      message: 'Tarjeta de anormalidad registrada exitosamente',
      fotos: fotosLinks
    };
  } catch (error) {
    console.error('Error al guardar tarjeta:', error);
    return {
      success: false,
      message: 'Error al guardar la tarjeta: ' + error.message
    };
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

/**
 * Envía correo de confirmación al creador de la tarjeta
 */
function sendEmailToCreador(creadorEmail, data, fotosLinks) {
  try {
    const subject = `✅ Tarjeta de Anormalidad Creada - ${data.prioridad}`;

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
            <p style="margin: 0;"><strong>📋 Estado:</strong> La tarjeta ha sido asignada a <strong>${data.responsableSolucion}</strong> para su revisión y solución.</p>
          </div>
          
          <p>Puede dar seguimiento a esta tarjeta accediendo al sistema.</p>
          
          <div style="text-align: center; margin: 20px 0;">
            <a href="${ScriptApp.getService().getUrl()}" 
               style="background-color: #28a745; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; display: inline-block;">
              📊 Ver en el Sistema
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
            <p style="margin: 0;"><strong>📋 Acción requerida:</strong> Por favor revisar esta anormalidad reportada y tomar las acciones correspondientes.</p>
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
              📊 Acceder al Sistema
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

function getTarjetasReports() {
  try {
    console.log('🔍 Iniciando getTarjetasReports...');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);
    if (!sheet) {
      console.log('❌ Hoja no encontrada:', SHEETS.REPORTES_TARJETAS);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    console.log('📊 Filas totales:', data.length);
    
    const tarjetas = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Verificar que row exista y tenga al menos 1 elemento
      if (!row || row.length === 0 || !row[0]) continue;

      // Convertir fotos de Drive a Base64
      let fotosUrls = [];
      if (row[12]) { // Columna 13: fotos (JSON con URLs)
        try {
          // Parsear JSON de URLs
          fotosUrls = JSON.parse(row[12]);
          
          // Asegurar que sea array
          if (!Array.isArray(fotosUrls)) {
            fotosUrls = [];
          }
          
          console.log(`Fila ${i+1}: ${fotosUrls.length} URLs de fotos`);
          
        } catch (e) {
          console.error('Error parseando JSON de fotos (fila ' + i + '):', e);
          fotosUrls = [];
        }
      }

      // MANEJO SEGURO DEL ID - CRÍTICO
      let tarjetaId;
      try {
        tarjetaId = row[16];
        if (!tarjetaId) {
          // Si no hay ID en row[16], crear uno basado en fecha
          if (row[10]) {
            const fecha = new Date(row[10]);
            if (!isNaN(fecha.getTime())) {
              tarjetaId = 'TAR-' + fecha.getTime();
            } else {
              tarjetaId = 'TAR-' + new Date().getTime();
            }
          } else {
            tarjetaId = 'TAR-' + new Date().getTime();
          }
        }
      } catch (e) {
        console.error('Error generando ID para fila', i, ':', e);
        tarjetaId = 'TAR-' + new Date().getTime(); // Fallback
      }

      // Fecha de creación segura
      let fechaCreacion = "";
      try {
        if (row[10]) {
          const fecha = new Date(row[10]);
          if (!isNaN(fecha.getTime())) {
            fechaCreacion = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
        }
      } catch (e) {
        console.error('Error procesando fecha creación fila', i, ':', e);
      }

      // Fecha de cierre segura (si existe)
      let fechaCierreTarjeta = "";
      if (row.length > 24 && row[24]) {
        try {
          const fechaCierre = new Date(row[24]);
          if (!isNaN(fechaCierre.getTime())) {
            fechaCierreTarjeta = Utilities.formatDate(fechaCierre, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
        } catch (e) {
          console.error('Error procesando fecha cierre fila', i, ':', e);
        }
      }

      tarjetas.push({
        rowIndex: i + 1,
        id: tarjetaId,
        zonaRiesgo: row[0] || "",
        nombreCedula: row[1] || "",
        ubicacion: row[2] || "",
        prioridad: row[3] || "",
        descripcionProblema: row[4] || "",
        tipoRiesgo: row[5] || "",
        problemaAsociado: row[6] || "",
        sistemaGestion: row[7] || "",
        responsableSolucion: row[8] || "",
        responsableNombreVisualizarReporte: row[18] || "",
        generadaPor: row[9] || "",
        fechaCreacion: fechaCreacion,
        estado: row[11] || "Abierta",
        fotos: fotosUrls,
        comentarioCierre: row[13] || "",
        responsableCierre: row[14] || "",
        requiereSAP: row[15] || "No",
        fechaCierreTarjeta: fechaCierreTarjeta
      });
    }

    console.log('✅ Tarjetas procesadas:', tarjetas.length);
    
    // Ordenar por prioridad solo si hay tarjetas
    if (tarjetas.length > 0) {
      const prioridadOrden = { "Alta": 1, "Media": 2, "Baja": 3 };
      tarjetas.sort((a, b) => (prioridadOrden[a.prioridad] || 999) - (prioridadOrden[b.prioridad] || 999));
    }

    return tarjetas;

  } catch (error) {
    console.error("❌ Error CRÍTICO en getTarjetasReports:", error);
    console.error("❌ Stack trace:", error.stack);
    return [];
  }
}

/**
 * Cierra una tarjeta de anormalidad con un comentario
 */
function closeTarjetaReport(rowIndex, comentario, responsableCierre, fechaCierreTarjeta = null) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);
    
    if (!sheet) {
      throw new Error('La hoja de tarjetas no existe');
    }

    // Usar la fecha pasada como parámetro o la fecha actual
    const fechaCierre = fechaCierreTarjeta 
      ? Utilities.formatDate(new Date(fechaCierreTarjeta), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
      : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    sheet.getRange(rowIndex, 12).setValue('Cerrada');
    sheet.getRange(rowIndex, 14).setValue(comentario);
    sheet.getRange(rowIndex, 15).setValue(responsableCierre);
    sheet.getRange(rowIndex, 25).setValue(fechaCierre);

    return {
      success: true,
      message: 'Tarjeta cerrada exitosamente'
    };

  } catch (error) {
    console.error('Error al cerrar tarjeta:', error);
    throw new Error('No se pudo cerrar la tarjeta: ' + error.message);
  }
}

/**
 * Agrega un comentario a un reporte N2
 */
function addCommentToReport(reportId, comment, autor) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2 + "_COMENTARIOS") || ss.insertSheet(SHEETS.REPORTES_N2 + "_COMENTARIOS");

  // Si la hoja es nueva, crea encabezados
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID Reporte", "Autor", "Comentario", "Fecha"]);
  }

  sheet.appendRow([
    reportId,
    autor,
    comment,
    new Date()
  ]);

  return { success: true, message: "Comentario agregado exitosamente" };
}

/**
 * Obtiene los comentarios de un reporte N2 específico
 */
function getCommentsForReport(reportId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2 + "_COMENTARIOS");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === reportId) {
      comments.push({
        autor: row[1],
        comentario: row[2],
        fecha: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
  }
  return comments;
}

/**
 * Cambia el responsable de un reporte
 */
function updateReportResponsible(reportId, newResponsible) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);
  if (!sheet) throw new Error("No existe la hoja N2");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === reportId) {
      sheet.getRange(i + 1, 2).setValue(newResponsible);
      return { success: true, message: "Responsable actualizado correctamente" };
    }
  }

  throw new Error("Reporte no encontrado");
}

function getCommentsCountForReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2 + "_COMENTARIOS");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const counts = {};

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (!counts[id]) counts[id] = 0;
    counts[id]++;
  }

  return counts;
}

function updateReportDate(reportId, newDate) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);            // usa mismo origen que updateReportResponsible
  const sheet = ss.getSheetByName(SHEETS.REPORTES_N2);          // usa la misma constante/hoja
  if (!sheet) throw new Error("No existe la hoja N2");

  const data = sheet.getDataRange().getValues();


  const idCol = 8;     // columna donde está el ID (ej. H -> 7, H index 8 si antes lo usaste así)
  const fechaCol = 6;  // columna donde quieres escribir la fecha (ajusta si es necesario)

  Logger.log("📩 ID recibido desde frontend: %s", reportId);
  Logger.log("🗂 Ejemplo IDs (filas 2-6): %s", data.slice(1, 6).map(r => r[idCol]).join(", "));

  const normalizedTarget = String(reportId).trim();

  for (let i = 1; i < data.length; i++) {
    const cellId = String(data[i][idCol]).trim();
    Logger.log("Comparando fila %d: hojaId=%s target=%s", i + 1, cellId, normalizedTarget);

    if (cellId === normalizedTarget) {
      // intentar guardar como DATE (si newDate viene 'YYYY-MM-DD' lo convertimos)
      let valueToWrite = newDate;
      try {
        // Si newDate es string "YYYY-MM-DD", esto lo convierte a Date
        const parsed = new Date(newDate);
        if (!isNaN(parsed.getTime())) {
          parsed.setMinutes(parsed.getMinutes() + parsed.getTimezoneOffset());
          valueToWrite = parsed;
        }
      } catch (e) {
        // si falla, dejamos el string (setValue aceptará string también)
      }

      sheet.getRange(i + 1, fechaCol + 1).setValue(valueToWrite);
      Logger.log("✅ Fecha actualizada para ID %s en fila %d", reportId, i + 1);
      return { success: true, message: "Fecha actualizada correctamente" };
    }
  }

  // si no encontró, devolver info para debugging (no throw si prefieres manejarlo en frontend)
  Logger.log("❌ No se encontró el reporte con ID %s", reportId);
  throw new Error("No se encontró el reporte con ID " + reportId);
}

function submitMaquinasReport(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.REPORTES_MAQUINAS);
      const headers = [
        'Fecha de Registro',
        'Mecanico Responsable',
        'Proceso',
        'AreaProceso',
        'Subsistema',
        'Anormalidad',
        'AreaResponsable',
        'Estado',
        'ID Reporte',
        'Criticidad' // <CHANGE> Agregada columna Criticidad
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    const reportId = 'MAQ-' + new Date().getTime();
    const fechaRegistro = parseLocalDate(formData.fecha);

    const rowData = [
      fechaRegistro,
      formData.mecanicoResponsable,
      formData.proceso,
      formData.areaProceso,
      formData.subsistema,
      formData.anormalidad,
      formData.areaResponsable,
      'Abierto',
      reportId,
      formData.criticidad || 'Media' // <CHANGE> Agregado campo criticidad
    ];

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(nextRow, 1).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('Reporte de máquina guardado exitosamente con ID:', reportId);

    return {
      success: true,
      reportId: reportId,
      message: 'Reporte de máquina guardado exitosamente'
    };

  } catch (error) {
    console.error('Error al guardar reporte de máquina:', error);
    return {
      success: false,
      message: 'No se pudo guardar el reporte: ' + error.message
    };
  }
}

/**
 * Obtiene todos los reportes de máquinas desde la hoja de cálculo
 */
function getMaquinasReports() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const reports = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const reportId = row[8] || 'MAQ-' + new Date(row[0]).getTime();

      reports.push({
        id: reportId,
        fechaRegistro: row[0]
          ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : "",
        mecanicoResponsable: row[1] || "",
        proceso: row[2] || "",
        areaProceso: row[3] || "",
        subsistema: row[4] || "",
        anormalidad: row[5] || "",
        areaResponsable: row[6] || "",
        estado: row[7] || "Abierto",
        criticidad: row[9] || "Media" // <CHANGE> Agregado campo criticidad
      });
    }

    reports.sort((a, b) => new Date(b.fechaRegistro) - new Date(a.fechaRegistro));

    console.log("Reportes de máquinas obtenidos: " + reports.length);

    return reports;

  } catch (error) {
    Logger.log("Error al obtener reportes de máquinas: " + error);
    return [];
  }
}

/**
 * Actualiza el estado de un reporte de máquina
 */
function updateMaquinaStatus(reportId, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);

    if (!sheet) {
      throw new Error('La hoja de reportes de máquinas no existe');
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const currentId = data[i][8] || 'MAQ-' + new Date(data[i][0]).getTime();

      if (currentId === reportId) {
        sheet.getRange(i + 1, 8).setValue(newStatus);

        if (!data[i][8]) {
          sheet.getRange(i + 1, 9).setValue(reportId);
        }

        return {
          success: true,
          message: 'Estado actualizado exitosamente'
        };
      }
    }

    throw new Error('Reporte no encontrado');

  } catch (error) {
    console.error('Error al actualizar estado de máquina:', error);
    throw new Error('No se pudo actualizar el estado: ' + error.message);
  }
}

/**
 * Agrega un comentario a un reporte de máquina
 */
function addMaquinasCommentToReport(reportId, comment, autor) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "Reportes_Maquinas_COMENTARIOS";
  let sheet = ss.getSheetByName(sheetName);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID Reporte", "Autor", "Comentario", "Fecha"]);
  }

  sheet.appendRow([
    reportId,
    autor,
    comment,
    new Date()
  ]);

  return { success: true, message: "Comentario agregado exitosamente" };
}

/**
 * Obtiene los comentarios de un reporte de máquina específico
 */
function getMaquinasCommentsForReport(reportId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Maquinas_COMENTARIOS");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === reportId) {
      comments.push({
        autor: row[1],
        comentario: row[2],
        fecha: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
  }
  return comments;
}

/**
 * Obtiene el conteo de comentarios para todos los reportes de máquinas
 */
function getMaquinasCommentsCountForReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Maquinas_COMENTARIOS");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const counts = {};

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (!counts[id]) counts[id] = 0;
    counts[id]++;
  }

  return counts;
}

// <CHANGE> Nueva función para actualizar criticidad
function updateMaquinaCriticidad(reportId, newCriticidad) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_MAQUINAS);

    if (!sheet) {
      throw new Error('La hoja de reportes de máquinas no existe');
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const currentId = data[i][8] || 'MAQ-' + new Date(data[i][0]).getTime();

      if (currentId === reportId) {
        // Actualizar criticidad en columna J (índice 9)
        sheet.getRange(i + 1, 10).setValue(newCriticidad);

        if (!data[i][8]) {
          sheet.getRange(i + 1, 9).setValue(reportId);
        }

        return {
          success: true,
          message: 'Criticidad actualizada exitosamente'
        };
      }
    }

    throw new Error('Reporte no encontrado');

  } catch (error) {
    console.error('Error al actualizar criticidad de máquina:', error);
    throw new Error('No se pudo actualizar la criticidad: ' + error.message);
  }
}

// <CHANGE> Obtiene todos los reportes consolidados
function getConsolidadoReports() {
  try {
    const n2 = getN2Reports();
    const tarjetas = getTarjetasReports();
    const maquinas = getMaquinasReports();
    const ciclos = getCiclosMejora();

    return {
      n2: n2,
      tarjetas: tarjetas,
      maquinas: maquinas,
      ciclos: ciclos
    };

  } catch (error) {
    console.error('Error al obtener consolidado:', error);
    throw new Error('No se pudo cargar el consolidado: ' + error.message);
  }
}

function addTarjetaCommentToReport(tarjetaId, comment, autor) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "Reportes_Tarjetas_COMENTARIOS";
  let sheet = ss.getSheetByName(sheetName);

  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ID Tarjeta", "Autor", "Comentario", "Fecha"]);
  }

  sheet.appendRow([
    tarjetaId,
    autor,
    comment,
    new Date()
  ]);

  return { success: true, message: "Comentario agregado exitosamente" };
}

function getTarjetasCommentsForReport(tarjetaId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Tarjetas_COMENTARIOS");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === tarjetaId) {
      comments.push({
        autor: row[1],
        comentario: row[2],
        fecha: Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
  }
  return comments;
}

function getTarjetasCommentsCountForReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Reportes_Tarjetas_COMENTARIOS");
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const counts = {};

  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (!counts[id]) counts[id] = 0;
    counts[id]++;
  }

  return counts;
}

/**
 * Registra un nuevo usuario en la hoja de líderes
 * Columnas: B=Cédula, C=Rol, D=Proceso, E=Correo, F=Empresa, G=Nombres, H=Apellidos
 */
function registrarUsuario(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.LIDERES);

    if (!sheet) {
      throw new Error('La hoja LIDERES no existe');
    }

    // Obtener datos de la columna B (cédulas) para verificar duplicados y encontrar última fila
    const lastRowSheet = sheet.getLastRow();
    const colB = sheet.getRange(1, 2, lastRowSheet, 1).getValues(); // Columna B completa

    // <CHANGE> Buscar la última fila con datos reales en columna B
    let nextRow = 1;
    for (let i = 0; i < colB.length; i++) {
      const cedulaExistente = String(colB[i][0] || '').trim();

      // Verificar si la cédula ya existe
      if (cedulaExistente === formData.cedula) {
        return {
          success: false,
          message: 'Esta cédula ya está registrada en el sistema'
        };
      }

      // Si la celda tiene datos, actualizar nextRow
      if (cedulaExistente !== '') {
        nextRow = i + 2; // +2 porque i es base 0 y queremos la siguiente fila
      }
    }

    // Convertir todos los campos a mayúsculas excepto el correo
    const rowData = [
      '',                                    // Columna A - Vacía
      formData.cedula.toUpperCase(),         // Columna B - Cédula
      'USUARIO',                             // Columna C - Rol
      formData.proceso.toUpperCase(),        // Columna D - Proceso
      formData.correo,                       // Columna E - Correo (sin cambios)
      formData.empresa.toUpperCase(),        // Columna F - Empresa
      formData.nombres.toUpperCase(),        // Columna G - Nombres
      formData.apellidos.toUpperCase()       // Columna H - Apellidos
    ];

    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    console.log('Usuario registrado en fila: ' + nextRow + ' - Cédula: ' + formData.cedula);

    return {
      success: true,
      message: 'Usuario registrado exitosamente'
    };

  } catch (error) {
    console.error('Error al registrar usuario:', error);
    return {
      success: false,
      message: 'Error al registrar: ' + error.message
    };
  }
}

// ========== CICLO DE MEJORA ==========

function getNextCicloId() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('REPORTES_CICLOS');

    if (!sheet) {
      return 'CM-001';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return 'CM-001';
    }

    const totalCiclos = lastRow - 1;
    const nextNumber = totalCiclos + 1;
    return 'CM-' + String(nextNumber).padStart(3, '0');

  } catch (error) {
    console.error('Error obteniendo siguiente ID de ciclo:', error);
    return 'CM-001';
  }
}

function submitCicloMejora(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Reportes_Ciclos');

    if (!sheet) {
      sheet = ss.insertSheet('Reportes_Ciclos');
      const headers = [
        'ID Ciclo', 'Fecha Registro', 'Nombre Ciclo', 'Aviso Mantenimiento',
        'Proceso', 'Equipo/Máquina', 'Líder', 'Integrantes',
        'Tipo Foco Mejora', 'Datos Foco Mejora',
        'Defecto Principal',
        'Causas Medio Ambiente', 'Causas Mano de Obra', 'Causas Materiales',
        'Causas Tiempo', 'Causas Método', 'Causas Máquina',
        'Análisis 5 Por Qué', 'Plan de Acción 5W+2H', 'Estado', 'Creado Por'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, headers.length).setBackground('#0f307f');
      sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
    }

    const cicloId = formData.cicloId || getNextCicloId();
    const fechaRegistro = parseLocalDate(formData.fecha);

    const espina = formData.espinaPescado || {};
    const causasAmbiente = (espina.medioAmbiente || []).join(' | ');
    const causasMano = (espina.manoDeObra || []).join(' | ');
    const causasMateriales = (espina.materiales || []).join(' | ');
    const causasTiempo = (espina.tiempo || []).join(' | ');
    const causasMetodo = (espina.metodo || []).join(' | ');
    const causasMaquina = (espina.maquina || []).join(' | ');

    const analisis5PorquesStr = formData.analisis5Porques ?
      JSON.stringify(formData.analisis5Porques) : '';

    // 🔥 CORRECCIÓN: GUARDAR PLAN DE ACCIÓN LIMPIO
    const focoMejora = formData.focoMejora || {};
    const tipoFoco = focoMejora.tipo || '';
    const datosFocoStr = JSON.stringify(focoMejora);
    
    let planAccionStr = '';
    if (formData.planAccion && Array.isArray(formData.planAccion)) {
      // Limpiar el array antes de convertirlo a string
      const planAccionLimpio = formData.planAccion.map(accion => {
        return {
          cual: (accion.cual || '').replace(/["'\\$%]/g, ''),
          que: (accion.que || '').replace(/["'\\$%]/g, ''),
          donde: (accion.donde || '').replace(/["'\\$%]/g, ''),
          quien: (accion.quien || '').replace(/["'\\$%]/g, ''),
          como: (accion.como || '').replace(/["'\\$%]/g, ''),
          cuando: (accion.cuando || '').replace(/["'\\$%]/g, ''),
          cuanto: (accion.cuanto || '').replace(/["'\\$%]/g, '')
        };
      });
      
      planAccionStr = JSON.stringify(planAccionLimpio);
      console.log('[Backend] Plan acción guardado (limpio):', planAccionStr.substring(0, 200));
    }

    const rowData = [
      cicloId, 
      fechaRegistro, 
      formData.nombreCiclo || '',
      formData.avisoMantenimiento || '', 
      formData.proceso || '',
      formData.equipoMaquina || '', 
      formData.lider || '',
      formData.integrantes || '',
      tipoFoco, 
      datosFocoStr,
      (formData.defecto || '').replace(/["'\\$%]/g, ''), // Limpiar defecto también
      causasAmbiente, 
      causasMano, 
      causasMateriales,
      causasTiempo, 
      causasMetodo, 
      causasMaquina,
      analisis5PorquesStr, 
      planAccionStr,
      'Abierto', 
      formData.creadoPor || ''
    ];

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(nextRow, 2).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('Ciclo de Mejora guardado exitosamente con ID:', cicloId);

    return {
      success: true,
      cicloId: cicloId,
      message: 'Ciclo de Mejora registrado exitosamente'
    };

  } catch (error) {
    console.error('Error al guardar Ciclo de Mejora:', error);
    return {
      success: false,
      message: 'Error al guardar: ' + error.message
    };
  }
}

/**
 * Actualiza el responsable de una tarjeta
 */
function updateTarjetaResponsible(rowIndex, newResponsible) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.REPORTES_TARJETAS);

    if (!sheet) {
      throw new Error('La hoja de tarjetas no existe');
    }

    // Actualizar en la columna 9 (índice 9 = columna I = Responsable)
    sheet.getRange(rowIndex, 9).setValue(newResponsible);

    return {
      success: true,
      message: 'Responsable actualizado correctamente'
    };

  } catch (error) {
    console.error('Error al actualizar responsable de tarjeta:', error);
    throw new Error('No se pudo actualizar el responsable: ' + error.message);
  }
}

// ========== FUNCIONES GESTIÓN DE CICLOS ==========

function getCiclosMejora() {
  try {
    console.log('[Backend] Iniciando getCiclosMejora...');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Reportes_Ciclos');

    if (!sheet) {
      sheet = ss.getSheetByName('CICLOS_MEJORA');
    }

    if (!sheet) {
      console.log('[Backend] ERROR: Ninguna hoja de ciclos encontrada');
      return [];
    }

    console.log('[Backend] Hoja encontrada:', sheet.getName());

    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('[Backend] Solo encabezados, sin datos');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    console.log('[Backend] Datos obtenidos, filas:', data.length);

    const ciclos = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Saltar filas completamente vacías
      if (!row[0] && !row[1] && !row[2]) continue;

      // CONVERTIR FECHA A STRING ISO
      let fechaStr = '';
      try {
        if (row[1] instanceof Date) {
          fechaStr = row[1].toISOString();
        } else if (row[1]) {
          fechaStr = new Date(row[1]).toISOString();
        }
      } catch (e) {
        fechaStr = '';
      }

      // 🔥 DIAGNÓSTICO: Ver el plan de acción crudo
      let planAccionStr = String(row[18] || '').trim();
      console.log(`[Backend] Fila ${i+1} - Plan acción CRUDO (primeros 500 chars):`, 
                  planAccionStr.substring(0, 500));
      console.log(`[Backend] Fila ${i+1} - Plan acción longitud:`, planAccionStr.length);
      
      let planAccionParseado = [];
      
      if (planAccionStr && planAccionStr !== '') {
  try {
    // 1. Limpiar caracteres básicos
    planAccionStr = planAccionStr
      .replace(/^[\x00-\x1F]+/, '')
      .trim();
    
    // 2. DIAGNÓSTICO: Ver el JSON exacto
    console.log(`[Backend] JSON antes de parsear (primeros 500 chars):`, 
                planAccionStr.substring(0, 500));
    
    // 3. Intentar parsear directamente primero
    try {
      planAccionParseado = JSON.parse(planAccionStr);
      console.log(`[Backend] Parseado directo exitoso, ${planAccionParseado.length} acciones`);
    } catch (parseError1) {
      console.log(`[Backend] Primer intento falló: ${parseError1.message}`);
      
      // 4. Intentar con función de reparación
      try {
        planAccionParseado = repararJSONPlanAccion(planAccionStr);
        console.log(`[Backend] JSON reparado exitoso, ${planAccionParseado.length} acciones`);
      } catch (parseError2) {
        console.error(`[Backend] Segundo intento falló: ${parseError2.message}`);
        
        // 5. Último intento: extraer solo el array manualmente
        try {
          // Buscar el array completo entre [ y ]
          const inicio = planAccionStr.indexOf('[');
          const fin = planAccionStr.lastIndexOf(']');
          
          if (inicio !== -1 && fin !== -1 && fin > inicio) {
            const jsonExtraido = planAccionStr.substring(inicio, fin + 1);
            // Limpiar escapes de barras invertidas
            const jsonLimpio = jsonExtraido.replace(/\\\\/g, '\\');
            planAccionParseado = JSON.parse(jsonLimpio);
            console.log(`[Backend] Extracción manual exitosa, ${planAccionParseado.length} acciones`);
          } else {
            throw new Error('No se encontró array en el string');
          }
        } catch (parseError3) {
          console.error(`[Backend] Tercer intento falló: ${parseError3.message}`);
          planAccionParseado = [];
        }
      }
    }
  } catch (error) {
    console.error(`❌ Error procesando plan de acción en fila ${i+1}:`, error);
    planAccionParseado = [];
  }

      }

      const ciclo = {
        id: String(row[0] || '').trim(),
        fecha: fechaStr,
        nombre: String(row[2] || '').trim(),
        aviso: String(row[3] || '').trim(),
        proceso: String(row[4] || '').trim(),
        equipo: String(row[5] || '').trim(),
        lider: String(row[6] || '').trim(),
        integrantes: String(row[7] || '').trim(),
        tipoFoco: String(row[8] || '').trim(),
        datosFoco: String(row[9] || '').trim(),
        defecto: String(row[10] || '').trim(),
        causasAmbiente: String(row[11] || '').trim(),
        causasMano: String(row[12] || '').trim(),
        causasMateriales: String(row[13] || '').trim(),
        causasTiempo: String(row[14] || '').trim(),
        causasMetodo: String(row[15] || '').trim(),
        causasMaquina: String(row[16] || '').trim(),
        analisis5Porques: String(row[17] || '').trim(),
        planAccion: planAccionParseado,
        estado: String(row[19] || 'Abierto').trim(),
        creadoPor: String(row[20] || '').trim()
      };

      ciclos.push(ciclo);
    }

    console.log('[Backend] Ciclos procesados:', ciclos.length);
    
    // Mostrar detalles del primer ciclo para diagnóstico
    if (ciclos.length > 0) {
      console.log('[Backend] Primer ciclo ID:', ciclos[0].id);
      console.log('[Backend] Primer ciclo - planAccion tipo:', typeof ciclos[0].planAccion);
      console.log('[Backend] Primer ciclo - planAccion es array?', Array.isArray(ciclos[0].planAccion));
      if (Array.isArray(ciclos[0].planAccion)) {
        console.log('[Backend] Primer ciclo - acciones:', ciclos[0].planAccion.length);
        if (ciclos[0].planAccion.length > 0) {
          console.log('[Backend] Primera acción:', JSON.stringify(ciclos[0].planAccion[0]).substring(0, 200));
        }
      }
    }

    // Ordenar por fecha descendente
    if (ciclos.length > 1) {
      ciclos.sort((a, b) => {
        try {
          const dateA = a.fecha ? new Date(a.fecha).getTime() : 0;
          const dateB = b.fecha ? new Date(b.fecha).getTime() : 0;
          return dateB - dateA;
        } catch (e) {
          return 0;
        }
      });
    }

    return ciclos;

  } catch (error) {
    console.error('[Backend] ERROR en getCiclosMejora:', error);
    console.error('[Backend] Stack trace:', error.stack);
    return [];
  }
}

// Obtener historial de seguimiento de un ciclo 
function getHistorialCiclo(cicloId) {
  try {
    console.log('🔍 Buscando historial para ciclo:', cicloId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('CICLOS_HISTORIAL');

    // Si no existe la hoja, crear una nueva
    if (!sheet) {
      console.log('📝 Creando nueva hoja de historial');
      sheet = ss.insertSheet('CICLOS_HISTORIAL');
      sheet.appendRow(['ID Ciclo', 'Fecha', 'Estado', 'Comentario', 'Autor']);
      sheet.getRange(1, 1, 1, 5).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
      return []; // Retornar vacío porque es nueva
    }

    const data = sheet.getDataRange().getValues();
    console.log('📊 Datos en hoja de historial:', data.length, 'filas');

    if (data.length <= 1) {
      console.log('📭 Hoja de historial vacía o solo encabezados');
      return [];
    }

    const historial = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();

      console.log(`Fila ${i}: ID="${rowCicloId}" buscando="${cicloId}"`);

      if (rowCicloId === cicloId) {
        const registro = {
          cicloId: rowCicloId,
          fecha: row[1] ? row[1].toISOString() : new Date().toISOString(),
          estado: String(row[2] || ''),
          comentario: String(row[3] || ''),
          autor: String(row[4] || 'Sistema')
        };

        console.log('✅ Registro encontrado:', registro);
        historial.push(registro);
      }
    }

    console.log('📋 Total registros encontrados:', historial.length);
    return historial;

  } catch (error) {
    console.error('💥 Error crítico en getHistorialCiclo:', error);
    return [];
  }
}

// Guardar seguimiento de ciclo - VERSIÓN MEJORADA
function guardarSeguimientoCiclo(seguimiento) {
  try {
    console.log('💾 Guardando seguimiento para ciclo:', seguimiento.cicloId);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Guardar en historial
    let historialSheet = ss.getSheetByName('CICLOS_HISTORIAL');
    if (!historialSheet) {
      console.log('📝 Creando nueva hoja de historial');
      historialSheet = ss.insertSheet('CICLOS_HISTORIAL');
      historialSheet.appendRow(['ID Ciclo', 'Fecha', 'Estado', 'Comentario', 'Autor']);
      historialSheet.getRange(1, 1, 1, 5).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
    }

    const fechaActual = new Date();

    // Agregar registro al historial
    historialSheet.appendRow([
      seguimiento.cicloId,
      fechaActual,
      seguimiento.estado,
      seguimiento.comentario,
      seguimiento.autor || 'Usuario desconocido'
    ]);

    // Formatear la fecha en la hoja
    const lastRow = historialSheet.getLastRow();
    historialSheet.getRange(lastRow, 2).setNumberFormat('dd/mm/yyyy hh:mm');

    console.log('✅ Seguimiento guardado en historial, fila:', lastRow);

    // 2. Actualizar estado en hoja de ciclos
    let ciclosSheet = ss.getSheetByName('Reportes_Ciclos');
    if (!ciclosSheet) {
      ciclosSheet = ss.getSheetByName('CICLOS_MEJORA'); // Buscar con otro nombre
    }

    if (ciclosSheet) {
      const data = ciclosSheet.getDataRange().getValues();
      let encontrado = false;

      for (let i = 1; i < data.length; i++) {
        const rowId = String(data[i][0] || '').trim();
        if (rowId === seguimiento.cicloId) {
          // Columna 20 es el estado (índice 19)
          ciclosSheet.getRange(i + 1, 20).setValue(seguimiento.estado);
          console.log('✅ Estado actualizado en hoja de ciclos, fila:', i + 1);
          encontrado = true;
          break;
        }
      }

      if (!encontrado) {
        console.warn('⚠️ Ciclo no encontrado en hoja principal:', seguimiento.cicloId);
      }
    } else {
      console.warn('⚠️ Hoja de ciclos no encontrada');
    }

    return {
      success: true,
      message: 'Seguimiento guardado correctamente',
      detalles: {
        historialRow: lastRow,
        fecha: fechaActual.toISOString()
      }
    };

  } catch (error) {
    console.error('❌ Error en guardarSeguimientoCiclo:', error);
    return {
      success: false,
      message: 'Error al guardar seguimiento: ' + error.toString()
    };
  }
}

function getAccionesCausa(cicloId, idCausa) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccionesCausa');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const acciones = [];
    const idCausaNum = parseInt(idCausa);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = row[0] ? row[0].toString().trim() : '';
      const rowIdCausa = row[1] !== undefined ? parseInt(row[1]) : null;
      
      if (rowCicloId === cicloId.toString() && rowIdCausa === idCausaNum) {
        const accion = {
          que: row[2] || '',
          donde: row[3] || '',
          quien: row[4] || '',
          como: row[5] || '',
          cuando: row[6] ? new Date(row[6]).toISOString() : null,
          cuanto: row[7] || '',
          estado: row[8] || 'pendiente',
          fechaCreacion: row[9] ? new Date(row[9]).toISOString() : new Date().toISOString(),
          autor: row[10] || 'Sistema'
        };
        
        acciones.push(accion);
      }
    }
    
    console.log(`Acciones encontradas para Ciclo:${cicloId}, Causa:${idCausa}:`, acciones.length);
    return acciones;
    
  } catch (error) {
    console.error('Error en getAccionesCausa:', error);
    return [];
  }
}

function guardarAccionCausa(accionData) {
  try {
    console.log('Guardando acción 5W+2H:', accionData);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('AccionesCausa');
    
    // Si no existe la hoja, crearla con nuevas columnas
    if (!sheet) {
      console.log('Creando nueva hoja AccionesCausa con estructura 5W+2H');
      sheet = ss.insertSheet('AccionesCausa');
      
      const headers = [
        'Ciclo ID',
        'ID Causa',
        'QUÉ (Descripción)',
        'DÓNDE',
        'QUIÉN (Responsable)',
        'CÓMO',
        'CUÁNDO (Fecha Límite)',
        'CUÁNTO (Recursos/Costo)',
        'Estado',
        'Fecha Creación',
        'Autor'
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // Crear nueva fila con todos los campos
    const newRow = [
      accionData.cicloId,
      parseInt(accionData.idCausa),
      accionData.que || '',
      accionData.donde || '',
      accionData.quien || '',
      accionData.como || '',
      accionData.cuando || '',
      accionData.cuanto || '',
      accionData.estado || 'pendiente',
      new Date(),
      accionData.autor || 'Usuario'
    ];
    
    console.log('Nueva fila 5W+2H:', newRow);
    
    // Agregar fila
    sheet.appendRow(newRow);
    
    return {
      success: true,
      message: 'Acción 5W+2H guardada correctamente',
      accion: accionData
    };
    
  } catch (error) {
    console.error('Error en guardarAccionCausa:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}


function eliminarAccionCausa(cicloId, idCausa, indexAccion) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccionesCausa');
    if (!sheet) return { success: false, message: 'Hoja no encontrada' };
    
    const data = sheet.getDataRange().getValues();
    let contador = 0;
    let filaAEliminar = -1;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === cicloId && row[1] == idCausa) {
        if (contador == indexAccion) {
          filaAEliminar = i + 1; // +1 porque getValues() empieza en 0 pero las filas en 1
          break;
        }
        contador++;
      }
    }
    
    if (filaAEliminar > 0) {
      sheet.deleteRow(filaAEliminar);
      return { success: true, message: 'Acción eliminada' };
    }
    
    return { success: false, message: 'Acción no encontrada' };
  } catch (error) {
    console.error('Error en eliminarAccionCausa:', error);
    return { success: false, message: error.toString() };
  }
}

function getContadorAccionesCausa(cicloId, idCausa) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccionesCausa');
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    
    let contador = 0;
    const idCausaNum = parseInt(idCausa);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = row[0] || '';
      const rowIdCausa = parseInt(row[1]) || 0;
      
      if (rowCicloId.toString() === cicloId.toString() && rowIdCausa === idCausaNum) {
        contador++;
      }
    }
    
    return contador;
  } catch (error) {
    console.error('Error en getContadorAccionesCausa:', error);
    return 0;
  }
}

function crearHojaAccionesCausa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('AccionesCausa');
  
  const headers = [
    'Ciclo ID',
    'ID Causa',
    'Descripción',
    'Responsable',
    'Fecha Límite',
    'Estado',
    'Fecha Creación',
    'Autor'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  return sheet;
}

// ========== FUNCIONES SEGUIMIENTO POR ACCIÓN-CAUSA ==========

/**
 * Obtiene el historial de seguimiento de una acción específica dentro de una causa
 */
function getHistorialAccionCausa(cicloId, causaIndex) {
  try {
    console.log('🔍 Buscando historial para acción-causa - Ciclo:', cicloId, 'CausaIndex:', causaIndex);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('CICLOS_HISTORIAL');
    
    if (!sheet) {
      console.log('📭 Hoja de historial no encontrada');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    console.log('📊 Datos en hoja de historial:', data.length, 'filas');
    
    if (data.length <= 1) {
      console.log('📭 Hoja de historial vacía o solo encabezados');
      return [];
    }
    
    const historial = [];
    const causaIndexNum = parseInt(causaIndex);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();
      const rowCausaIndex = row[1] !== undefined ? parseInt(row[1]) : -1;
      
      // Buscar registros para este ciclo Y esta causa específica
      // Si row[1] está vacío, significa que es un seguimiento general del ciclo
      // Si row[1] tiene valor, es un seguimiento específico de una causa
      if (rowCicloId === cicloId && rowCausaIndex === causaIndexNum) {
        const registro = {
          cicloId: rowCicloId,
          causaIndex: rowCausaIndex,
          fecha: row[2] ? row[2].toISOString() : new Date().toISOString(),
          estado: String(row[3] || ''),
          comentario: String(row[4] || ''),
          autor: String(row[5] || 'Sistema')
        };
        
        console.log('✅ Registro encontrado para acción-causa:', registro);
        historial.push(registro);
      }
    }
    
    console.log('📋 Total registros encontrados para acción-causa:', historial.length);
    
    // Ordenar por fecha descendente (más reciente primero)
    historial.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
    
    return historial;
    
  } catch (error) {
    console.error('💥 Error crítico en getHistorialAccionCausa:', error);
    return [];
  }
}

/**
 * Guarda un seguimiento para una acción específica dentro de una causa
 */
function guardarSeguimientoAccionCausa(seguimiento) {
  try {
    console.log('💾 Guardando seguimiento para acción-causa:', seguimiento);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Verificar y preparar la hoja CICLOS_HISTORIAL
    let historialSheet = ss.getSheetByName('CICLOS_HISTORIAL');
    if (!historialSheet) {
      console.log('📝 Creando nueva hoja de historial');
      historialSheet = ss.insertSheet('CICLOS_HISTORIAL');
      // Actualizar encabezados para incluir Causa Index
      historialSheet.appendRow(['Ciclo ID', 'Causa Index', 'Fecha', 'Estado', 'Comentario', 'Autor', 'Tipo']);
      historialSheet.getRange(1, 1, 1, 7).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
    } else {
      // Verificar si la hoja tiene los encabezados correctos
      const headers = historialSheet.getRange(1, 1, 1, historialSheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Causa Index')) {
        // Agregar columna Causa Index si no existe
        historialSheet.insertColumnAfter(1);
        historialSheet.getRange(1, 2).setValue('Causa Index');
        historialSheet.getRange(1, 1, 1, historialSheet.getLastColumn()).setBackground('#0f307f').setFontColor('#ffffff').setFontWeight('bold');
      }
    }
    
    const fechaActual = new Date();
    
    // Agregar registro al historial
    historialSheet.appendRow([
      seguimiento.cicloId,
      parseInt(seguimiento.causaIndex),
      fechaActual,
      seguimiento.estado,
      seguimiento.comentario,
      seguimiento.autor || 'Usuario desconocido',
      'Acción-Causa' // Tipo para diferenciar de seguimientos generales del ciclo
    ]);
    
    // Formatear la fecha en la hoja
    const lastRow = historialSheet.getLastRow();
    historialSheet.getRange(lastRow, 3).setNumberFormat('dd/mm/yyyy hh:mm');
    
    console.log('✅ Seguimiento de acción-causa guardado en CICLOS_HISTORIAL, fila:', lastRow);
    
    // 2. También necesitamos modificar la función getHistorialCiclo para que ignore registros de acción-causa
    // (Eso se hará en esa función)
    
    return {
      success: true,
      message: 'Seguimiento guardado correctamente',
      detalles: {
        historialRow: lastRow,
        fecha: fechaActual.toISOString()
      }
    };
    
  } catch (error) {
    console.error('❌ Error en guardarSeguimientoAccionCausa:', error);
    return {
      success: false,
      message: 'Error al guardar seguimiento: ' + error.toString()
    };
  }
}

/**
 * Obtiene el estado actual de una acción específica
 */
function getEstadoAccionCausa(cicloId, causaIndex) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('CICLOS_HISTORIAL');
    
    if (!sheet) return 'Abierto';
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 'Abierto';
    
    const causaIndexNum = parseInt(causaIndex);
    let ultimoEstado = 'Abierto';
    
    // Buscar el último estado registrado para esta acción
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowCicloId = String(row[0] || '').trim();
      const rowCausaIndex = row[1] !== undefined ? parseInt(row[1]) : -1;
      
      // Solo considerar registros específicos de esta causa
      if (rowCicloId === cicloId && rowCausaIndex === causaIndexNum) {
        ultimoEstado = String(row[3] || 'Abierto');
        break;
      }
    }
    
    return ultimoEstado;
    
  } catch (error) {
    console.error('Error en getEstadoAccionCausa:', error);
    return 'Abierto';
  }
}

