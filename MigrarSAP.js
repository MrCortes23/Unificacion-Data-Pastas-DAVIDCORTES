function migrarSAP() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = spreadsheet.getSheetByName('Reportes_Tarjetas');
  var hojaDestino = spreadsheet.getSheetByName('SAP');
  
  if (!hojaDestino) {
    hojaDestino = spreadsheet.insertSheet('SAP');
  }
  
  var datos = hojaOrigen.getDataRange().getValues();
  var filasFiltradas = [];

  // Encontrar índices de columnas necesarias
  var headers = datos[0];
  var requiereSAPIndex = headers.indexOf("RequiereSAP");
  var estadoIndex = headers.indexOf("estado");
  var numSAPIndex = headers.indexOf("NumSAP");
  
  // Verificar que las columnas existen
  if (requiereSAPIndex === -1) {
    Logger.log("❌ ERROR: No se encontró la columna 'RequiereSAP'");
    return;
  }
  if (estadoIndex === -1) {
    Logger.log("❌ ERROR: No se encontró la columna 'estado'");
    return;
  }
  if (numSAPIndex === -1) {
    Logger.log("❌ ERROR: No se encontró la columna 'NumSAP'");
    return;
  }
  
  Logger.log("🔍 Índices encontrados:");
  Logger.log("   RequiereSAP: " + requiereSAPIndex);
  Logger.log("   estado: " + estadoIndex);
  Logger.log("   NumSAP: " + numSAPIndex);
  
  for (var i = 0; i < datos.length; i++) {
    if (i === 0) {
      // Siempre agregar encabezados
      filasFiltradas.push(datos[i]);
    } else {
      var fila = datos[i];
      var requiereSAP = fila[requiereSAPIndex];
      var estado = fila[estadoIndex];
      var numSAP = fila[numSAPIndex];
      
      // Convertir a strings para comparación segura
      var requiereSAPStr = requiereSAP ? requiereSAP.toString().trim().toLowerCase() : "";
      var estadoStr = estado ? estado.toString().trim().toLowerCase() : "";
      var numSAPStr = numSAP ? numSAP.toString().trim() : "";
      
      // Verificar las 3 condiciones
      var cumpleCondiciones = 
        requiereSAPStr === "si" &&            // Condición 1: Requiere SAP = "Si"
        (estadoStr === "abierta" || estadoStr === "abierto") &&  // Condición 2: Estado = Abierta/Abierto
        numSAPStr === "";                     // Condición 3: NumSAP está vacío
      
      if (cumpleCondiciones) {
        filasFiltradas.push(fila);
        
        // Log detallado para debugging
        Logger.log("✅ FILA " + i + " CUMPLE:");
        Logger.log("   RequiereSAP: '" + requiereSAPStr + "'");
        Logger.log("   estado: '" + estadoStr + "'");
        Logger.log("   NumSAP: '" + numSAPStr + "'");
      }
    }
  }
  
  // Limpiar hoja destino
  hojaDestino.clear();
  
  // Escribir datos filtrados si hay algo
  if (filasFiltradas.length > 0) {
    hojaDestino.getRange(1, 1, filasFiltradas.length, filasFiltradas[0].length).setValues(filasFiltradas);
    
    // Aplicar formato para mejor visualización
    hojaDestino.getRange(1, 1, 1, filasFiltradas[0].length).setFontWeight("bold");
    hojaDestino.getRange(1, 1, 1, filasFiltradas[0].length).setBackground("#e8f0fe");
    hojaDestino.autoResizeColumns(1, filasFiltradas[0].length);
  }
  
  // Resumen
  Logger.log('======================================');
  Logger.log('✅ MIGRACIÓN COMPLETADA');
  Logger.log('📊 TOTAL: ' + (filasFiltradas.length - 1) + ' filas copiadas a SAP');
  Logger.log('📄 Hoja original Reportes_Tarjetas permanece intacta');
  Logger.log('======================================');
  
  // Mostrar mensaje al usuario
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Migración completada: ' + (filasFiltradas.length - 1) + ' filas copiadas a SAP',
    'Migración SAP',
    5
  );
}