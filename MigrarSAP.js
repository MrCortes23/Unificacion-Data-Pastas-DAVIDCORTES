function migrarSAP() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = spreadsheet.getSheetByName('Reportes_Tarjetas');
  var hojaDestino = spreadsheet.getSheetByName('SAP');
  
  if (!hojaDestino) {
    hojaDestino = spreadsheet.insertSheet('SAP');
  }
  
  var datos = hojaOrigen.getDataRange().getValues();
  var filasFiltradas = [];

  for (var i = 0; i < datos.length; i++) {
    if (i === 0) {
      filasFiltradas.push(datos[i]);
    }
    else if (datos[i][15] && datos[i][15].toString().trim().toLowerCase() === "si") {
      filasFiltradas.push(datos[i]);
    }
  }
  
  hojaDestino.clear();
  
  if (filasFiltradas.length > 0) {
    hojaDestino.getRange(1, 1, filasFiltradas.length, filasFiltradas[0].length).setValues(filasFiltradas);
  }
  
  Logger.log('✅ Copia completada: ' + (filasFiltradas.length - 1) + ' filas copiadas a SAP');
  Logger.log('✅ Hoja original Reportes_Tarjetas permanece intacta');
}