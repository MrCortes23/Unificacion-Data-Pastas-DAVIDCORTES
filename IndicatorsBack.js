
function getTarjetasStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const tipoRiesgoIndex = headers.indexOf('Tipo_Riesgo');
    const problemaIndex = headers.indexOf('Problema_asociado');
    const estadoIndex = headers.indexOf('estado');
    
    if (tipoRiesgoIndex === -1 || problemaIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas
    const stats = {};
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tipoRiesgo = row[tipoRiesgoIndex] || 'Sin tipo';
      const problema = row[problemaIndex] || 'Sin problema';
      const estado = row[estadoIndex] || 'Sin estado';
      
      // Inicializar estructura si no existe
      if (!stats[tipoRiesgo]) {
        stats[tipoRiesgo] = {
          problemas: {},
          totalAbierta: 0,
          totalCerrada: 0,
          totalGeneral: 0
        };
      }
      
      if (!stats[tipoRiesgo].problemas[problema]) {
        stats[tipoRiesgo].problemas[problema] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[tipoRiesgo].problemas[problema].abierta++;
        stats[tipoRiesgo].totalAbierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[tipoRiesgo].problemas[problema].cerrada++;
        stats[tipoRiesgo].totalCerrada++;
      }
      
      stats[tipoRiesgo].problemas[problema].total++;
      stats[tipoRiesgo].totalGeneral++;
    }
    
    // Calcular porcentajes y preparar datos para tabla
    const tableData = [];
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    Object.keys(stats).sort().forEach(tipoRiesgo => {
      const tipoData = stats[tipoRiesgo];
      let tipoAbierta = 0;
      let tipoCerrada = 0;
      let tipoTotal = 0;
      
      // Agregar filas para cada problema
      Object.keys(tipoData.problemas).sort().forEach(problema => {
        const probData = tipoData.problemas[problema];
        const porcentaje = probData.total > 0 ? 
          ((probData.cerrada / probData.total) * 100).toFixed(2) : '0.00';
        
        tableData.push({
          tipoRiesgo: problema === 'Total' ? '' : tipoRiesgo,
          problema: problema,
          abierta: probData.abierta,
          cerrada: probData.cerrada,
          total: probData.total,
          porcentaje: porcentaje + '%'
        });
        
        tipoAbierta += probData.abierta;
        tipoCerrada += probData.cerrada;
        tipoTotal += probData.total;
      });
      
      // Agregar total por tipo de riesgo
      const tipoPorcentaje = tipoTotal > 0 ? 
        ((tipoCerrada / tipoTotal) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        tipoRiesgo: tipoRiesgo,
        problema: 'Total ' + tipoRiesgo,
        abierta: tipoAbierta,
        cerrada: tipoCerrada,
        total: tipoTotal,
        porcentaje: tipoPorcentaje + '%',
        isTotal: true
      });
      
      totalGlobalAbierta += tipoAbierta;
      totalGlobalCerrada += tipoCerrada;
      totalGlobal += tipoTotal;
    });
    
    // Agregar total global
    const globalPorcentaje = totalGlobal > 0 ? 
      ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
    
    tableData.push({
      tipoRiesgo: '',
      problema: 'Suma total',
      abierta: totalGlobalAbierta,
      cerrada: totalGlobalCerrada,
      total: totalGlobal,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });
    
    return {
      success: true,
      data: tableData,
      summary: {
        totalTarjetas: totalGlobal,
        abiertas: totalGlobalAbierta,
        cerradas: totalGlobalCerrada,
        porcentajeGestion: globalPorcentaje
      }
    };
    
  } catch (error) {
    console.error('Error en getTarjetasStats:', error);
    return { success: false, message: error.toString() };
  }
}

function exportTarjetasStatsToCSV() {
  try {
    const stats = getTarjetasStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Tipo de Riesgo,Problema Asociado,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      csvContent += `"${row.tipoRiesgo || ''}","${row.problema || ''}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${row.porcentaje || '0.00%'}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportTarjetasStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

function getResponsableCargoStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const responsableIndex = headers.indexOf('Responsable_Solucion');
    const estadoIndex = headers.indexOf('estado');
    
    if (responsableIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas
    const stats = {};
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const responsable = row[responsableIndex] || 'Sin responsable';
      const estado = row[estadoIndex] || 'Sin estado';
      
      // Inicializar estructura si no existe
      if (!stats[responsable]) {
        stats[responsable] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[responsable].abierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[responsable].cerrada++;
      }
      
      stats[responsable].total++;
    }
    
    // Preparar datos para tabla
    const tableData = [];
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    Object.keys(stats).sort().forEach(responsable => {
      const responsableData = stats[responsable];
      const porcentaje = responsableData.total > 0 ? 
        ((responsableData.cerrada / responsableData.total) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        responsable: responsable,
        abierta: responsableData.abierta,
        cerrada: responsableData.cerrada,
        total: responsableData.total,
        porcentaje: porcentaje + '%'
      });
      
      totalGlobalAbierta += responsableData.abierta;
      totalGlobalCerrada += responsableData.cerrada;
      totalGlobal += responsableData.total;
    });
    
    // Agregar total global
    const globalPorcentaje = totalGlobal > 0 ? 
      ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
    
    tableData.push({
      responsable: 'Suma total',
      abierta: totalGlobalAbierta,
      cerrada: totalGlobalCerrada,
      total: totalGlobal,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });
    
    return {
      success: true,
      data: tableData
    };
    
  } catch (error) {
    console.error('Error en getResponsableCargoStats:', error);
    return { success: false, message: error.toString() };
  }
}

// Función para estadísticas por líder de solución
function getLiderSolucionStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const liderIndex = headers.indexOf('Responsable_Solucion_Nombre_Visualizar_Reporte');
    const estadoIndex = headers.indexOf('estado');
    
    if (liderIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas
    const stats = {};
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const lider = row[liderIndex] || 'Sin líder';
      const estado = row[estadoIndex] || 'Sin estado';
      
      // Inicializar estructura si no existe
      if (!stats[lider]) {
        stats[lider] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[lider].abierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[lider].cerrada++;
      }
      
      stats[lider].total++;
    }
    
    // Preparar datos para tabla
    const tableData = [];
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    Object.keys(stats).sort().forEach(lider => {
      const liderData = stats[lider];
      const porcentaje = liderData.total > 0 ? 
        ((liderData.cerrada / liderData.total) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        lider: lider,
        abierta: liderData.abierta,
        cerrada: liderData.cerrada,
        total: liderData.total,
        porcentaje: porcentaje + '%'
      });
      
      totalGlobalAbierta += liderData.abierta;
      totalGlobalCerrada += liderData.cerrada;
      totalGlobal += liderData.total;
    });
    
    // Agregar total global
    const globalPorcentaje = totalGlobal > 0 ? 
      ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
    
    tableData.push({
      lider: 'Suma total',
      abierta: totalGlobalAbierta,
      cerrada: totalGlobalCerrada,
      total: totalGlobal,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });
    
    return {
      success: true,
      data: tableData
    };
    
  } catch (error) {
    console.error('Error en getLiderSolucionStats:', error);
    return { success: false, message: error.toString() };
  }
}

function getZonaRiesgoStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const zonaIndex = headers.indexOf('Zona_Riesgo');
    const estadoIndex = headers.indexOf('estado');
    
    if (zonaIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas
    const stats = {};
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const zona = row[zonaIndex] || 'Sin zona';
      const estado = row[estadoIndex] || 'Sin estado';
      
      // Inicializar estructura si no existe
      if (!stats[zona]) {
        stats[zona] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[zona].abierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[zona].cerrada++;
      }
      
      stats[zona].total++;
    }
    
    // Preparar datos para tabla
    const tableData = [];
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    Object.keys(stats).sort().forEach(zona => {
      const zonaData = stats[zona];
      const porcentaje = zonaData.total > 0 ? 
        ((zonaData.cerrada / zonaData.total) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        zonaRiesgo: zona,
        abierta: zonaData.abierta,
        cerrada: zonaData.cerrada,
        total: zonaData.total,
        porcentaje: porcentaje + '%'
      });
      
      totalGlobalAbierta += zonaData.abierta;
      totalGlobalCerrada += zonaData.cerrada;
      totalGlobal += zonaData.total;
    });
    
    // Agregar total global
    const globalPorcentaje = totalGlobal > 0 ? 
      ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
    
    tableData.push({
      zonaRiesgo: 'Suma total',
      abierta: totalGlobalAbierta,
      cerrada: totalGlobalCerrada,
      total: totalGlobal,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });
    
    return {
      success: true,
      data: tableData
    };
    
  } catch (error) {
    console.error('Error en getZonaRiesgoStats:', error);
    return { success: false, message: error.toString() };
  }
}

// Función para estadísticas por tipo de reporte
function getTipoReporteStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const tipoIndex = headers.indexOf('Generada_por');
    const estadoIndex = headers.indexOf('estado');
    
    if (tipoIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas
    const stats = {};
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tipo = row[tipoIndex] || 'Sin tipo';
      const estado = row[estadoIndex] || 'Sin estado';
      
      // Inicializar estructura si no existe
      if (!stats[tipo]) {
        stats[tipo] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[tipo].abierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[tipo].cerrada++;
      }
      
      stats[tipo].total++;
    }
    
    // Preparar datos para tabla
    const tableData = [];
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    Object.keys(stats).sort().forEach(tipo => {
      const tipoData = stats[tipo];
      const porcentaje = tipoData.total > 0 ? 
        ((tipoData.cerrada / tipoData.total) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        tipoReporte: tipo || 'Sin especificar',
        abierta: tipoData.abierta,
        cerrada: tipoData.cerrada,
        total: tipoData.total,
        porcentaje: porcentaje + '%'
      });
      
      totalGlobalAbierta += tipoData.abierta;
      totalGlobalCerrada += tipoData.cerrada;
      totalGlobal += tipoData.total;
    });
    
    // Agregar total global
    const globalPorcentaje = totalGlobal > 0 ? 
      ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
    
    tableData.push({
      tipoReporte: 'Suma total',
      abierta: totalGlobalAbierta,
      cerrada: totalGlobalCerrada,
      total: totalGlobal,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });
    
    return {
      success: true,
      data: tableData
    };
    
  } catch (error) {
    console.error('Error en getTipoReporteStats:', error);
    return { success: false, message: error.toString() };
  }
}

// Función para estadísticas mensuales
function getMensualStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const fechaIndex = headers.indexOf('Fecha_Creacion');
    const estadoIndex = headers.indexOf('estado');
    
    if (fechaIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas por año-mes
    const stats = {};
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fechaStr = row[fechaIndex];
      const estado = row[estadoIndex] || 'Sin estado';
      
      let fecha;
      try {
        fecha = new Date(fechaStr);
      } catch (e) {
        continue; // Saltar fechas inválidas
      }
      
      if (!fecha || isNaN(fecha.getTime())) {
        continue;
      }
      
      const year = fecha.getFullYear();
      const month = fecha.getMonth() + 1; // Enero = 1
      const key = `${year}-${month.toString().padStart(2, '0')}`;
      
      // Inicializar estructura si no existe
      if (!stats[key]) {
        stats[key] = {
          year: year,
          month: month,
          mesNombre: getMonthName(month),
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[key].abierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[key].cerrada++;
      }
      
      stats[key].total++;
    }
    
    // Preparar datos para tabla
    const tableData = [];
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    // Agrupar por año
    const statsByYear = {};
    
    // Primero agregar datos mensuales ordenados
    Object.keys(stats).sort().forEach(key => {
      const monthData = stats[key];
      const porcentaje = monthData.total > 0 ? 
        ((monthData.cerrada / monthData.total) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        ano: monthData.year,
        mes: monthData.mesNombre,
        abierta: monthData.abierta,
        cerrada: monthData.cerrada,
        total: monthData.total,
        porcentaje: porcentaje + '%'
      });
      
      // Acumular por año
      if (!statsByYear[monthData.year]) {
        statsByYear[monthData.year] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      statsByYear[monthData.year].abierta += monthData.abierta;
      statsByYear[monthData.year].cerrada += monthData.cerrada;
      statsByYear[monthData.year].total += monthData.total;
      
      totalGlobalAbierta += monthData.abierta;
      totalGlobalCerrada += monthData.cerrada;
      totalGlobal += monthData.total;
    });
    
    // Agregar totales por año
    Object.keys(statsByYear).sort().forEach(year => {
      const yearData = statsByYear[year];
      const porcentaje = yearData.total > 0 ? 
        ((yearData.cerrada / yearData.total) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        ano: `Total ${year}`,
        mes: '',
        abierta: yearData.abierta,
        cerrada: yearData.cerrada,
        total: yearData.total,
        porcentaje: porcentaje + '%',
        isYearTotal: true
      });
    });
    
    // Agregar total global
    const globalPorcentaje = totalGlobal > 0 ? 
      ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
    
    tableData.push({
      ano: '',
      mes: 'Suma total',
      abierta: totalGlobalAbierta,
      cerrada: totalGlobalCerrada,
      total: totalGlobal,
      porcentaje: globalPorcentaje + '%',
      isGlobalTotal: true
    });
    
    return {
      success: true,
      data: tableData
    };
    
  } catch (error) {
    console.error('Error en getMensualStats:', error);
    return { success: false, message: error.toString() };
  }
}

// Función auxiliar para obtener nombre del mes
function getMonthName(monthNumber) {
  const months = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  return months[monthNumber - 1] || `Mes ${monthNumber}`;
}

// Funciones de exportación a CSV
function exportZonaRiesgoStatsToCSV() {
  try {
    const stats = getZonaRiesgoStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Zona de Riesgo,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const zona = (row.zonaRiesgo || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${zona}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportZonaRiesgoStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

function exportTipoReporteStatsToCSV() {
  try {
    const stats = getTipoReporteStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Tipo de Reporte,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const tipo = (row.tipoReporte || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${tipo}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportTipoReporteStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

function exportMensualStatsToCSV() {
  try {
    const stats = getMensualStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Año,Mes,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const ano = (row.ano || '').replace(/"/g, '""');
      const mes = (row.mes || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${ano}","${mes}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportMensualStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

// Funciones de exportación existentes
function exportTarjetasStatsToCSV() {
  try {
    const stats = getTarjetasStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Tipo de Riesgo,Problema Asociado,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const tipoRiesgo = (row.tipoRiesgo || '').replace(/"/g, '""');
      const problema = (row.problema || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${tipoRiesgo}","${problema}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportTarjetasStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

function exportResponsableCargoStatsToCSV() {
  try {
    const stats = getResponsableCargoStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Responsable Solución,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const responsable = (row.responsable || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${responsable}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportResponsableCargoStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

function exportLiderSolucionStatsToCSV() {
  try {
    const stats = getLiderSolucionStats();
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Líder de Solución,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const lider = (row.lider || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${lider}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportLiderSolucionStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

// Función para obtener años disponibles
function getAvailableYearsForStats() {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índice de columna de fecha
    const fechaIndex = headers.indexOf('Fecha_Creacion');
    
    if (fechaIndex === -1) {
      return { success: false, message: 'Columna de fecha no encontrada' };
    }
    
    const yearsSet = new Set();
    
    // Recorrer datos para extraer años
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fechaStr = row[fechaIndex];
      
      if (fechaStr) {
        let fecha;
        try {
          fecha = new Date(fechaStr);
        } catch (e) {
          continue;
        }
        
        if (fecha && !isNaN(fecha.getTime())) {
          const year = fecha.getFullYear();
          yearsSet.add(year.toString());
        }
      }
    }
    
    const years = Array.from(yearsSet).sort((a, b) => b - a); // Orden descendente
    
    return {
      success: true,
      years: years
    };
    
  } catch (error) {
    console.error('Error en getAvailableYearsForStats:', error);
    return { success: false, message: error.toString() };
  }
}

// Versión mejorada que maneja registros sin fecha
function getRiesgoAnualStats(year) {
  try {
    const ss = SpreadsheetApp.openById('1QIUKYX42uuMlsssR-0CizPI-lJwDS6xH760kg9uYDII');
    const sheet = ss.getSheetByName('Reportes_Tarjetas'); 
    
    if (!sheet) {
      return { success: false, message: 'Hoja no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Encontrar índices de columnas
    const fechaIndex = headers.indexOf('Fecha_Creacion');
    const tipoRiesgoIndex = headers.indexOf('Tipo_Riesgo');
    const problemaIndex = headers.indexOf('Problema_asociado');
    const estadoIndex = headers.indexOf('estado');
    
    if (fechaIndex === -1 || tipoRiesgoIndex === -1 || problemaIndex === -1 || estadoIndex === -1) {
      return { success: false, message: 'Columnas requeridas no encontradas' };
    }
    
    // Estructura para almacenar estadísticas
    const stats = {};
    
    // Variables para totales
    let totalGlobalAbierta = 0;
    let totalGlobalCerrada = 0;
    let totalGlobal = 0;
    
    // Procesar datos
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fechaStr = row[fechaIndex];
      const tipoRiesgo = row[tipoRiesgoIndex] || 'Sin tipo';
      const problema = row[problemaIndex] || 'Sin problema';
      const estado = row[estadoIndex] || 'Sin estado';
      
      // Obtener año de la fecha
      let rowYear = "Sin año";
      if (fechaStr) {
        try {
          const fecha = new Date(fechaStr);
          if (fecha && !isNaN(fecha.getTime())) {
            rowYear = fecha.getFullYear().toString();
          }
        } catch (e) {
          // Si hay error, mantener "Sin año"
        }
      }
      
      // Filtrar por año si no es "todos"
      if (year !== "todos" && rowYear !== year) {
        continue;
      }
      
      const key = `${rowYear}|${tipoRiesgo}|${problema}`;
      
      // Inicializar estructura si no existe
      if (!stats[rowYear]) {
        stats[rowYear] = {};
      }
      
      if (!stats[rowYear][tipoRiesgo]) {
        stats[rowYear][tipoRiesgo] = {
          problemas: {},
          totalAbierta: 0,
          totalCerrada: 0,
          totalGeneral: 0
        };
      }
      
      if (!stats[rowYear][tipoRiesgo].problemas[problema]) {
        stats[rowYear][tipoRiesgo].problemas[problema] = {
          abierta: 0,
          cerrada: 0,
          total: 0
        };
      }
      
      // Contar por estado
      if (estado === 'Abierta' || estado === 'Abierto') {
        stats[rowYear][tipoRiesgo].problemas[problema].abierta++;
        stats[rowYear][tipoRiesgo].totalAbierta++;
        totalGlobalAbierta++;
      } else if (estado === 'Cerrada' || estado === 'Cerrado') {
        stats[rowYear][tipoRiesgo].problemas[problema].cerrada++;
        stats[rowYear][tipoRiesgo].totalCerrada++;
        totalGlobalCerrada++;
      }
      
      stats[rowYear][tipoRiesgo].problemas[problema].total++;
      stats[rowYear][tipoRiesgo].totalGeneral++;
      totalGlobal++;
    }
    
    // Preparar datos para tabla
    const tableData = [];
    
    // Obtener años ordenados (más reciente primero, "Sin año" al final)
    const years = Object.keys(stats).sort((a, b) => {
      if (a === "Sin año") return 1;
      if (b === "Sin año") return -1;
      return b - a;
    });
    
    years.forEach(ano => {
      const tiposRiesgo = Object.keys(stats[ano]).sort();
      
      tiposRiesgo.forEach(tipoRiesgo => {
        const tipoData = stats[ano][tipoRiesgo];
        
        // Agregar filas para cada problema
        Object.keys(tipoData.problemas).sort().forEach(problema => {
          const probData = tipoData.problemas[problema];
          const porcentaje = probData.total > 0 ? 
            ((probData.cerrada / probData.total) * 100).toFixed(2) : '0.00';
          
          tableData.push({
            ano: ano,
            tipoRiesgo: tipoRiesgo,
            problema: problema,
            abierta: probData.abierta,
            cerrada: probData.cerrada,
            total: probData.total,
            porcentaje: porcentaje + '%'
          });
        });
        
        // Agregar subtotal por tipo de riesgo en el año
        const tipoPorcentaje = tipoData.totalGeneral > 0 ? 
          ((tipoData.totalCerrada / tipoData.totalGeneral) * 100).toFixed(2) : '0.00';
        
        tableData.push({
          ano: ano,
          tipoRiesgo: tipoRiesgo,
          problema: 'Total ' + tipoRiesgo,
          abierta: tipoData.totalAbierta,
          cerrada: tipoData.totalCerrada,
          total: tipoData.totalGeneral,
          porcentaje: tipoPorcentaje + '%',
          isRiesgoTotal: true
        });
      });
      
      // Calcular total por año
      let anoAbierta = 0;
      let anoCerrada = 0;
      let anoTotal = 0;
      
      tiposRiesgo.forEach(tipoRiesgo => {
        const tipoData = stats[ano][tipoRiesgo];
        anoAbierta += tipoData.totalAbierta;
        anoCerrada += tipoData.totalCerrada;
        anoTotal += tipoData.totalGeneral;
      });
      
      const anoPorcentaje = anoTotal > 0 ? 
        ((anoCerrada / anoTotal) * 100).toFixed(2) : '0.00';
      
      // Agregar total por año
      tableData.push({
        ano: ano,
        tipoRiesgo: 'TOTAL AÑO',
        problema: '',
        abierta: anoAbierta,
        cerrada: anoCerrada,
        total: anoTotal,
        porcentaje: anoPorcentaje + '%',
        isYearTotal: true
      });
    });
    
    // Agregar total global (solo si es "todos" o hay un solo año)
    if (year === "todos" || years.length > 0) {
      const globalPorcentaje = totalGlobal > 0 ? 
        ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00';
      
      tableData.push({
        ano: '',
        tipoRiesgo: 'TOTAL GENERAL',
        problema: '',
        abierta: totalGlobalAbierta,
        cerrada: totalGlobalCerrada,
        total: totalGlobal,
        porcentaje: globalPorcentaje + '%',
        isGlobalTotal: true
      });
    }
    
    return {
      success: true,
      data: tableData,
      summary: {
        totalTarjetas: totalGlobal,
        abiertas: totalGlobalAbierta,
        cerradas: totalGlobalCerrada,
        porcentajeGestion: totalGlobal > 0 ? 
          ((totalGlobalCerrada / totalGlobal) * 100).toFixed(2) : '0.00'
      }
    };
    
  } catch (error) {
    console.error('Error en getRiesgoAnualStats:', error);
    return { success: false, message: error.toString() };
  }
}

// Función para exportar estadísticas de riesgo anual a CSV
function exportRiesgoAnualStatsToCSV(year) {
  try {
    const stats = getRiesgoAnualStats(year);
    
    if (!stats.success) {
      return { success: false, message: stats.message };
    }
    
    // Crear contenido CSV
    let csvContent = "Año,Tipo de Riesgo,Problema Asociado,Abierta,Cerrada,Suma total,% Gestión\n";
    
    stats.data.forEach(row => {
      const ano = (row.ano || '').replace(/"/g, '""');
      const tipoRiesgo = (row.tipoRiesgo || '').replace(/"/g, '""');
      const problema = (row.problema || '').replace(/"/g, '""');
      const porcentaje = (row.porcentaje || '0.00%').replace(',', '.');
      
      csvContent += `"${ano}","${tipoRiesgo}","${problema}",${row.abierta || 0},${row.cerrada || 0},${row.total || 0},"${porcentaje}"\n`;
    });
    
    return {
      success: true,
      csv: csvContent
    };
    
  } catch (error) {
    console.error('Error en exportRiesgoAnualStatsToCSV:', error);
    return { success: false, message: error.toString() };
  }
}