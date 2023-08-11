var nombresHojasOrigen = [
  'Desgloce Produccion (Compras)',
  'Desgloce Produccion (Nomina)',
  'Desgloce Granja',
  'Desgloce DISTRIBUCION',
  'Desgloce Comercializacion',
  'Desgloce Administracion'
];

function aplicarTransformacionesEnCarpeta() {
  var carpetaId = '1Lm4rlmBY57vmSXQj4fcLb0INpy9E0svh'; // Reemplaza con el ID real de la carpeta
  var archivos = obtenerArchivosEnCarpeta(carpetaId);

  for (var i = 0; i < nombresHojasOrigen.length; i++) {
    var nombreHojaOrigen = nombresHojasOrigen[i];

    while (archivos.hasNext()) {
      var archivo = archivos.next();
      var idArchivo = archivo.getId();
      var archivoOrigen = SpreadsheetApp.openById(idArchivo);
      var hojaOrigen = archivoOrigen.getSheetByName(nombreHojaOrigen);
      if (hojaOrigen) {
        var tablaDatos = hojaOrigen.getDataRange().getValues();
        transformarTabla(archivoOrigen, tablaDatos);
      }
    }
    // Reiniciar el iterador para el próximo nombre de hoja
    archivos = obtenerArchivosEnCarpeta(carpetaId);
  }
}

function obtenerArchivosEnCarpeta(carpetaId) {
  var carpeta = DriveApp.getFolderById(carpetaId);
  return carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
}

function transformarTabla(archivoOrigen, tablaDatos) {
  var tablaAreas = generarTablaAreas(tablaDatos);
  var desglosePorArea = generarDesglosePorArea(tablaDatos);

  eliminarPestanasDesgloce(archivoOrigen);
  crearHojaAreasUnicas(archivoOrigen, tablaAreas);
  crearPestanasDesgloce(archivoOrigen, tablaAreas, desglosePorArea);
}

function eliminarPestanasDesgloce(archivoOrigen) {
  var pestañasDesgloce = archivoOrigen.getSheets();
  for (var i = 0; i < pestañasDesgloce.length; i++) { // Comienza desde 1 para evitar la hoja 'Áreas Únicas'
    var nombrePestana = pestañasDesgloce[i].getName();
    if (nombrePestana.startsWith('Desglose ') || nombrePestana === 'Áreas Únicas') {
      archivoOrigen.deleteSheet(pestañasDesgloce[i]);
    }
  }
}

function generarTablaAreas(tablaDatos) {
  var tablaAreas = [['Centro de Costos', 'Área', 'Clasificación para Finanzas', 'Total']];
  var areasUnicas = [];
  for (var i = 1; i < tablaDatos.length; i++) {
    var area = tablaDatos[i][2]; // Ajusta el índice a la columna donde se encuentra el nombre del área
    if (areasUnicas.indexOf(area) === -1) {
      areasUnicas.push(area);
      tablaAreas.push(['CF' + (1000 + i), area, 'Producción ( Compras )', calcularTotalPorArea(tablaDatos, area)]);
    }
  }
  return tablaAreas;
}

function calcularTotalPorArea(tablaDatos, area) {
  var total = 0;
  for (var i = 1; i < tablaDatos.length; i++) {
    if (tablaDatos[i][2] === area) {
      total += tablaDatos[i][3];
    }
  }
  return total;
}

function generarDesglosePorArea(tablaDatos) {
  var desglosePorArea = {};
  for (var i = 1; i < tablaDatos.length; i++) {
    var area = tablaDatos[i][2];
    if (!desglosePorArea[area]) {
      desglosePorArea[area] = [];
    }
    
    var plan = tablaDatos[i][5];
    var empleado = desglosePorArea[area].find(emp => emp[0] === tablaDatos[i][0]);
    if (!empleado) {
      empleado = [tablaDatos[i][0], tablaDatos[i][1], '', '', '', '', tablaDatos[i][6], tablaDatos[i][4],''];
      desglosePorArea[area].push(empleado);
    }

    if (plan === '5.5 GB') {
      empleado[2] = tablaDatos[i][3];
    } else if (plan === '7.5 GB') {
      empleado[3] = tablaDatos[i][3];
    } else if (plan === '11 GB') {
      empleado[4] = tablaDatos[i][3];
    } else if (plan === '2 GB') {
      empleado[5] = tablaDatos[i][3];
    }
  }
  return desglosePorArea;
}

function crearHojaAreasUnicas(archivoOrigen, tablaAreas) {
  var hojaNueva = archivoOrigen.insertSheet('Áreas Únicas');
  hojaNueva.getRange(2, 1, tablaAreas.length, tablaAreas[0].length).setValues(tablaAreas);
}



function crearPestanasDesgloce(archivoOrigen, tablaAreas, desglosePorArea) {
  for (var i = 1; i < tablaAreas.length; i++) {
    var area = tablaAreas[i][1];
    var hojaNueva = archivoOrigen.insertSheet('Desglose ' + area);

    // Agregar los encabezados de las columnas
    var headers = ['Teléfono', 'Nombre', '5.5 GB', '7.5 GB', '11 GB', '2 GB', 'Seguro', 'Anualidad', 'Consumo Adicional', 'Suma total', 'Observaciones'];
    hojaNueva.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Obtener los datos del desglose por área
    var datosDesglose = desglosePorArea[area];

    // Crear un nuevo array con los datos modificados
    var nuevaTabla = datosDesglose.map(function (row) {
      var sumaColumnas = row.slice(2, 9).reduce(function (a, b) {
        return a + (parseFloat(b) || 0);
      }, 0).toFixed(2);
      return [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], sumaColumnas, '-'];
    });

    // Establecer los valores en la hoja nueva
    hojaNueva.getRange(2, 1, nuevaTabla.length, nuevaTabla[0].length).setValues(nuevaTabla);

    // Fórmulas para calcular la suma total por columna
    for (var j = 2; j <= 10; j++) {
      hojaNueva.getRange(nuevaTabla.length + 2, j).setFormula('=SUM(' + hojaNueva.getRange(2, j, nuevaTabla.length, 1).getA1Notation() + ')');
    }

    // Fórmula para agregar "Suma Total" en la columna de nombre
    hojaNueva.getRange(nuevaTabla.length + 2, 2).setValue('Suma Total');

    
  }

  
}
