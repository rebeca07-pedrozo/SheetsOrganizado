function revisarRegistrosDatos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Datos");  
  if (!hoja) {
    Logger.log("Hoja no encontrada.");
    return;
  }

  var lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay filas con datos.");
    return;
  }

  var columnaFecha = 16; 
  var datos = hoja.getRange(2, columnaFecha, lastRow - 1, 1).getValues();

  var tz = ss.getSpreadsheetTimeZone();
  var hoy = new Date();
  var ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);
  
  var fechaAyerStr = Utilities.formatDate(ayer, tz, "yyyy-MM-dd");
  
  var huboRegistros = false;
  var filasEncontradas = [];
  var filasNoParseables = [];

  for (var r = 0; r < datos.length; r++) {
    var cell = datos[r][0];
    
    if (cell === "" || cell === null || cell === undefined) continue;

    var fecha = null;

    if (cell instanceof Date && !isNaN(cell.getTime())) {
      fecha = cell;
    } else if (typeof cell === 'string' && cell.trim() !== '') {
      var s = cell.trim().replace(/\//g, '-');
      var parsed = Date.parse(s);
      if (!isNaN(parsed)) fecha = new Date(parsed);
    }

    if (fecha && !isNaN(fecha.getTime())) {
      var fechaCeldaStr = Utilities.formatDate(fecha, tz, "yyyy-MM-dd");
      
      if (fechaCeldaStr === fechaAyerStr) {
        huboRegistros = true;
        filasEncontradas.push({ 
          fila_excel: r + 2, 
          valor: cell 
        });
      }
    } else {
      filasNoParseables.push({ fila_excel: r + 2, contenido: cell });
    }
  }

  var urlHoja = ss.getUrl();
  var destinatariosJefes = [
    "anamilena.roa@segurosbolivar.com",
    "diana.cordoba@segurosbolivar.com",
    "nohora.jaimes@segurosbolivar.com",
    "sebastian.daza@segurosbolivar.com"
  ];

  if (!huboRegistros) {
    var asuntoJefes = "Alerta: No hubo registros en datos HOGAR el " + fechaAyerStr;
    var mensajeJefes = "Hola,\n\nNo se detectaron nuevos registros en la hoja 'Datos' durante el día de ayer (" + fechaAyerStr + ").\n\n" +
                       "Puedes revisar el archivo aquí: " + urlHoja;

    MailApp.sendEmail(destinatariosJefes.join(","), asuntoJefes, mensajeJefes);

    var asuntoRebeCopia = "[COPIA ALERTA] " + asuntoJefes;
    var mensajeRebeCopia = mensajeJefes + "\n\n--- Notas Técnicas ---\n" +
                           (filasNoParseables.length ? "Se encontraron celdas con texto no reconocido como fecha en estas filas: " + JSON.stringify(filasNoParseables) : "No hay basura en la columna.");
    
    MailApp.sendEmail("rebeca.pedrozo@segurosbolivar.com", asuntoRebeCopia, mensajeRebeCopia);

  } else {
    var asuntoExito = "Reporte: Sí hubo registros en en datos HOGAR el " + fechaAyerStr;
    var mensajeExito = "Se detectaron " + filasEncontradas.length + " registros de ayer en la columna O.\n\n" +
                       "DETALLE DE FILAS ENCONTRADAS:\n" + 
                       JSON.stringify(filasEncontradas, null, 2) + "\n\n" +
                       "Si crees que esto es un error, revisa las filas mencionadas arriba en el Excel.\n\n" +
                       "Archivo: " + urlHoja;

    MailApp.sendEmail("rebeca.pedrozo@segurosbolivar.com", asuntoExito, mensajeExito);
  }
}