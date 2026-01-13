function revisarRegistrosDatos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Datos");
  
  if (!hoja) {
    Logger.log("Hoja 'Datos' no encontrada.");
    return;
  }

  var lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay filas con datos en 'Datos'.");
    return;
  }

  var columnaK = 11; 
  var datos = hoja.getRange(2, columnaK, lastRow - 1, 1).getValues();

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
        filasEncontradas.push({ row: r + 2, fecha: Utilities.formatDate(fecha, tz, "yyyy-MM-dd HH:mm:ss") });
      }
    } else if (cell !== "") {
      filasNoParseables.push({ row: r + 2, raw: cell });
    }
  }

  var urlHoja = ss.getUrl();

  if (!huboRegistros) {
    var destinatariosJefes = [
      "anamilena.roa@segurosbolivar.com",
      "diana.cordoba@segurosbolivar.com",
      "nohora.jaimes@segurosbolivar.com",
      "sebastian.daza@segurosbolivar.com"
    ];

    var asunto = "Alerta: No hubo leads en RENTA VOLUNTARIA: DATOS el " + fechaAyerStr;
    var mensaje = "No se detectaron registros en la hoja 'Datos' el " + fechaAyerStr + ".\n\n" +
                  "Puedes revisar la hoja aquí: " + urlHoja;

    MailApp.sendEmail(destinatariosJefes.join(","), asunto, mensaje);

    // Enviar copia a Rebeca
    var copiaAsunto = "[COPIA] " + asunto;
    var mensajeRebe = mensaje + "\n\n" +
                      (filasNoParseables.length ? "Atención: hay filas con formato de fecha no reconocido:\n" + JSON.stringify(filasNoParseables) : "");
    MailApp.sendEmail("rebeca.pedrozo@segurosbolivar.com", copiaAsunto, mensajeRebe);

  } else {
    var asuntoExito = "Reporte: Sí hubo leads en Datos el " + fechaAyerStr;
    var mensajeExito = "Se detectaron " + filasEncontradas.length + " registros en la hoja 'Datos' el " + fechaAyerStr + ".\n\n" +
                       "Ejemplos:\n" + JSON.stringify(filasEncontradas.slice(0,5)) + "\n\n" +
                       (filasNoParseables.length ? "También hubo filas no parseables:\n" + JSON.stringify(filasNoParseables) + "\n\n" : "") +
                       "Puedes revisar la hoja aquí: " + urlHoja;

    MailApp.sendEmail("rebeca.pedrozo@segurosbolivar.com", asuntoExito, mensajeExito);
  }
}