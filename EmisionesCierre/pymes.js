function EmisionesPymes_Leads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Ventas 8 dic"); 
  if (!hoja) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "Ventas 19 oct"');
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const encabezados = ["CC", "Mail1", "Ventas", "Fuente", "Medio", "Campaña", "Contenido anuncio", "Fecha Lead"];
  hoja.getRange(1, 16, 1, encabezados.length).setValues([encabezados]).setHorizontalAlignment("center");
  hoja.getRange(2, 16, ultimaFila - 1, encabezados.length).clearContent();

  const cedulas = hoja.getRange(2, 11, ultimaFila - 1, 1).getValues().map(r => r[0] ? String(r[0]).trim() : "");
  const correos = hoja.getRange(2, 13, ultimaFila - 1, 1).getValues().map(r => r[0] ? String(r[0]).trim().toLowerCase() : "");

  const leads = ss.getSheetByName("Leads");
  if (!leads) {
    SpreadsheetApp.getUi().alert('No se encontró la hoja "Leads"');
    return;
  }

  const leadsData = leads.getRange(2, 1, leads.getLastRow() - 1, 16).getValues();
  const leadsCCMap = new Map();
  const leadsMailMap = new Map();

  leadsData.forEach(row => {
    const cc = row[3] ? String(row[3]).trim() : "";
    const mail = row[8] ? String(row[8]).trim().toLowerCase() : "";
    if (cc) leadsCCMap.set(cc, row);
    if (mail) leadsMailMap.set(mail, row);
  });

  function formatearFecha(fecha) {
    if (!fecha) return "";
    const d = new Date(fecha);
    if (isNaN(d)) return "";
    const yyyy = d.getFullYear();
    const MM = ("0" + (d.getMonth() + 1)).slice(-2);
    const dd = ("0" + d.getDate()).slice(-2);
    const HH = ("0" + d.getHours()).slice(-2);
    const mm = ("0" + d.getMinutes()).slice(-2);
    const ss = ("0" + d.getSeconds()).slice(-2);
    return `${yyyy}-${MM}-${dd} ${HH}:${mm}:${ss}`;
  }

  const resultados = [];

  for (let i = 0; i < cedulas.length; i++) {
    const ccVal = leadsCCMap.has(cedulas[i]) ? 1 : 0;
    const mailVal = leadsMailMap.has(correos[i]) ? 1 : 0;
    const ventas = ccVal + mailVal;

    let fuente = "", medio = "", campaña = "", fecha = "";

    if (ventas > 0) {
      const filaEncontrada = leadsCCMap.get(cedulas[i]) || leadsMailMap.get(correos[i]);
      if (filaEncontrada) {
        fuente = filaEncontrada[10] ? String(filaEncontrada[10]) : "";
        medio = filaEncontrada[11] ? String(filaEncontrada[11]) : "";
        campaña = filaEncontrada[12] ? String(filaEncontrada[12]) : "";
        fecha = formatearFecha(filaEncontrada[15]);
      }
    }

    resultados.push([ccVal, mailVal, ventas, fuente, medio, campaña, "", fecha]);
  }

  hoja.getRange(2, 16, resultados.length, encabezados.length)
       .setValues(resultados)
       .setHorizontalAlignment("right");
}
