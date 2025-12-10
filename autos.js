function EmisionesAutosCruzados(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!nombreHoja) nombreHoja = "Emisiones 24 nov"; 

  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    ui.alert(`Error: No se encontró la hoja '${nombreHoja}'.`);
    return;
  }
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const extraerValorDeJson = (valor) => {
    let texto = String(valor || '');
    if (texto.trim().startsWith('{') && texto.includes(':')) {
      try {
        const obj = JSON.parse(texto);
        return Object.values(obj).join(''); 
      } catch (e) {
        return texto;
      }
    }
    return texto;
  };

  const limpiarDoc = d => {
    const textoReal = extraerValorDeJson(d);
    return textoReal.replace(/[^a-z0-9]/gi, '').toLowerCase();
  };

  const limpiarPlaca = p => {
    const textoReal = extraerValorDeJson(p);
    return textoReal.replace(/[^a-z0-9]/gi, '').toLowerCase();
  };

  const limpiarCorreo = c => String(c || '').trim().toLowerCase();  


  const datosEmisiones = hoja.getRange(2, 1, ultimaFila - 1, 20).getValues();


  const hojaLeads = ss.getSheetByName("TOTAL LEADS - CRUZAR");
  if (!hojaLeads) { ui.alert("Falta hoja TOTAL LEADS - CRUZAR"); return; }
  
  const datosLeads = hojaLeads.getDataRange().getValues().slice(1); 
  
  const mapTLC_CC = new Map(), mapTLC_Placa = new Map(), mapTLC_Correo = new Map();
  
  datosLeads.forEach(r => {
    const cedula = r[1] ? limpiarDoc(r[1]) : ''; 
    const correo = r[2] ? limpiarCorreo(r[2]) : ''; 
    const placa = r[4] ? limpiarPlaca(r[4]) : '';  

    if (cedula && !mapTLC_CC.has(cedula)) mapTLC_CC.set(cedula, r);
    if (correo && !mapTLC_Correo.has(correo)) mapTLC_Correo.set(correo, r);
    if (placa && !mapTLC_Placa.has(placa)) mapTLC_Placa.set(placa, r);
  });


  const hojaBases = ss.getSheetByName("BASES");
  if (!hojaBases) { ui.alert("Falta hoja BASES"); return; }
  const basesDatos = hojaBases.getDataRange().getValues().slice(1);
  
  basesDatos.sort((a, b) => {
    const fA = new Date(a[2]); const fB = new Date(b[2]);
    return (isNaN(fB) ? -1 : (isNaN(fA) ? 1 : fB - fA));
  });

  const mapBASES_CC = new Map(), mapBASES_Correo = new Map();
  basesDatos.forEach(r => {
    const cedula = r[0] ? limpiarDoc(r[0]) : '';
    const correo = r[1] ? limpiarCorreo(r[1]) : '';
    if (cedula && !mapBASES_CC.has(cedula)) mapBASES_CC.set(cedula, r);
    if (correo && !mapBASES_Correo.has(correo)) mapBASES_Correo.set(correo, r);
  });
  
  function formatearFecha(fecha) {
    if (!fecha) return '';
    const d = new Date(fecha);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, ss.getSpreadsheetTimeZone() || 'GMT-5', "yyyy-MM-dd HH:mm:ss");
  }


  console.log("--- INICIANDO CRUCE ---");
  
  const resultadosFinales = datosEmisiones.map((r, i) => {
    
    const placa = limpiarPlaca(r[2]); 

    const doc = limpiarDoc(r[12]);     
    const correo = limpiarCorreo(r[14]); 

    if (i === 0) {
        console.log(`FILA 1 | Placa Original: '${r[0]}' -> Limpia: '${placa}'`);
        console.log(`FILA 1 | Doc Original: '${r[11]}' -> Limpio: '${doc}'`);
    }

    let foundRowTLC = null;
    let foundRowBASES = null;

    if (doc && mapTLC_CC.has(doc)) foundRowTLC = mapTLC_CC.get(doc);
    if (placa && !foundRowTLC && mapTLC_Placa.has(placa)) foundRowTLC = mapTLC_Placa.get(placa);
    if (correo && !foundRowTLC && mapTLC_Correo.has(correo)) foundRowTLC = mapTLC_Correo.get(correo);
    
    if (!foundRowTLC) {
      if (doc && mapBASES_CC.has(doc)) foundRowBASES = mapBASES_CC.get(doc);
      if (correo && !foundRowBASES && mapBASES_Correo.has(correo)) foundRowBASES = mapBASES_Correo.get(correo);
    }
    
    let countTLC_CC = (doc && mapTLC_CC.has(doc)) ? 1 : 0;
    let countTLC_Placa = (placa && mapTLC_Placa.has(placa)) ? 1 : 0;
    let countTLC_Correo = (correo && mapTLC_Correo.has(correo)) ? 1 : 0;
    let countBASES_CC = (doc && mapBASES_CC.has(doc)) ? 1 : 0;
    let countBASES_Correo = (correo && mapBASES_Correo.has(correo)) ? 1 : 0;

    let fuenteFinal = '', medioFinal = '', campañaFinal = '', adnameFinal = '', fechaFinal = '';

    if (foundRowTLC) {
      fuenteFinal = foundRowTLC[7] || '';   
      medioFinal = foundRowTLC[11] || '';   
      campañaFinal = foundRowTLC[8] || '';  
      adnameFinal = foundRowTLC[13] || '';  
      fechaFinal = formatearFecha(foundRowTLC[10]); 
    } else if (foundRowBASES) {
      fuenteFinal = foundRowBASES[7] || 'BASES';
      medioFinal = foundRowBASES[3] || '';
      campañaFinal = foundRowBASES[6] || '';
      fechaFinal = formatearFecha(foundRowBASES[2]);
    }

    const conteos = [countTLC_CC, countTLC_Placa, countTLC_Correo, countBASES_CC, countBASES_Correo];
    const totalConteo = conteos.reduce((a, b) => a + b, 0);

    return [...conteos, totalConteo, totalConteo, fuenteFinal, medioFinal, campañaFinal, adnameFinal, fechaFinal];
  });

  const nuevosEncabezados = ["TOTAL LEADS CC", "TOTAL LEADS Placa", "TOTAL LEADS Correo", "Bases CC", "Bases Correo", "Total Leads", "Ventas", "Fuente", "Medio", "Campaña", "Adname", "Fecha Lead"];

  if (resultadosFinales.length > 0) {

    hoja.getRange(1, 17, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    hoja.getRange(2, 17, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);
  }
  
  SpreadsheetApp.getUi().alert("Proceso completado correctamente :)");
}