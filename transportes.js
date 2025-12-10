function CruzarVentasLeadsTransportes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaVentas = ss.getSheetByName("Ventas 8 dic"); 
  const hojaLeads = ss.getSheetByName("Leads");          
  
  if (!hojaVentas) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja 'Ventas 28 sep'.");
    return;
  }
  if (!hojaLeads) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la hoja 'Leads'.");
    return;
  }

  const ultimaFilaVentas = hojaVentas.getLastRow();
  if (ultimaFilaVentas < 2) return;

  const limpiarDoc = d => String(d || '').replace(/\./g, '').replace(/\s/g, '').toLowerCase();
  const limpiarCorreo = c => String(c || '').trim().toLowerCase();
  
  function formatearFecha(fecha) {
    if (!fecha) return '';
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return '';
    const timeZone = ss.getSpreadsheetTimeZone() || 'GMT-5';
    return Utilities.formatDate(d, timeZone, "yyyy-MM-dd HH:mm:ss");
  }

  const leadsDatos = hojaLeads.getDataRange().getValues().slice(1);
  const mapLeads_CC = new Map(), mapLeads_Correo = new Map();

  leadsDatos.forEach(r => {
    const cedula = r[3] ? limpiarDoc(r[3]) : ''; 
    const correo = r[0] ? limpiarCorreo(r[0]) : '';
    
    if (cedula && !mapLeads_CC.has(cedula)) mapLeads_CC.set(cedula, r);
    if (correo && !mapLeads_Correo.has(correo)) mapLeads_Correo.set(correo, r);
  });

  const DATASIZE = 11; 
  const datosVentas = hojaVentas.getRange(2, 3, ultimaFilaVentas - 1, DATASIZE).getValues();

  const resultadosFinales = datosVentas.map(r => {

    const doc = limpiarDoc(r[8]);      
    const correo = limpiarCorreo(r[10]);
    
    let countCC = 0, countMail = 0;
    let foundRowLead = null;

    if (doc && mapLeads_CC.has(doc)) {
      foundRowLead = mapLeads_CC.get(doc);
      countCC = 1;
    } else if (correo && mapLeads_Correo.has(correo)) {
      foundRowLead = mapLeads_Correo.get(correo);
      countMail = 1;
    }

    let fuente = '', medio = '', campana = '', fechaLead = '';

    if (foundRowLead) {

      fuente = foundRowLead[12] || '';
      medio = foundRowLead[14] || '';
      campana = foundRowLead[15] || '';      
      fechaLead = formatearFecha(foundRowLead[11]); 
    }
    
    const ventas = (countCC || countMail) ? 1 : 0; 
    
    return [
      countCC,       
      countMail,    
      ventas,        
      fuente,        
      medio,        
      campana,       
      fechaLead,     
      '',            
      '',            
      ''             
    ];
  });

  const nuevosEncabezados = [
    "CC", "Mail1", "Ventas",
    "Fuente", "Medio", "Campaña", "Fecha Lead",
    "Póliza", "Ciudad", "Actividad económica"
  ];

  if (resultadosFinales.length > 0) {
    const colInicio = 16; 
    hojaVentas.getRange(1, colInicio, 1, nuevosEncabezados.length).setValues([nuevosEncabezados]);
    hojaVentas.getRange(2, colInicio, resultadosFinales.length, resultadosFinales[0].length).setValues(resultadosFinales);
  }

  
}