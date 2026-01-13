function EmsionesHogarRefactorizada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEmisiones = ss.getSheetByName("Emisiones Enero");
  const hojaLeads = ss.getSheetByName("Leads para cruce");

  if (!hojaEmisiones || !hojaLeads) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja 'Emisiones Enero' o 'Leads para cruce'");
    return;
  }

  const ultimaFilaEmisiones = hojaEmisiones.getLastRow();
  if (ultimaFilaEmisiones < 2) return;

  // 1. CARGAR MAPA DE DATOS (Leads para cruce)
  // Leemos hasta la columna G (7 columnas mínimo)
  const basesData = hojaLeads.getRange(2, 1, hojaLeads.getLastRow() - 1, 10).getValues();
  const basesMap = new Map();

  basesData.forEach(row => {
    // Según tu indicación: CC en F (5) y Correo en G (6)
    const cedula = row[5] != null ? String(row[5]).trim() : ""; 
    const email = row[6] != null ? String(row[6]).trim().toLowerCase() : ""; 
    
    // Guardamos la fila para poder extraer UTMs después
    if (cedula && cedula !== "") basesMap.set(cedula, row);
    if (email && email !== "") basesMap.set(email, row);
  });

  // 2. OBTENER DATOS DE EMISIONES (K=CC, M=Correo)
  const dataEmisiones = hojaEmisiones.getRange(2, 11, ultimaFilaEmisiones - 1, 3).getValues(); 

  // 3. PROCESAR CRUCE
  const resultados = dataEmisiones.map(fila => {
    const ccBusqueda = String(fila[0]).trim();
    const correoBusqueda = String(fila[2]).trim().toLowerCase();

    let cruzadoCC = 0;
    let cruzadoCorreo = 0;
    let fuente = "", med = "", campaña = "", fechaLead = "";

    // Buscar coincidencia en el mapa
    const leadEncontrado = basesMap.get(ccBusqueda) || basesMap.get(correoBusqueda);

    if (leadEncontrado) {
      cruzadoCC = (ccBusqueda !== "" && basesMap.has(ccBusqueda)) ? 1 : 0;
      cruzadoCorreo = (correoBusqueda !== "" && basesMap.has(correoBusqueda)) ? 1 : 0;
      
      // Mapeo según tus nuevas indicaciones:
      fuente = leadEncontrado[1];   // Columna B (UTM_SOURCE)
      med = leadEncontrado[2];      // Columna C (UTM_MEDIUM)
      campaña = leadEncontrado[3];  // Columna D (UTM_CAMPAIGN)
      
      // Si tienes la fecha en alguna columna, ajusta el índice aquí (ejemplo: Columna E = row[4])
      fechaLead = leadEncontrado[4] ? formatearFechaHora(leadEncontrado[4]) : "";
    }

    const suma = cruzadoCC + cruzadoCorreo;

    return [
      cruzadoCC,      // Col O
      cruzadoCorreo,  // Col P
      suma,           // Col Q
      fuente,         // Col R
      med,            // Col S
      campaña,        // Col T
      fechaLead       // Col U
    ];
  });

  // 4. ESCRIBIR RESULTADOS
  const encabezados = ["CC Encontrado", "Email Encontrado", "Suma", "Fuente", "Med", "Campaña", "Fecha Lead"];
  hojaEmisiones.getRange(1, 15, 1, encabezados.length).setValues([encabezados]);
  
  hojaEmisiones.getRange(2, 15, resultados.length, encabezados.length)
               .setValues(resultados)
               .setHorizontalAlignment("right");
}

function formatearFechaHora(fecha) {
  if (!fecha || isNaN(new Date(fecha).getTime())) return "";
  return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}