function separarDatosFlexible() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName("treble (3)");

  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja llamada 'Hoja 6'");
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return; 

  hoja.getRange(1, 21, 1, 5).setValues([["Nombre", "Tipo Doc", "Número Doc", "Placa", "Correo"]]);

  const rango = hoja.getRange("B2:B" + ultimaFila);
  const valores = rango.getValues();

  const salida = valores.map((fila) => {
    let texto = "";
    if (fila[0] !== null && fila[0] !== undefined) {
      texto = String(fila[0]).trim();
    }

    if (!texto) return ["", "", "", "", ""];

    const correoMatch = texto.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    const correo = correoMatch ? correoMatch[0] : "";

    const docMatch = texto.match(/(CC|CE|TI|NIT|PAS)?\s*(\d{6,12})/i);
    const tipoDoc = docMatch && docMatch[1] ? docMatch[1].toUpperCase() : "CC";
    const numDoc = docMatch && docMatch[2] ? docMatch[2] : "";

    const placaMatch = texto.match(/\b([A-Z]{3}\d{2,3}[A-Z]?)\b/i);
    const placa = placaMatch ? placaMatch[1].toUpperCase() : "";

    let nombre = texto;

    if (correo) nombre = nombre.replace(correo, "");
    if (docMatch) nombre = nombre.replace(docMatch[0], "");
    if (placa) nombre = nombre.replace(placa, "");

    nombre = nombre.split("//")[0];
    nombre = nombre.split("-")[0];
    
    nombre = nombre.replace(/\b(PROPIETARIO|PROPIETARIA|CUÑADO|MODELO|VALOR|SOLICITUD)\b/gi, "");
    
    nombre = nombre.replace(/[0-9]+/g, ""); 
    nombre = nombre.replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]/g, ""); 
    
    nombre = nombre.replace(/\s+/g, " ").trim();

    return [nombre, tipoDoc, numDoc, placa, correo];
  });

  hoja.getRange(2, 21, salida.length, 5).setValues(salida);
}function separarDatosFlexible() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName("treble (3)");

  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja llamada 'Hoja 6'");
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return; 

  hoja.getRange(1, 21, 1, 5).setValues([["Nombre", "Tipo Doc", "Número Doc", "Placa", "Correo"]]);

  const rango = hoja.getRange("B2:B" + ultimaFila);
  const valores = rango.getValues();

  const salida = valores.map((fila) => {
    let texto = "";
    if (fila[0] !== null && fila[0] !== undefined) {
      texto = String(fila[0]).trim();
    }

    if (!texto) return ["", "", "", "", ""];

    const correoMatch = texto.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    const correo = correoMatch ? correoMatch[0] : "";

    const docMatch = texto.match(/(CC|CE|TI|NIT|PAS)?\s*(\d{6,12})/i);
    const tipoDoc = docMatch && docMatch[1] ? docMatch[1].toUpperCase() : "CC";
    const numDoc = docMatch && docMatch[2] ? docMatch[2] : "";

    const placaMatch = texto.match(/\b([A-Z]{3}\d{2,3}[A-Z]?)\b/i);
    const placa = placaMatch ? placaMatch[1].toUpperCase() : "";

    let nombre = texto;

    if (correo) nombre = nombre.replace(correo, "");
    if (docMatch) nombre = nombre.replace(docMatch[0], "");
    if (placa) nombre = nombre.replace(placa, "");

    nombre = nombre.split("//")[0];
    nombre = nombre.split("-")[0];
    
    nombre = nombre.replace(/\b(PROPIETARIO|PROPIETARIA|CUÑADO|MODELO|VALOR|SOLICITUD)\b/gi, "");
    
    nombre = nombre.replace(/[0-9]+/g, ""); 
    nombre = nombre.replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]/g, ""); 
    
    nombre = nombre.replace(/\s+/g, " ").trim();

    return [nombre, tipoDoc, numDoc, placa, correo];
  });

  hoja.getRange(2, 21, salida.length, 5).setValues(salida);
}function separarDatosFlexible() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName("treble (3)");

  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja llamada 'Hoja 6'");
    return;
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return; 

  hoja.getRange(1, 21, 1, 5).setValues([["Nombre", "Tipo Doc", "Número Doc", "Placa", "Correo"]]);

  const rango = hoja.getRange("B2:B" + ultimaFila);
  const valores = rango.getValues();

  const salida = valores.map((fila) => {
    let texto = "";
    if (fila[0] !== null && fila[0] !== undefined) {
      texto = String(fila[0]).trim();
    }

    if (!texto) return ["", "", "", "", ""];

    const correoMatch = texto.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    const correo = correoMatch ? correoMatch[0] : "";

    const docMatch = texto.match(/(CC|CE|TI|NIT|PAS)?\s*(\d{6,12})/i);
    const tipoDoc = docMatch && docMatch[1] ? docMatch[1].toUpperCase() : "CC";
    const numDoc = docMatch && docMatch[2] ? docMatch[2] : "";

    const placaMatch = texto.match(/\b([A-Z]{3}\d{2,3}[A-Z]?)\b/i);
    const placa = placaMatch ? placaMatch[1].toUpperCase() : "";

    let nombre = texto;

    if (correo) nombre = nombre.replace(correo, "");
    if (docMatch) nombre = nombre.replace(docMatch[0], "");
    if (placa) nombre = nombre.replace(placa, "");

    nombre = nombre.split("//")[0];
    nombre = nombre.split("-")[0];
    
    nombre = nombre.replace(/\b(PROPIETARIO|PROPIETARIA|CUÑADO|MODELO|VALOR|SOLICITUD)\b/gi, "");
    
    nombre = nombre.replace(/[0-9]+/g, ""); 
    nombre = nombre.replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]/g, ""); 
    
    nombre = nombre.replace(/\s+/g, " ").trim();

    return [nombre, tipoDoc, numDoc, placa, correo];
  });

  hoja.getRange(2, 21, salida.length, 5).setValues(salida);
}