// ==========================================
// FORMATEO Y PREPARACIÓN DE LA HOJA DE CONCILIACIÓN
//
// Se encarga de tres cosas en orden:
//
// 1. SINCRONIZACIÓN DE FILAS: copia las transacciones nuevas del sheet raw de
//    Mercury (MERCURY_SOURCE_SHEET_ID) a la hoja activa, omitiendo las que ya
//    existen (detección por id o por fecha+monto+descripción).
//
// 2. GESTIÓN DE COLUMNAS:
//    - "comprobante" y "JSON Data": usadas por invoice_matching.js
//    - "assigned_category": la IA asigna automáticamente una categoría tomando
//      como referencia la lista de la hoja "asigned_category" del mismo archivo.
//      Solo se procesa si la celda está vacía.
//    - "comments": columna libre para anotaciones manuales de los usuarios.
//      Nunca se sobreescribe si ya tiene contenido.
//
// 3. FORMATO VISUAL: header azul, fila fija, filtro, y ocultamiento de columnas
//    técnicas para mantener la vista limpia.
// ==========================================

function formatearHoja() {
  Logger.log("🎨 Iniciando formateo y sincronización de la hoja...");

  const spreadsheet = SpreadsheetApp.openById(CONCILIATION_SHEET_ID);
  const targetSheet = spreadsheet.getSheets()[0];

  // ── 1. SINCRONIZAR FILAS NUEVAS DESDE MERCURY ───────────────────────────────
  let currentData = targetSheet.getLastRow() > 0 ? targetSheet.getDataRange().getValues() : [];
  let headers     = currentData.length > 0 ? currentData[0] : [];

  const existingTxIds   = new Set();
  const currentIdIdx    = headers.indexOf('id');
  let   currentDateIdx  = headers.indexOf('Timestamp');
  if (currentDateIdx === -1) currentDateIdx = headers.indexOf('Date (UTC)');
  const currentAmountIdx = headers.indexOf('Amount');
  const currentDescIdx   = headers.indexOf('Description');

  for (let r = 1; r < currentData.length; r++) {
    if (currentIdIdx !== -1 && currentData[r][currentIdIdx]) {
      existingTxIds.add(currentData[r][currentIdIdx].toString().trim());
    } else if (currentDateIdx !== -1 && currentAmountIdx !== -1) {
      let dateVal = currentData[r][currentDateIdx];
      if (dateVal instanceof Date) dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const amountVal = currentData[r][currentAmountIdx] ? Math.abs(parseFloat(currentData[r][currentAmountIdx])).toFixed(2) : "0.00";
      const descVal   = (currentData[r][currentDescIdx] || '').toString().toLowerCase().trim();
      existingTxIds.add(dateVal + '_' + amountVal + '_' + descVal);
    }
  }

  let sourceData = [];
  try {
    sourceData = SpreadsheetApp.openById(MERCURY_SOURCE_SHEET_ID).getSheets()[0].getDataRange().getValues();
  } catch(e) {
    Logger.log("❌ Error leyendo el sheet fuente de Mercury: " + e.toString());
    return;
  }

  const sourceHeaders  = sourceData[0];
  const sourceIdIdx    = sourceHeaders.indexOf('id');
  const amountIdx      = sourceHeaders.indexOf('Amount');
  let   dateIdx        = sourceHeaders.indexOf('Timestamp');
  if (dateIdx === -1) dateIdx = sourceHeaders.indexOf('Date (UTC)');
  const descIdx        = sourceHeaders.indexOf('Description');

  if (sourceIdIdx === -1 || amountIdx === -1 || dateIdx === -1) {
    Logger.log("❌ El sheet fuente no tiene el formato esperado.");
    return;
  }

  const indicesAOmitir = COLUMNAS_A_SACAR.map(col => sourceHeaders.indexOf(col)).filter(idx => idx !== -1);

  if (headers.length === 0) {
    headers = sourceHeaders.filter(function(_, idx) { return !indicesAOmitir.includes(idx); });
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const nuevasFilas = [];
  for (let s = 1; s < sourceData.length; s++) {
    const row    = sourceData[s];
    const txId   = row[sourceIdIdx];
    let   txDate = row[dateIdx];
    if (txDate instanceof Date) txDate = Utilities.formatDate(txDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const txAmount = row[amountIdx] ? Math.abs(parseFloat(row[amountIdx])).toFixed(2) : "0.00";
    const txDesc   = (row[descIdx] || '').toString().toLowerCase().trim();
    const compKey  = txDate + '_' + txAmount + '_' + txDesc;

    if ((txId && existingTxIds.has(txId.toString().trim())) || existingTxIds.has(compKey)) continue;

    const nuevaFila = new Array(headers.length).fill("");
    let t = 0;
    for (let c = 0; c < row.length; c++) {
      if (!indicesAOmitir.includes(c)) { nuevaFila[t] = row[c]; t++; }
    }
    nuevasFilas.push(nuevaFila);
  }

  if (nuevasFilas.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
    Logger.log("📥 " + nuevasFilas.length + " filas nuevas agregadas.");
  } else {
    Logger.log("ℹ️ Sin transacciones nuevas.");
  }

  // ── 2. GESTIÓN DE COLUMNAS ───────────────────────────────────────────────────
  // Re-leer headers después de posibles cambios
  headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];

  function ensureColumn(name) {
    let col = headers.indexOf(name) + 1;
    if (col === 0) {
      col = targetSheet.getLastColumn() + 1;
      targetSheet.getRange(1, col).setValue(name).setFontWeight("bold");
      headers.push(name);
    }
    return col;
  }

  const comprobanteCol      = ensureColumn('comprobante');
  const jsonDataCol         = ensureColumn('JSON Data');
  const assignedCategoryCol = ensureColumn('assigned_category');
  const commentsCol         = ensureColumn('comments');

  // ── 3. ASIGNAR CATEGORÍAS CON IA ─────────────────────────────────────────────
  // Lee la lista de categorías desde la hoja "asigned_category" del mismo archivo.
  // Solo procesa filas con la celda assigned_category vacía.
  const categorias = _leerCategorias(spreadsheet);

  if (categorias.length > 0) {
    Logger.log("📋 Categorías encontradas: " + categorias.join(', '));
    const sa    = getServiceAccount();
    const token = obtenerTokenDeAcceso(sa);
    _asignarCategoriasIA(targetSheet, headers, assignedCategoryCol, categorias, sa, token);
  } else {
    Logger.log("⚠️ No se encontraron categorías en la hoja 'asigned_category'. Se omite asignación.");
  }

  // ── 4. FORMATO VISUAL ────────────────────────────────────────────────────────
  const lastCol     = targetSheet.getLastColumn();
  const headerRange = targetSheet.getRange(1, 1, 1, lastCol);

  headerRange.setBackground("#4a86e8").setFontColor("#ffffff").setFontWeight("bold");
  targetSheet.setFrozenRows(1);
  if (!targetSheet.getFilter()) headerRange.createFilter();

  Logger.log("✅ Formateo completado.");
}

// ── FUNCIÓN STANDALONE: se puede ejecutar sola o es llamada por formatearHoja() ──

// Asigna la categoría correcta a cada transacción usando IA.
// Lee la lista de categorías desde la hoja "asigned_category" del spreadsheet
// de conciliación. Solo procesa filas con la celda assigned_category vacía.
function asignarCategorias() {
  const spreadsheet     = SpreadsheetApp.openById(CONCILIATION_SHEET_ID);
  const targetSheet     = spreadsheet.getSheets()[0];
  const headers         = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const assignedCatCol  = headers.indexOf('assigned_category') + 1;

  if (assignedCatCol === 0) {
    Logger.log("❌ No existe la columna 'assigned_category'. Ejecutá formatearHoja() primero.");
    return;
  }

  const categorias = _leerCategorias(spreadsheet);
  if (categorias.length === 0) {
    Logger.log("❌ No se encontraron categorías. Verificá que exista la hoja 'asigned_category' con categorías en la columna A.");
    return;
  }

  Logger.log("📋 Categorías encontradas: " + categorias.join(', '));

  const sa    = getServiceAccount();
  const token = obtenerTokenDeAcceso(sa);
  _asignarCategoriasIA(targetSheet, headers, assignedCatCol, categorias, sa, token);
}

// ── HELPERS INTERNOS ─────────────────────────────────────────────────────────

function _leerCategorias(spreadsheet) {
  // Busca la hoja con nombre exacto o insensible a mayúsculas
  const sheets   = spreadsheet.getSheets();
  const catSheet = sheets.find(function(s) {
    return s.getName().toLowerCase().replace(/\s/g,'') === 'asigned_category';
  });

  if (!catSheet) {
    Logger.log("⚠️ Hojas disponibles: " + sheets.map(function(s) { return s.getName(); }).join(', '));
    return [];
  }

  const valores = catSheet.getRange(1, 1, catSheet.getLastRow(), 1).getValues();
  // Saltamos la primera fila si parece un header
  const inicio  = (valores[0] && valores[0][0].toString().toLowerCase().includes('categ')) ? 1 : 0;
  return valores.slice(inicio).map(function(r) { return r[0].toString().trim(); }).filter(Boolean);
}

function _asignarCategoriasIA(sheet, headers, assignedCategoryCol, categorias, sa, token) {
  const lastRow   = sheet.getLastRow();
  if (lastRow < 2) return;

  const descIdx   = headers.indexOf('Description');
  const amountIdx = headers.indexOf('Amount');
  let   dateIdx   = headers.indexOf('Timestamp');
  if (dateIdx === -1) dateIdx = headers.indexOf('Date (UTC)');

  const totalCols  = sheet.getLastColumn();
  const dataRango  = sheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  const catRango   = sheet.getRange(2, assignedCategoryCol, lastRow - 1, 1).getValues();
  const listaTexto = categorias.join('\n- ');

  const peticiones = [];
  const mapeo      = [];

  for (let i = 0; i < dataRango.length; i++) {
    // No sobreescribir si ya tiene categoría asignada
    if (catRango[i][0] && catRango[i][0].toString().trim() !== '') continue;

    const desc   = (dataRango[i][descIdx]   || '').toString();
    const amount = (dataRango[i][amountIdx] || '').toString();
    const date   = (dataRango[i][dateIdx]   || '').toString();

    if (!desc && !amount) continue;

    peticiones.push({
      prompt: 'Eres un contador que clasifica transacciones bancarias corporativas.\n\n' +
              'Transacción:\n' +
              '- Fecha: '       + date   + '\n' +
              '- Monto: $'      + amount + '\n' +
              '- Descripción: ' + desc   + '\n\n' +
              'Categorías disponibles:\n- ' + listaTexto + '\n\n' +
              'Asigná la categoría más apropiada.\n' +
              'Respondé ÚNICAMENTE con el nombre exacto de la categoría, sin explicación ni puntuación.'
    });
    mapeo.push(i + 2); // fila real en el sheet (1-indexed + header)
  }

  if (peticiones.length === 0) {
    Logger.log("ℹ️ Todas las filas ya tienen categoría asignada.");
    return;
  }

  Logger.log("🏷️ Asignando categorías para " + peticiones.length + " filas...");
  const respuestas = ejecutarBatchIA(peticiones, token, sa);

  for (let k = 0; k < respuestas.length; k++) {
    if (!respuestas[k] || respuestas[k].getResponseCode() !== 200) continue;
    try {
      const texto = JSON.parse(respuestas[k].getContentText())
                      .candidates[0].content.parts[0].text.trim();
      // Preferir coincidencia exacta (case-insensitive); si no hay, escribir igual lo que devolvió la IA
      const match = categorias.find(function(c) { return c.toLowerCase() === texto.toLowerCase(); });
      sheet.getRange(mapeo[k], assignedCategoryCol).setValue(match || texto);
      Logger.log("🏷️ Fila " + mapeo[k] + " → " + (match || texto));
    } catch(e) {
      Logger.log("⚠️ Error procesando categoría fila " + mapeo[k] + ": " + e.toString());
    }
  }

  Logger.log("✅ Categorías asignadas.");
}
