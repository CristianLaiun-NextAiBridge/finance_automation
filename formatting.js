// ==========================================
// FORMATEO DE LA HOJA LEDGER
//
// Lee los datos raw de la pestaña "mercury" y los transforma al escribirlos
// en la pestaña "Ledger", aplicando:
//
// - Columnas del Ledger: Checked | Date | Description | Category |
//   Comments | Amount In (+) | Amount Out (-) | Account balance | Status | Receipt | JSON Data
// - Description: unificación de Description + Bank Description + Reference + Note,
//   eliminando duplicados y separando con " | "
// - Amount: dividido en Amount In (+) para créditos y Amount Out (-) para débitos
// - Timestamp, Bank Description, Reference, Note y columnas técnicas: excluidas
// - Formato visual: header azul, fila fija, filtro
//
// asignarCategorias() es una función INDEPENDIENTE — no se llama desde aquí.
// ==========================================

// Columnas base del Ledger — orden definitivo
const LEDGER_BASE_HEADERS = [
  'Checked', 'Date', 'Description', 'Category', 'Comments',
  'Amount In (+)', 'Amount Out (-)', 'Account balance'
];

function formatearHoja() {
  Logger.log("🎨 Iniciando formateo del Ledger...");

  const spreadsheet = SpreadsheetApp.openById(CONCILIATION_SHEET_ID);
  const ledger      = spreadsheet.getSheetByName(LEDGER_TAB_NAME);

  if (!ledger) {
    Logger.log('❌ No existe la pestaña "' + LEDGER_TAB_NAME + '".');
    return;
  }

  const mercuryTab = spreadsheet.getSheetByName(MERCURY_TAB_NAME);
  if (!mercuryTab) {
    Logger.log('❌ No existe la pestaña "' + MERCURY_TAB_NAME + '". Ejecutá actualizarTablaMercury() primero.');
    return;
  }

  const srcData    = mercuryTab.getDataRange().getValues();
  const srcHeaders = srcData[0];

  // Índices en la pestaña mercury — usar los nombres originales de esa pestaña
  const SI = {
    id:          srcHeaders.indexOf('id'),
    date:        srcHeaders.indexOf('Date (UTC)'),
    description: srcHeaders.indexOf('Description'),
    bankDesc:    srcHeaders.indexOf('Bank Description'),
    reference:   srcHeaders.indexOf('Reference'),
    note:        srcHeaders.indexOf('Note'),
    amount:      srcHeaders.indexOf('Amount'),
    balance:     srcHeaders.indexOf('Balance')
  };

  // ── 1. INICIALIZAR HEADERS DEL LEDGER SI ESTÁ VACÍO ─────────────────────────
  if (ledger.getLastRow() === 0) {
    ledger.getRange(1, 1, 1, LEDGER_BASE_HEADERS.length).setValues([LEDGER_BASE_HEADERS]);
  }

  // ── 2. LEER ESTADO ACTUAL DEL LEDGER ────────────────────────────────────────
  let ledgerHeaders = ledger.getRange(1, 1, 1, ledger.getLastColumn()).getValues()[0];

  // Asegurar columnas de gestión extra (van al final)
  function ensureColumn(name) {
    let col = ledgerHeaders.indexOf(name) + 1;
    if (col === 0) {
      col = ledger.getLastColumn() + 1;
      ledger.getRange(1, col).setValue(name).setFontWeight('bold');
      ledgerHeaders.push(name);
    }
    return col;
  }

  ensureColumn('Receipt');
  ensureColumn('JSON Data');

  // Dropdown de categorías en Category (si la hoja setup existe)
  const setupTab = spreadsheet.getSheetByName(SETUP_TAB_NAME);
  if (setupTab) {
    const categorias = _leerCategorias(setupTab);
    if (categorias.length > 0) {
      const assignedCol = ledgerHeaders.indexOf('Category') + 1;
      _aplicarDropdownCategorias(ledger, assignedCol, categorias);
    }
  }

  // ── 3. CONSTRUIR SET DE IDs YA EN LEDGER (deduplicación) ────────────────────
  const existingKeys    = new Set();
  const ledgerDateIdx   = ledgerHeaders.indexOf('Date');
  const ledgerAmtInIdx  = ledgerHeaders.indexOf('Amount In (+)');
  const ledgerAmtOutIdx = ledgerHeaders.indexOf('Amount Out (-)');
  const ledgerDescIdx   = ledgerHeaders.indexOf('Description');
  // ledgerHeaders.indexOf('Checked') = 0 → no se usa para deduplicación

  if (ledger.getLastRow() > 1) {
    const ledgerData = ledger.getRange(2, 1, ledger.getLastRow() - 1, ledgerHeaders.length).getValues();
    ledgerData.forEach(function(row) {
      let date = row[ledgerDateIdx];
      if (date instanceof Date) date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const amt  = row[ledgerAmtInIdx] || row[ledgerAmtOutIdx] || 0;
      const desc = (row[ledgerDescIdx] || '').toString().toLowerCase().trim();
      existingKeys.add(date + '|' + Math.abs(parseFloat(amt)).toFixed(2) + '|' + desc);
    });
  }

  // ── 4. TRANSFORMAR Y AGREGAR FILAS NUEVAS ───────────────────────────────────
  const nuevasFilas = [];

  for (let s = 1; s < srcData.length; s++) {
    const row = srcData[s];

    // Construir descripción unificada
    const descParts = [
      (row[SI.description] || '').toString().trim(),
      (row[SI.bankDesc]    || '').toString().trim(),
      (row[SI.reference]   || '').toString().trim(),
      (row[SI.note]        || '').toString().trim()
    ].filter(Boolean);
    const seen = new Set();
    const description = descParts.filter(function(p) {
      const k = p.toLowerCase();
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    }).join(' | ') || 'Sin detalles';

    // Monto dividido
    const rawAmount = parseFloat((row[SI.amount] || 0).toString().replace(/,/g, ''));
    const amountIn  = rawAmount > 0 ? rawAmount        : '';
    const amountOut = rawAmount < 0 ? Math.abs(rawAmount) : '';

    let date = row[SI.date];
    if (date instanceof Date) date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const amt    = amountIn || amountOut || 0;
    const clave  = date + '|' + Math.abs(parseFloat(amt)).toFixed(2) + '|' + description.toLowerCase().trim();

    if (existingKeys.has(clave)) continue;

    // Construir fila del Ledger
    const fila = new Array(ledgerHeaders.length).fill('');
    fila[ledgerHeaders.indexOf('Checked')]           = false;
    fila[ledgerHeaders.indexOf('Date')]        = date;
    fila[ledgerHeaders.indexOf('Description')]       = description;
    fila[ledgerHeaders.indexOf('Category')] = '';
    fila[ledgerHeaders.indexOf('Comments')]          = '';
    fila[ledgerHeaders.indexOf('Amount In (+)')]     = amountIn;
    fila[ledgerHeaders.indexOf('Amount Out (-)')]    = amountOut;
    fila[ledgerHeaders.indexOf('Account balance')]           = row[SI.balance];

    nuevasFilas.push(fila);
    existingKeys.add(clave);
  }

  if (nuevasFilas.length > 0) {
    ledger.getRange(ledger.getLastRow() + 1, 1, nuevasFilas.length, ledgerHeaders.length).setValues(nuevasFilas);
    Logger.log('📥 ' + nuevasFilas.length + ' filas nuevas agregadas al Ledger.');
  } else {
    Logger.log('ℹ️ Sin transacciones nuevas para el Ledger.');
  }

  // ── 5. FORMATO VISUAL ────────────────────────────────────────────────────────
  const lastCol     = ledger.getLastColumn();
  const lastRow     = ledger.getLastRow();
  const headerRange = ledger.getRange(1, 1, 1, lastCol);

  headerRange.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
  ledger.setFrozenRows(1);
  if (!ledger.getFilter()) headerRange.createFilter();

  // Checkboxes en columna Checked (no toca celdas que ya tienen valor)
  const checkedCol = ledgerHeaders.indexOf('Checked') + 1;
  if (checkedCol > 0 && lastRow > 1) {
    const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    ledger.getRange(2, checkedCol, lastRow - 1, 1).setDataValidation(checkboxRule);
  }

  // Conditional formatting: fila completa verde claro cuando Checked = TRUE
  if (checkedCol > 0 && lastRow > 1) {
    const checkedColLetter = _colLetter(checkedCol);
    const dataRange        = ledger.getRange(2, 1, ledger.getMaxRows() - 1, lastCol);

    const greenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + checkedColLetter + '2=TRUE')
      .setBackground('#b7e1cd') // verde claro
      .setRanges([dataRange])
      .build();

    // Reemplazar reglas anteriores de Checked para no duplicar
    const existingRules = ledger.getConditionalFormatRules().filter(function(r) {
      const cond = r.getBooleanCondition();
      return !(cond && cond.getCriteriaValues()[0] &&
               cond.getCriteriaValues()[0].toString().includes(checkedColLetter));
    });
    existingRules.push(greenRule);
    ledger.setConditionalFormatRules(existingRules);
  }

  Logger.log('✅ Formateo del Ledger completado.');
}


// ==========================================
// ASIGNACIÓN DE CATEGORÍAS — FUNCIÓN INDEPENDIENTE
// Ejecutar por separado, no forma parte del formateo.
// ==========================================

function asignarCategorias() {
  const spreadsheet = SpreadsheetApp.openById(CONCILIATION_SHEET_ID);
  const ledger      = spreadsheet.getSheetByName(LEDGER_TAB_NAME);
  const setupTab    = spreadsheet.getSheetByName(SETUP_TAB_NAME);

  if (!ledger) {
    Logger.log('❌ No existe la pestaña "' + LEDGER_TAB_NAME + '".');
    return;
  }
  if (!setupTab) {
    Logger.log('❌ No existe la pestaña "' + SETUP_TAB_NAME + '".');
    return;
  }

  const headers        = ledger.getRange(1, 1, 1, ledger.getLastColumn()).getValues()[0];
  const assignedCatCol = headers.indexOf('Category') + 1;

  if (assignedCatCol === 0) {
    Logger.log("❌ No existe la columna 'Category'. Ejecutá formatearHoja() primero.");
    return;
  }

  const categorias = _leerCategorias(setupTab);
  if (categorias.length === 0) {
    Logger.log('❌ No se encontraron categorías en la pestaña "' + SETUP_TAB_NAME + '".');
    return;
  }

  Logger.log('📋 Categorías: ' + categorias.join(', '));
  _aplicarDropdownCategorias(ledger, assignedCatCol, categorias);

  const sa    = getServiceAccount();
  const token = obtenerTokenDeAcceso(sa);
  _asignarCategoriasIA(ledger, headers, assignedCatCol, categorias, sa, token);
}


// ── HELPERS INTERNOS ─────────────────────────────────────────────────────────

// Convierte número de columna (1-based) a letra(s): 1→A, 26→Z, 27→AA
function _colLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function _leerCategorias(setupTab) {
  try {
    const valores = setupTab.getRange(1, 1, setupTab.getLastRow(), 1).getValues();
    const inicio  = (valores[0] && valores[0][0].toString().toLowerCase().includes('categ')) ? 1 : 0;
    return valores.slice(inicio).map(function(r) { return r[0].toString().trim(); }).filter(Boolean);
  } catch(e) {
    Logger.log('⚠️ Error leyendo categorías: ' + e.toString());
    return [];
  }
}

function _aplicarDropdownCategorias(sheet, col, categorias) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categorias, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);
}

function _asignarCategoriasIA(sheet, headers, assignedCategoryCol, categorias, sa, token) {
  const lastRow   = sheet.getLastRow();
  if (lastRow < 2) return;

  const descIdx   = headers.indexOf('Description');
  const amtInIdx  = headers.indexOf('Amount In (+)');
  const amtOutIdx = headers.indexOf('Amount Out (-)');
  const dateIdx   = headers.indexOf('Date');

  const totalCols  = sheet.getLastColumn();
  const dataRango  = sheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  const catRango   = sheet.getRange(2, assignedCategoryCol, lastRow - 1, 1).getValues();
  const listaTexto = categorias.join('\n- ');

  const peticiones = [];
  const mapeo      = [];

  for (let i = 0; i < dataRango.length; i++) {
    if (catRango[i][0] && catRango[i][0].toString().trim() !== '') continue;

    const desc   = (dataRango[i][descIdx]   || '').toString();
    const amount = (dataRango[i][amtInIdx]  || dataRango[i][amtOutIdx] || '').toString();
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
    mapeo.push(i + 2);
  }

  if (peticiones.length === 0) {
    Logger.log('ℹ️ Todas las filas ya tienen categoría asignada.');
    return;
  }

  Logger.log('🏷️ Asignando categorías para ' + peticiones.length + ' filas...');
  const respuestas = ejecutarBatchIA(peticiones, token, sa);

  for (let k = 0; k < respuestas.length; k++) {
    if (!respuestas[k] || respuestas[k].getResponseCode() !== 200) continue;
    try {
      const texto = JSON.parse(respuestas[k].getContentText())
                      .candidates[0].content.parts[0].text.trim();
      const match = categorias.find(function(c) { return c.toLowerCase() === texto.toLowerCase(); });
      sheet.getRange(mapeo[k], assignedCategoryCol).setValue(match || texto);
      Logger.log('🏷️ Fila ' + mapeo[k] + ' → ' + (match || texto));
    } catch(e) {
      Logger.log('⚠️ Error fila ' + mapeo[k] + ': ' + e.toString());
    }
  }

  Logger.log('✅ Categorías asignadas.');
}
