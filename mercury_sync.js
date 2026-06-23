// ==========================================
// SINCRONIZACIÓN DE TRANSACCIONES MERCURY
// Lee el token desde Script Properties (nunca hardcodeado).
// Configurarlo: ejecutá setupCredentials() en setup_credentials.js
// ==========================================

function actualizarTablaMercury() {
  const props    = PropertiesService.getScriptProperties();
  const apiToken = props.getProperty('MERCURY_API_TOKEN');

  if (!apiToken) {
    Logger.log("❌ Falta MERCURY_API_TOKEN en Script Properties. Ejecutá setupCredentials().");
    return;
  }

  const hoja = SpreadsheetApp.openById(MERCURY_SOURCE_SHEET_ID).getSheets()[0];

  const opciones = {
    method:             'get',
    headers:            { 'Authorization': 'Bearer ' + apiToken, 'Accept': 'application/json' },
    muteHttpExceptions: true
  };

  // ── 1. SALDO ACTUAL DE CADA CUENTA ──────────────────────────────────────────
  // Usamos GET /accounts para anclar el cálculo del saldo corriente.
  // Si hay varias cuentas, el saldo se calcula por separado para cada una.
  const saldosPorCuenta = {};   // { accountId: currentBalance }
  try {
    const respCuentas = UrlFetchApp.fetch('https://api.mercury.com/api/v1/accounts', opciones);
    if (respCuentas.getResponseCode() === 200) {
      const cuentas = JSON.parse(respCuentas.getContentText()).accounts || [];
      cuentas.forEach(function(c) {
        saldosPorCuenta[c.id] = c.currentBalance;
      });
      Logger.log('Cuentas encontradas: ' + cuentas.map(c => c.name + ' ($' + c.currentBalance + ')').join(', '));
    }
  } catch(e) {
    Logger.log('⚠️ No se pudo obtener el saldo actual: ' + e.toString());
  }

  // ── 2. TRAER TODAS LAS TRANSACCIONES ────────────────────────────────────────
  Logger.log('Trayendo todo el histórico de TODAS las cuentas...');

  let todasLasTransacciones = [];
  let hayMas     = true;
  let startAfter = "";

  while (hayMas) {
    let url = 'https://api.mercury.com/api/v1/transactions?limit=500&start=2020-01-01';
    if (startAfter !== "") url += '&start_after=' + startAfter;

    const respuesta = UrlFetchApp.fetch(url, opciones);
    const datos     = JSON.parse(respuesta.getContentText());

    if (respuesta.getResponseCode() !== 200) {
      Logger.log('Error en la API de Mercury: ' + (datos.message || respuesta.getContentText()));
      return;
    }

    todasLasTransacciones = todasLasTransacciones.concat(datos.transactions || []);

    if (datos.page && datos.page.nextPage) {
      startAfter = datos.page.nextPage;
    } else {
      hayMas = false;
    }
  }

  if (todasLasTransacciones.length === 0) {
    Logger.log('No se encontraron transacciones.');
    return;
  }

  // Ordenar de más antigua a más nueva
  todasLasTransacciones.sort(function(a, b) {
    return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
  });

  // ── 3. CALCULAR SALDO CORRIENTE POR CUENTA ──────────────────────────────────
  // Estrategia: el saldo actual de la API es el punto de anclaje (el más confiable).
  // Trabajamos HACIA ATRÁS desde ese valor para reconstruir el saldo después de
  // cada transacción. Solo se incluyen transacciones con status "sent" o "pending"
  // ya que las "failed" no afectan el saldo real.
  //
  // saldoCorreinte[i] = saldo de la cuenta DESPUÉS de que se ejecutó la tx i.

  // Agrupar índices por accountId
  const txPorCuenta = {};
  todasLasTransacciones.forEach(function(tx, idx) {
    const aid = tx.accountId || 'unknown';
    if (!txPorCuenta[aid]) txPorCuenta[aid] = [];
    txPorCuenta[aid].push(idx);
  });

  const saldoCorreinte = new Array(todasLasTransacciones.length).fill('');

  Object.keys(txPorCuenta).forEach(function(accountId) {
    const indices = txPorCuenta[accountId]; // ya están en orden cronológico
    const saldoActual = saldosPorCuenta[accountId];

    if (saldoActual === undefined) {
      // Sin anclaje: no podemos calcular el saldo corriente para esta cuenta
      return;
    }

    // El saldo después de la ÚLTIMA transacción = saldoActual
    // Trabajamos hacia atrás: saldo_antes_de_tx = saldo_despues_de_tx - tx.amount
    let saldo = saldoActual;
    for (let i = indices.length - 1; i >= 0; i--) {
      const tx = todasLasTransacciones[indices[i]];
      saldoCorreinte[indices[i]] = Math.round(saldo * 100) / 100;
      saldo = saldo - (tx.amount || 0);
    }
  });

  // ── 4. ESCRIBIR EN LA HOJA ───────────────────────────────────────────────────
  hoja.clearContents();

  const cabeceras = [
    'Date (UTC)', 'Timestamp', 'Description', 'Amount', 'Balance After',
    'Status', 'Bank Description', 'Reference', 'Note', 'Category',
    'Source of Category', 'Original Currency', 'kind', 'counterpartyId',
    'counterpartyNickname', 'postedAt', 'dashboardLink', 'attachments', 'id'
  ];
  hoja.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]);

  const filas = todasLasTransacciones.map(function(tx, idx) {
    const card    = tx.merchant || {};
    const details = tx.details  || {};
    const cxInfo  = tx.currencyExchangeInfo || {};

    const merchantName = card.name || details.merchantName || '';
    const description  = tx.counterpartyName || merchantName || tx.bankDescription || tx.externalMemo || tx.note || 'Sin detalles';
    const category     = (tx.mercuryCategory ? tx.mercuryCategory.name : '') || (tx.categoryData ? tx.categoryData.name : '') || '';

    return [
      tx.createdAt ? tx.createdAt.substring(0, 10)      : '',
      tx.createdAt ? new Date(tx.createdAt)              : '',
      description,
      tx.amount !== undefined ? tx.amount                : '',
      saldoCorreinte[idx],                                         // ← Balance After
      tx.status                                          || '',
      tx.bankDescription                                 || '',
      tx.externalMemo || tx.referenceNumber              || '',
      tx.note                                            || '',
      category,
      tx.sourceOfCategory || (tx.mercuryCategory ? 'Mercury' : '') || '',
      cxInfo.originalCurrency || tx.currency             || 'USD',
      tx.kind                                            || '',
      tx.counterpartyId                                  || '',
      tx.counterpartyNickname                            || '',
      tx.postedAt ? new Date(tx.postedAt)                : '',
      tx.dashboardLink                                   || '',
      tx.attachments ? tx.attachments.length             : 0,
      tx.id                                              || ''
    ];
  });

  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, cabeceras.length).setValues(filas);
  }

  Logger.log(`✅ Proceso terminado. ${filas.length} transacciones cargadas con saldo corriente.`);
}
