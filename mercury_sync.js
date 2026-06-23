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
    method:           'get',
    headers:          { 'Authorization': 'Bearer ' + apiToken, 'Accept': 'application/json' },
    muteHttpExceptions: true
  };

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

  hoja.clearContents();

  const cabeceras = [
    'Date (UTC)', 'Timestamp', 'Description', 'Amount', 'Status',
    'Bank Description', 'Reference', 'Note', 'Category',
    'Source of Category', 'Original Currency', 'kind', 'counterpartyId',
    'counterpartyNickname', 'postedAt', 'dashboardLink', 'attachments', 'id'
  ];
  hoja.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]);

  todasLasTransacciones.sort(function(a, b) {
    return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
  });

  const filas = todasLasTransacciones.map(function(tx) {
    const card    = tx.merchant || {};
    const details = tx.details  || {};
    const cxInfo  = tx.currencyExchangeInfo || {};

    const merchantName = card.name || details.merchantName || '';
    const description  = tx.counterpartyName || merchantName || tx.bankDescription || tx.externalMemo || tx.note || 'Sin detalles';
    const category     = (tx.mercuryCategory ? tx.mercuryCategory.name : '') || (tx.categoryData ? tx.categoryData.name : '') || '';

    return [
      tx.createdAt ? tx.createdAt.substring(0, 10)  : '',
      tx.createdAt ? new Date(tx.createdAt)          : '',
      description,
      tx.amount !== undefined ? tx.amount            : '',
      tx.status              || '',
      tx.bankDescription     || '',
      tx.externalMemo        || tx.referenceNumber   || '',
      tx.note                || '',
      category,
      tx.sourceOfCategory    || (tx.mercuryCategory ? 'Mercury' : '') || '',
      cxInfo.originalCurrency || tx.currency         || 'USD',
      tx.kind                || '',
      tx.counterpartyId      || '',
      tx.counterpartyNickname || '',
      tx.postedAt ? new Date(tx.postedAt)            : '',
      tx.dashboardLink       || '',
      tx.attachments ? tx.attachments.length         : 0,
      tx.id                  || ''
    ];
  });

  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, cabeceras.length).setValues(filas);
  }

  Logger.log(`✅ Proceso terminado. ${filas.length} transacciones cargadas, ordenadas de más antigua a más nueva.`);
}
