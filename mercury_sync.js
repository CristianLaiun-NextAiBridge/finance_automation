// ==========================================
// SINCRONIZACIÓN DE TRANSACCIONES MERCURY
//
// Conecta con la API de Mercury para traer el historial de transacciones bancarias
// y escribirlo en el Google Sheet definido por MERCURY_SOURCE_SHEET_ID.
//
// En la primera ejecución (hoja vacía) hace una carga completa desde 2020-01-01.
// En las siguientes, reemplaza únicamente la ventana de los últimos 30 días para
// reflejar cambios de estado (pending → sent) sin reprocesar todo el histórico.
//
// Por cada transacción calcula el campo "Balance After": el saldo de la cuenta
// luego de ese movimiento. El cálculo trabaja hacia atrás desde el saldo real
// actual que devuelve GET /accounts, garantizando que la última fila siempre
// coincida con el saldo vigente.
//
// El token de Mercury se lee desde Script Properties (MERCURY_API_TOKEN).
// Para configurarlo ejecutá setupCredentials() en setup_credentials.js.
// ==========================================

const CABECERAS_MERCURY = [
  'Date (UTC)', 'Timestamp', 'Description', 'Amount', 'Balance After',
  'Status', 'Category', 'Source of Category', 'Original Currency', 'kind',
  'counterpartyId', 'counterpartyNickname', 'postedAt', 'dashboardLink', 'attachments', 'id'
];

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

  // ── 1. SALDO ACTUAL POR CUENTA ───────────────────────────────────────────────
  const saldosPorCuenta = {};
  try {
    const respCuentas = UrlFetchApp.fetch('https://api.mercury.com/api/v1/accounts', opciones);
    if (respCuentas.getResponseCode() === 200) {
      (JSON.parse(respCuentas.getContentText()).accounts || []).forEach(function(c) {
        saldosPorCuenta[c.id] = c.currentBalance;
        Logger.log('Cuenta: ' + c.name + ' | Saldo actual: $' + c.currentBalance);
      });
    }
  } catch(e) {
    Logger.log('⚠️ No se pudo obtener el saldo actual: ' + e.toString());
  }

  // ── 2. DETERMINAR MODO: COMPLETO O VENTANA ───────────────────────────────────
  const lastRow     = hoja.getLastRow();
  const esPrimeraVez = lastRow <= 1;

  const hace30Dias = new Date();
  hace30Dias.setDate(hace30Dias.getDate() - 30);
  const fechaCorte = Utilities.formatDate(hace30Dias, 'UTC', 'yyyy-MM-dd');
  const fechaInicio = esPrimeraVez ? '2020-01-01' : fechaCorte;

  Logger.log(esPrimeraVez
    ? '🚀 Primera ejecución: carga completa desde 2020-01-01'
    : '🔄 Modo incremental: actualizando ventana desde ' + fechaCorte
  );

  // ── 3. TRAER TRANSACCIONES ───────────────────────────────────────────────────
  let transacciones = [];
  let hayMas        = true;
  let startAfter    = "";

  while (hayMas) {
    let url = 'https://api.mercury.com/api/v1/transactions?limit=500&start=' + fechaInicio;
    if (startAfter !== "") url += '&start_after=' + startAfter;

    const resp  = UrlFetchApp.fetch(url, opciones);
    const datos = JSON.parse(resp.getContentText());

    if (resp.getResponseCode() !== 200) {
      Logger.log('❌ Error en la API de Mercury: ' + (datos.message || resp.getContentText()));
      return;
    }

    transacciones = transacciones.concat(datos.transactions || []);
    hayMas        = !!(datos.page && datos.page.nextPage);
    if (hayMas) startAfter = datos.page.nextPage;
  }

  if (transacciones.length === 0) {
    Logger.log('ℹ️ No se encontraron transacciones para el período.');
    return;
  }

  transacciones.sort(function(a, b) {
    return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
  });

  // ── 4. CALCULAR BALANCE AFTER ────────────────────────────────────────────────
  // Anclaje: currentBalance de la API es el saldo real después de la última tx.
  // Calculamos hacia atrás para asignar el saldo correcto a cada transacción.
  const balanceAfter = new Array(transacciones.length).fill('');

  const txPorCuenta = {};
  transacciones.forEach(function(tx, idx) {
    const aid = tx.accountId || 'unknown';
    if (!txPorCuenta[aid]) txPorCuenta[aid] = [];
    txPorCuenta[aid].push(idx);
  });

  Object.keys(txPorCuenta).forEach(function(accountId) {
    const indices     = txPorCuenta[accountId];
    const saldoActual = saldosPorCuenta[accountId];
    if (saldoActual === undefined) return;

    let saldo = saldoActual;
    for (let i = indices.length - 1; i >= 0; i--) {
      balanceAfter[indices[i]] = Math.round(saldo * 100) / 100;
      saldo = saldo - (transacciones[indices[i]].amount || 0);
    }
  });

  // ── 5. ENCONTRAR Y ELIMINAR FILAS DE LA VENTANA EN LA HOJA ──────────────────
  if (esPrimeraVez) {
    hoja.clearContents();
    hoja.getRange(1, 1, 1, CABECERAS_MERCURY.length).setValues([CABECERAS_MERCURY]);
  } else {
    // Buscar la primera fila cuya fecha >= fechaCorte
    const fechasDatos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    let filaCorte = -1;

    for (let i = 0; i < fechasDatos.length; i++) {
      let dateVal = fechasDatos[i][0];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, 'UTC', 'yyyy-MM-dd');
      }
      if (dateVal.toString() >= fechaCorte) {
        filaCorte = i + 2; // +1 por índice, +1 por header
        break;
      }
    }

    if (filaCorte !== -1) {
      const filasAEliminar = hoja.getLastRow() - filaCorte + 1;
      hoja.deleteRows(filaCorte, filasAEliminar);
      Logger.log('🗑️ Eliminadas ' + filasAEliminar + ' filas de la ventana anterior.');
    }
  }

  // ── 6. ESCRIBIR NUEVAS FILAS ─────────────────────────────────────────────────
  const filas = transacciones.map(function(tx, idx) {
    const card    = tx.merchant || {};
    const details = tx.details  || {};
    const cxInfo  = tx.currencyExchangeInfo || {};

    const merchantName = card.name || details.merchantName || '';
    const category     = (tx.mercuryCategory ? tx.mercuryCategory.name : '') || (tx.categoryData ? tx.categoryData.name : '') || '';

    // Unificar todas las fuentes de descripción en un solo campo, eliminando duplicados
    const descParts = [
      tx.counterpartyName  || merchantName      || '',
      tx.bankDescription                        || '',
      tx.externalMemo      || tx.referenceNumber || '',
      tx.note                                   || ''
    ].map(function(s) { return s.toString().trim(); }).filter(Boolean);

    const seen = new Set();
    const description = descParts.filter(function(p) {
      const key = p.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    }).join(' | ') || 'Sin detalles';

    return [
      tx.createdAt ? tx.createdAt.substring(0, 10)               : '',
      tx.createdAt ? new Date(tx.createdAt)                       : '',
      description,
      tx.amount !== undefined ? tx.amount                         : '',
      balanceAfter[idx],
      tx.status                                                   || '',
      category,
      tx.sourceOfCategory || (tx.mercuryCategory ? 'Mercury' : '') || '',
      cxInfo.originalCurrency || tx.currency                      || 'USD',
      tx.kind                                                     || '',
      tx.counterpartyId                                           || '',
      tx.counterpartyNickname                                     || '',
      tx.postedAt ? new Date(tx.postedAt)                         : '',
      tx.dashboardLink                                            || '',
      tx.attachments ? tx.attachments.length                      : 0,
      tx.id                                                       || ''
    ];
  });

  const insertRow = hoja.getLastRow() + 1;
  hoja.getRange(insertRow, 1, filas.length, CABECERAS_MERCURY.length).setValues(filas);

  Logger.log('✅ ' + filas.length + ' transacciones escritas desde fila ' + insertRow + '.');
}
