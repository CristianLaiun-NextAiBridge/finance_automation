// ==========================================
// CONFIGURACIÓN (IDs no sensibles — pueden vivir aquí)
// ==========================================

// ID de la carpeta donde están los PDFs y JSONs en Google Drive
const FOLDER_COMPROBANTES_ID = "1rNWOhiTWnX7SvzAf-mg_LaVEoeJN0DP5";

// ID del archivo de Google Sheets original que contiene la data de Mercury
const MERCURY_SOURCE_SHEET_ID = "1hWPRWPka_yOFpp6nnwVcS4zxNZ62VU2ipVMXGIgsm4c";

// Columnas que queremos excluir de la hoja original
const COLUMNAS_A_SACAR = ['counterpartyId', 'counterpartyNickname', 'id'];

// ==========================================
// CREDENCIALES — leídas desde Script Properties (nunca hardcodeadas)
// Para configurarlas: ejecutá setupCredentials() desde setup_credentials.js
// o cargalas manualmente en: Apps Script > Configuración > Propiedades de secuencia de comandos
// ==========================================
function getServiceAccount() {
  const props = PropertiesService.getScriptProperties();
  const projectId   = props.getProperty('SA_PROJECT_ID');
  const clientEmail = props.getProperty('SA_CLIENT_EMAIL');
  const privateKey  = props.getProperty('SA_PRIVATE_KEY');

  if (!projectId || !clientEmail || !privateKey) {
    throw new Error(
      'Faltan credenciales de Service Account. ' +
      'Ejecutá setupCredentials() en setup_credentials.js ' +
      'o cargalas en: Apps Script > Configuración > Propiedades de secuencia de comandos.'
    );
  }

  return { project_id: projectId, client_email: clientEmail, private_key: privateKey };
}


// ==========================================
// FUNCIÓN PRINCIPAL OPTIMIZADA
// ==========================================
function matchearFacturasMercuryConDrive() {
  Logger.log("🚀 Iniciando proceso incremental y batch de conciliación contable...");

  let token = "";
  try {
    const sa = getServiceAccount();
    token = obtenerTokenDeAcceso(sa);
  } catch (e) {
    Logger.log("❌ Error generando Token de IA: " + e.toString());
    return;
  }

  // 1. OBTENER DATA ACTUAL DE LA HOJA DESTINO
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let currentData = [];
  if (targetSheet.getLastRow() > 0) {
    currentData = targetSheet.getDataRange().getValues();
  }

  let headers = currentData.length > 0 ? currentData[0] : [];

  const existingTxIds = new Set();
  const currentIdIdx     = headers.indexOf('id');
  let   currentDateIdx   = headers.indexOf('Timestamp');
  if (currentDateIdx === -1) currentDateIdx = headers.indexOf('Date (UTC)');
  const currentAmountIdx = headers.indexOf('Amount');
  const currentDescIdx   = headers.indexOf('Description');

  if (currentData.length > 1) {
    for (let r = 1; r < currentData.length; r++) {
      if (currentIdIdx !== -1 && currentData[r][currentIdIdx]) {
        existingTxIds.add(currentData[r][currentIdIdx].toString().trim());
      } else if (currentDateIdx !== -1 && currentAmountIdx !== -1) {
        let dateVal = currentData[r][currentDateIdx];
        if (dateVal instanceof Date) {
          dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        let rawAmount = currentData[r][currentAmountIdx];
        let amountVal = rawAmount ? Math.abs(parseFloat(rawAmount)).toFixed(2) : "0.00";
        const descVal = (currentData[r][currentDescIdx] || '').toString().toLowerCase().trim();
        existingTxIds.add(`${dateVal}_${amountVal}_${descVal}`);
      }
    }
  }

  // 2. TRAER DATA DEL ARCHIVO FUENTE DE MERCURY
  let sourceData = [];
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(MERCURY_SOURCE_SHEET_ID);
    sourceData = sourceSpreadsheet.getSheets()[0].getDataRange().getValues();
  } catch (e) {
    Logger.log("❌ Error accediendo al archivo fuente de Mercury: " + e.toString());
    return;
  }

  const sourceHeaders = sourceData[0];
  const sourceIdIdx   = sourceHeaders.indexOf('id');
  const amountIdx     = sourceHeaders.indexOf('Amount');
  let   dateIdx       = sourceHeaders.indexOf('Timestamp');
  if (dateIdx === -1) dateIdx = sourceHeaders.indexOf('Date (UTC)');
  const descIdx       = sourceHeaders.indexOf('Description');

  if (sourceIdIdx === -1 || amountIdx === -1 || dateIdx === -1) {
    Logger.log("❌ El archivo origen de Mercury no tiene el formato esperado (falta id, Amount o Timestamp).");
    return;
  }

  const indicesAOmitir = COLUMNAS_A_SACAR.map(col => sourceHeaders.indexOf(col)).filter(idx => idx !== -1);

  if (headers.length === 0) {
    headers = sourceHeaders.filter((_, idx) => !indicesAOmitir.includes(idx));
    targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  let comprobanteCol = headers.indexOf('comprobante') + 1;
  if (comprobanteCol === 0) {
    comprobanteCol = targetSheet.getLastColumn() + 1;
    targetSheet.getRange(1, comprobanteCol).setValue('comprobante').setFontWeight("bold");
    headers.push('comprobante');
  }

  let jsonInfoCol = headers.indexOf('JSON Data') + 1;
  if (jsonInfoCol === 0) {
    jsonInfoCol = targetSheet.getLastColumn() + 1;
    targetSheet.getRange(1, jsonInfoCol).setValue('JSON Data').setFontWeight("bold");
    headers.push('JSON Data');
  }

  // 3. COMPARACIÓN Y CARGA DE NUEVOS REGISTROS
  const nuevasFilasAPegar = [];
  for (let s = 1; s < sourceData.length; s++) {
    const sourceRow = sourceData[s];
    const txId = sourceRow[sourceIdIdx];

    let txDate = sourceRow[dateIdx];
    if (txDate instanceof Date) {
      txDate = Utilities.formatDate(txDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    let rawSourceAmount = sourceRow[amountIdx];
    let txAmount = rawSourceAmount ? Math.abs(parseFloat(rawSourceAmount)).toFixed(2) : "0.00";
    const txDesc = (sourceRow[descIdx] || '').toString().toLowerCase().trim();
    const compKey = `${txDate}_${txAmount}_${txDesc}`;

    if ((txId && existingTxIds.has(txId.toString().trim())) || existingTxIds.has(compKey)) {
      continue;
    }

    const nuevaFila = new Array(headers.length).fill("");
    let targetIdx = 0;
    for (let c = 0; c < sourceRow.length; c++) {
      if (!indicesAOmitir.includes(c)) {
        nuevaFila[targetIdx] = sourceRow[c];
        targetIdx++;
      }
    }
    nuevasFilasAPegar.push(nuevaFila);
  }

  if (nuevasFilasAPegar.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, nuevasFilasAPegar.length, nuevasFilasAPegar[0].length).setValues(nuevasFilasAPegar);
    Logger.log(`📥 Se añadieron ${nuevasFilasAPegar.length} nuevos registros.`);
  } else {
    Logger.log("ℹ️ No se detectaron transacciones nuevas en Mercury.");
  }

  // 4. CARGAR COMPROBANTES DESDE DRIVE EN MEMORIA
  Logger.log("📂 Sincronizando repositorio de comprobantes de Drive...");
  const folder = DriveApp.getFolderById(FOLDER_COMPROBANTES_ID);
  const files = folder.getFiles();
  const jsonFiles = [];
  const fileUrls = {};
  const jsonUrlsByBaseName = {};

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const baseName = name.replace(/\.[^/.]+$/, "");
    const isJson = name.toLowerCase().endsWith('.json');

    if (!isJson || !fileUrls[baseName]) {
      fileUrls[baseName] = file.getUrl();
    }

    if (isJson) {
      jsonUrlsByBaseName[baseName] = file.getUrl();
      try {
        const content = file.getBlob().getDataAsString();
        jsonFiles.push({
          baseName:   baseName,
          data:       JSON.parse(content),
          contentStr: content,
          jsonUrl:    file.getUrl()
        });
      } catch(e) {
        Logger.log(`⚠️ Error leyendo JSON de Drive: ${name}`);
      }
    }
  }

  // =========================================================================
  // 5. PROCESAMIENTO INTELIGENTE Y AUTOCORRECCIÓN (VERSIÓN BATCH / PARALELO)
  // =========================================================================
  const sa = getServiceAccount();
  const dataFinal = targetSheet.getDataRange().getValues();
  const finalAmountIdx    = headers.indexOf('Amount');
  let   finalDateIdx      = headers.indexOf('Timestamp');
  if (finalDateIdx === -1) finalDateIdx = headers.indexOf('Date (UTC)');
  const finalDescIdx      = headers.indexOf('Description');
  const finalBankDescIdx  = headers.indexOf('Bank Description');

  const peticionesAI  = [];
  const mapeoPeticiones = [];
  const filasYaMatcheadas = new Set();

  for (let i = 1; i < dataFinal.length; i++) {
    const row = dataFinal[i];
    const valorComprobanteActual = (row[comprobanteCol - 1] || "").toString();

    if (valorComprobanteActual.includes("http") && !valorComprobanteActual.startsWith("=")) {
      let baseNameEncontrado = null;
      for (let name in fileUrls) {
        if (fileUrls[name] === valorComprobanteActual) {
          baseNameEncontrado = name;
          break;
        }
      }
      targetSheet.getRange(i + 1, comprobanteCol).setFormula(`=HYPERLINK("${valorComprobanteActual}", "Link")`);
      if (baseNameEncontrado && jsonUrlsByBaseName[baseNameEncontrado]) {
        targetSheet.getRange(i + 1, jsonInfoCol).setFormula(`=HYPERLINK("${jsonUrlsByBaseName[baseNameEncontrado]}", "Link")`);
      }
      continue;
    }

    if (valorComprobanteActual !== "") continue;

    const amountStr = row[finalAmountIdx];
    if (!amountStr || amountStr === "") continue;

    const cleanAmountStr = amountStr.toString().replace(/,/g, '');
    const txAmount = Math.abs(parseFloat(cleanAmountStr));

    const txDate   = row[finalDateIdx];
    const bankDesc = finalBankDescIdx !== -1 ? row[finalBankDescIdx] : '';
    const txDesc   = `${row[finalDescIdx] || ''} ${bankDesc || ''}`.trim();

    const candidatos = preFiltrarCandidatosAmplio(txAmount, txDate, txDesc, jsonFiles);

    for (let j = 0; j < candidatos.length; j++) {
      const jsonCandidato = candidatos[j];
      const prompt = `Eres un auditor contable automatizando conciliaciones bancarias para una empresa corporativa.
Transacción Bancaria de la cuenta Mercury:
- Fecha de impacto bancario: ${txDate}
- Monto debitado: $${txAmount}
- Descripción del Banco: ${txDesc}

Datos extraídos del comprobante físico (JSON):
${jsonCandidato.contentStr}

Instrucciones: Determina si este comprobante ampara este movimiento. Sé flexible con fechas (pagos con tarjeta pueden reflejarse días después). Match si: 1. Monto exacto. 2. Proveedor concuerda. 3. Fecha lógica.
Responde ÚNICAMENTE con la palabra "SI" o "NO".`;

      peticionesAI.push({
        url:            `https://us-central1-aiplatform.googleapis.com/v1/projects/${sa.project_id}/locations/us-central1/publishers/google/models/gemini-2.5-flash:generateContent`,
        method:         "post",
        contentType:    "application/json",
        headers:        { "Authorization": "Bearer " + token },
        payload:        JSON.stringify({
          contents:         [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.0 }
        }),
        muteHttpExceptions: true
      });

      mapeoPeticiones.push({ fila: i + 1, json: jsonCandidato, txAmount: txAmount });
    }
  }

  // Ejecutar el Batch de llamadas a IA EN LOTES
  if (peticionesAI.length > 0) {
    Logger.log(`⚡ Procesando ${peticionesAI.length} validaciones con IA en lotes...`);

    let respuestas = [];
    const TAMANO_LOTE = 40;

    for (let i = 0; i < peticionesAI.length; i += TAMANO_LOTE) {
      const lote = peticionesAI.slice(i, i + TAMANO_LOTE);
      try {
        respuestas = respuestas.concat(UrlFetchApp.fetchAll(lote));
        Utilities.sleep(1500);
      } catch (e) {
        Logger.log(`⚠️ Error procesando el lote ${i}: ${e.toString()}`);
        respuestas = respuestas.concat(new Array(lote.length).fill(null));
      }
    }

    for (let k = 0; k < respuestas.length; k++) {
      const mapeo = mapeoPeticiones[k];
      if (filasYaMatcheadas.has(mapeo.fila)) continue;
      if (!respuestas[k]) continue;

      if (respuestas[k].getResponseCode() === 200) {
        try {
          const jsonRespuesta = JSON.parse(respuestas[k].getContentText());
          const iaRespuesta   = jsonRespuesta.candidates[0].content.parts[0].text.trim().toUpperCase().replace(/Í/g, 'I');

          if (iaRespuesta.includes("SI")) {
            const fileLink = fileUrls[mapeo.json.baseName] || "";
            const jsonLink = mapeo.json.jsonUrl || "";

            if (fileLink !== "") targetSheet.getRange(mapeo.fila, comprobanteCol).setFormula(`=HYPERLINK("${fileLink}", "Link")`);
            if (jsonLink !== "") targetSheet.getRange(mapeo.fila, jsonInfoCol).setFormula(`=HYPERLINK("${jsonLink}", "Link")`);

            Logger.log(`✅ Match IA exitoso: Fila ${mapeo.fila} [$${mapeo.txAmount}] -> ${mapeo.json.baseName}`);
            filasYaMatcheadas.add(mapeo.fila);
          }
        } catch (e) {
          Logger.log(`⚠️ Error procesando respuesta IA en fila ${mapeo.fila}: ${e.toString()}`);
        }
      } else {
        Logger.log(`❌ Error de API en fila ${mapeo.fila}: Código ${respuestas[k].getResponseCode()} - ${respuestas[k].getContentText()}`);
      }
    }
  } else {
    Logger.log("ℹ️ No hay nuevas filas pendientes de validación con IA.");
  }

  // 6. ESTILIZADO DE LA HOJA
  const lastCol = targetSheet.getLastColumn();
  if (lastCol > 0) {
    const headerRange = targetSheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground("#4a86e8").setFontColor("#ffffff").setFontWeight("bold");
    targetSheet.setFrozenRows(1);

    if (!targetSheet.getFilter()) headerRange.createFilter();

    const columnasAOcultar = ['Bank Description', 'Reference', 'Note', 'Category', 'Source of Category', 'Source of Categ', 'attachments', 'Timestamp', 'Date (UTC)', 'dashboardLink', 'postedAt'];
    columnasAOcultar.forEach(colName => {
      const idx = headers.indexOf(colName);
      if (idx !== -1) targetSheet.hideColumns(idx + 1);
    });
  }

  Logger.log("✅ ¡Proceso finalizado con éxito!");
}


// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

function preFiltrarCandidatosAmplio(txAmount, txDate, txDesc, jsonFiles) {
  const candidatos   = [];
  const fechaBanco   = new Date(txDate);
  const txDescLower  = txDesc.toLowerCase();
  const palabrasClave = txDescLower.replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(w => w.length > 1);

  for (let i = 0; i < jsonFiles.length; i++) {
    const json     = jsonFiles[i];
    const strLower = json.contentStr.toLowerCase();

    const amountStr       = txAmount.toString();
    const amountWithComma = amountStr.replace('.', ',');
    let montoAparece = strLower.includes(amountStr) || strLower.includes(amountWithComma);

    if (!montoAparece && json.data && json.data.total) {
      const cleanTotal = json.data.total.toString().replace(/,/g, '');
      montoAparece = Math.abs(parseFloat(cleanTotal)) === txAmount;
    }

    if (!montoAparece) continue;

    let fechaAceptable = false;
    if (json.data && json.data.fecha) {
      const fechaComprobante = new Date(json.data.fecha);
      if (!isNaN(fechaComprobante.getTime())) {
        const diferenciaDias = (fechaBanco - fechaComprobante) / (1000 * 60 * 60 * 24);
        if (diferenciaDias >= -5 && diferenciaDias <= 20) fechaAceptable = true;
      } else {
        const fStr = json.data.fecha.toString().replace(/[^0-9]/g, '');
        if (fStr.length === 8) {
          const fComp = new Date(fStr.substring(0,4), parseInt(fStr.substring(4,6))-1, fStr.substring(6,8));
          const diff  = (fechaBanco - fComp) / (1000 * 60 * 60 * 24);
          if (diff >= -5 && diff <= 20) fechaAceptable = true;
        }
      }
    }

    const proveedorCoincide = palabrasClave.some(palabra => strLower.includes(palabra));

    if (fechaAceptable || proveedorCoincide) {
      candidatos.push(json);
    }
  }

  return candidatos;
}

function obtenerTokenDeAcceso(sa) {
  const header  = JSON.stringify({ alg: "RS256", typ: "JWT" });
  const ahora   = Math.floor(Date.now() / 1000);
  const payload = JSON.stringify({
    iss:   sa.client_email,
    scope: "https://www.googleapis.com/auth/cloud-platform",
    aud:   "https://oauth2.googleapis.com/token",
    exp:   ahora + 3600,
    iat:   ahora
  });

  const base64Header  = Utilities.base64EncodeWebSafe(header).replace(/=+$/, '');
  const base64Payload = Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '');
  const firmaInput    = base64Header + "." + base64Payload;
  const firmaBinaria  = Utilities.computeRsaSha256Signature(firmaInput, sa.private_key);
  const base64Firma   = Utilities.base64EncodeWebSafe(firmaBinaria).replace(/=+$/, '');
  const jwt           = firmaInput + "." + base64Firma;

  const respuesta = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method:  "post",
    payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt }
  });
  return JSON.parse(respuesta.getContentText()).access_token;
}
