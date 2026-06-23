// ==========================================
// MATCHING DE FACTURAS CON COMPROBANTES DE DRIVE
//
// Para cada transacción de la hoja de conciliación que no tiene Receipt
// asignado, busca en la carpeta de Drive (FOLDER_COMPROBANTES_ID) el archivo
// (PDF/imagen + JSON) que la ampara.
//
// El proceso es en dos pasos:
// 1. PRE-FILTRO: descarta candidatos que no coincidan en monto y fecha/proveedor
//    para reducir las llamadas a la IA.
// 2. VALIDACIÓN IA (batch): Gemini 2.5 Flash decide si el Receipt ampara la
//    transacción. Las llamadas se agrupan en lotes de 40 para no saturar
//    UrlFetchApp.fetchAll().
//
// Al confirmar un match escribe hipervínculos en las columnas "Receipt"
// (PDF) y "JSON Data" (JSON extraído).
//
// Depende de: Code.js (getServiceAccount, obtenerTokenDeAcceso,
//             preFiltrarCandidatosAmplio, ejecutarBatchIA)
// ==========================================

function matchearFacturasConDrive() {
  Logger.log("🔍 Iniciando matching de facturas con Receipts de Drive...");

  let token;
  let sa;
  try {
    sa    = getServiceAccount();
    token = obtenerTokenDeAcceso(sa);
  } catch(e) {
    Logger.log("❌ Error generando token: " + e.toString());
    return;
  }

  const targetSheet = SpreadsheetApp.openById(CONCILIATION_SHEET_ID).getSheetByName(LEDGER_TAB_NAME);
  const dataFinal   = targetSheet.getDataRange().getValues();
  const headers     = dataFinal[0];

  const ReceiptCol  = headers.indexOf('Receipt')    + 1;
  const jsonInfoCol     = headers.indexOf('JSON Data')       + 1;
  const finalAmtInIdx   = headers.indexOf('Amount In (+)');
  const finalAmtOutIdx  = headers.indexOf('Amount Out (-)');
  const finalDescIdx    = headers.indexOf('Description');
  const finalDateIdx    = headers.indexOf('Date (UTC)');

  if (ReceiptCol === 0 || jsonInfoCol === 0) {
    Logger.log("❌ No se encontraron las columnas 'Receipt' o 'JSON Data'. Ejecutá formatearHoja() primero.");
    return;
  }

  // ── 1. CARGAR COMPROBANTES DE DRIVE EN MEMORIA ───────────────────────────────
  Logger.log("📂 Cargando Receipts desde Drive...");
  const folder            = DriveApp.getFolderById(FOLDER_COMPROBANTES_ID);
  const files             = folder.getFiles();
  const jsonFiles         = [];
  const fileUrls          = {};
  const jsonUrlsByBaseName = {};

  while (files.hasNext()) {
    const file     = files.next();
    const name     = file.getName();
    const baseName = name.replace(/\.[^/.]+$/, "");
    const isJson   = name.toLowerCase().endsWith('.json');

    if (!isJson && !fileUrls[baseName]) fileUrls[baseName] = file.getUrl();

    if (isJson) {
      jsonUrlsByBaseName[baseName] = file.getUrl();
      try {
        const content = file.getBlob().getDataAsString();
        jsonFiles.push({ baseName: baseName, data: JSON.parse(content), contentStr: content, jsonUrl: file.getUrl() });
      } catch(e) {
        Logger.log("⚠️ Error leyendo JSON de Drive: " + name);
      }
    }
  }

  Logger.log("📁 " + jsonFiles.length + " JSONs de Receipts cargados.");

  // ── 2. AUTOCORRECCIÓN DE LINKS VIEJOS + ARMAR PETICIONES IA ─────────────────
  const peticiones        = [];
  const mapeo             = [];
  const filasYaMatcheadas = new Set();

  for (let i = 1; i < dataFinal.length; i++) {
    const row                    = dataFinal[i];
    const valorComprobanteActual = (row[ReceiptCol - 1] || "").toString();

    // Autocorrección: links crudos → fórmula HYPERLINK
    if (valorComprobanteActual.includes("http") && !valorComprobanteActual.startsWith("=")) {
      targetSheet.getRange(i + 1, ReceiptCol).setFormula('=HYPERLINK("' + valorComprobanteActual + '", "Link")');
      for (const name in fileUrls) {
        if (fileUrls[name] === valorComprobanteActual && jsonUrlsByBaseName[name]) {
          targetSheet.getRange(i + 1, jsonInfoCol).setFormula('=HYPERLINK("' + jsonUrlsByBaseName[name] + '", "Link")');
          break;
        }
      }
      continue;
    }

    if (valorComprobanteActual !== "") continue;

    const rawAmount = row[finalAmtInIdx] || row[finalAmtOutIdx] || '';
    if (!rawAmount) continue;
    const txAmount  = Math.abs(parseFloat(rawAmount.toString().replace(/,/g, '')));

    const txDate = row[finalDateIdx];
    const txDesc = (row[finalDescIdx] || '').toString().trim();

    const candidatos = preFiltrarCandidatosAmplio(txAmount, txDate, txDesc, jsonFiles);

    for (let j = 0; j < candidatos.length; j++) {
      const cand = candidatos[j];
      peticiones.push({
        prompt: 'Eres un auditor contable automatizando conciliaciones bancarias para una empresa corporativa.\n' +
                'Transacción Bancaria de la cuenta Mercury:\n' +
                '- Fecha de impacto bancario: ' + txDate + '\n' +
                '- Monto debitado: $' + txAmount + '\n' +
                '- Descripción del Banco: ' + txDesc + '\n\n' +
                'Datos extraídos del Receipt físico (JSON):\n' +
                cand.contentStr + '\n\n' +
                'Instrucciones: Determina si este Receipt ampara este movimiento. ' +
                'Sé flexible con fechas (pagos con tarjeta pueden reflejarse días después). ' +
                'Match si: 1. Monto exacto. 2. Proveedor concuerda. 3. Fecha lógica.\n' +
                'Responde ÚNICAMENTE con la palabra "SI" o "NO".'
      });
      mapeo.push({ fila: i + 1, json: cand, txAmount: txAmount });
    }
  }

  // ── 3. EJECUTAR BATCH IA Y ESCRIBIR MATCHES ──────────────────────────────────
  if (peticiones.length === 0) {
    Logger.log("ℹ️ No hay filas pendientes de matching.");
    return;
  }

  Logger.log("⚡ Procesando " + peticiones.length + " validaciones con IA...");
  const respuestas = ejecutarBatchIA(peticiones, token, sa);

  for (let k = 0; k < respuestas.length; k++) {
    const m = mapeo[k];
    if (filasYaMatcheadas.has(m.fila) || !respuestas[k]) continue;

    if (respuestas[k].getResponseCode() === 200) {
      try {
        const texto = JSON.parse(respuestas[k].getContentText())
                        .candidates[0].content.parts[0].text.trim().toUpperCase().replace(/Í/g, 'I');

        if (texto.includes("SI")) {
          const fileLink = fileUrls[m.json.baseName]      || "";
          const jsonLink = m.json.jsonUrl                  || "";

          if (fileLink) targetSheet.getRange(m.fila, ReceiptCol).setFormula('=HYPERLINK("' + fileLink + '", "Link")');
          if (jsonLink) targetSheet.getRange(m.fila, jsonInfoCol).setFormula('=HYPERLINK("' + jsonLink + '", "Link")');

          Logger.log("✅ Match: Fila " + m.fila + " [$" + m.txAmount + "] → " + m.json.baseName);
          filasYaMatcheadas.add(m.fila);
        }
      } catch(e) {
        Logger.log("⚠️ Error procesando respuesta fila " + m.fila + ": " + e.toString());
      }
    } else {
      Logger.log("❌ Error API fila " + m.fila + ": " + respuestas[k].getResponseCode());
    }
  }

  Logger.log("✅ Matching finalizado. " + filasYaMatcheadas.size + " facturas vinculadas.");
}
