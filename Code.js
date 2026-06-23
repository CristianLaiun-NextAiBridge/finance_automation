// ==========================================
// CONSTANTES DE CONFIGURACIÓN
// ==========================================

// ID de la carpeta de Drive donde están los PDFs y JSONs de comprobantes
const FOLDER_COMPROBANTES_ID = "1rNWOhiTWnX7SvzAf-mg_LaVEoeJN0DP5";

// ID del Google Sheet de conciliación — contiene todas las pestañas del proyecto
const CONCILIATION_SHEET_ID = "19-SS1TaCDNcSZvZsTHpIdwudZSzrssbagIHmNl4hj5U";

// Nombre de la pestaña donde se vuelca la raw data de Mercury
const MERCURY_TAB_NAME = 'mercury';

// Nombre de la pestaña principal de conciliación
const LEDGER_TAB_NAME = 'Ledger';

// Nombre de la pestaña de configuración (categorías, etc.)
const SETUP_TAB_NAME = 'setup';

// Columnas de la pestaña mercury que no se copian a la hoja de conciliación
const COLUMNAS_A_SACAR = ['counterpartyId', 'counterpartyNickname', 'id'];

// ==========================================
// CREDENCIALES
// Leídas desde Script Properties — nunca hardcodeadas.
// Configurar ejecutando setupCredentials() en setup_credentials.js
// ==========================================
function getServiceAccount() {
  const props       = PropertiesService.getScriptProperties();
  const projectId   = props.getProperty('SA_PROJECT_ID');
  const clientEmail = props.getProperty('SA_CLIENT_EMAIL');
  const privateKey  = props.getProperty('SA_PRIVATE_KEY');

  if (!projectId || !clientEmail || !privateKey) {
    throw new Error(
      'Faltan credenciales de Service Account. ' +
      'Ejecutá setupCredentials() en setup_credentials.js.'
    );
  }
  return { project_id: projectId, client_email: clientEmail, private_key: privateKey };
}

// ==========================================
// ORQUESTADOR PRINCIPAL
// Ejecuta el pipeline completo en orden:
// 1. Formatea la hoja y asigna categorías (formatting.js)
// 2. Matchea facturas con comprobantes de Drive (invoice_matching.js)
// ==========================================
function procesarConciliacion() {
  Logger.log("🚀 Iniciando pipeline de conciliación contable...");
  formatearHoja();
  matchearFacturasConDrive();
  Logger.log("✅ Pipeline finalizado.");
}

// ==========================================
// HELPERS COMPARTIDOS
// Usados tanto por formatting.js como por invoice_matching.js
// ==========================================

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

  const resp = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method:  "post",
    payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt }
  });
  return JSON.parse(resp.getContentText()).access_token;
}

function preFiltrarCandidatosAmplio(txAmount, txDate, txDesc, jsonFiles) {
  const candidatos    = [];
  const fechaBanco    = new Date(txDate);
  const txDescLower   = txDesc.toLowerCase();
  const palabrasClave = txDescLower.replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(w => w.length > 1);

  for (let i = 0; i < jsonFiles.length; i++) {
    const json     = jsonFiles[i];
    const strLower = json.contentStr.toLowerCase();

    const amountStr       = txAmount.toString();
    const amountWithComma = amountStr.replace('.', ',');
    let montoAparece      = strLower.includes(amountStr) || strLower.includes(amountWithComma);

    if (!montoAparece && json.data && json.data.total) {
      const cleanTotal = json.data.total.toString().replace(/,/g, '');
      montoAparece     = Math.abs(parseFloat(cleanTotal)) === txAmount;
    }
    if (!montoAparece) continue;

    let fechaAceptable = false;
    if (json.data && json.data.fecha) {
      const fechaComprobante = new Date(json.data.fecha);
      if (!isNaN(fechaComprobante.getTime())) {
        const diff = (fechaBanco - fechaComprobante) / (1000 * 60 * 60 * 24);
        if (diff >= -5 && diff <= 20) fechaAceptable = true;
      } else {
        const fStr = json.data.fecha.toString().replace(/[^0-9]/g, '');
        if (fStr.length === 8) {
          const fComp = new Date(fStr.substring(0,4), parseInt(fStr.substring(4,6))-1, fStr.substring(6,8));
          const diff  = (fechaBanco - fComp) / (1000 * 60 * 60 * 24);
          if (diff >= -5 && diff <= 20) fechaAceptable = true;
        }
      }
    }

    const proveedorCoincide = palabrasClave.some(p => strLower.includes(p));
    if (fechaAceptable || proveedorCoincide) candidatos.push(json);
  }

  return candidatos;
}

// Ejecuta llamadas a Vertex AI en lotes de tamaño seguro y devuelve las respuestas
// en el mismo orden que las peticiones.
function ejecutarBatchIA(peticiones, token, sa) {
  if (peticiones.length === 0) return [];

  const requests = peticiones.map(function(p) {
    return {
      url:            'https://us-central1-aiplatform.googleapis.com/v1/projects/' +
                      sa.project_id +
                      '/locations/us-central1/publishers/google/models/gemini-2.5-flash:generateContent',
      method:         'post',
      contentType:    'application/json',
      headers:        { 'Authorization': 'Bearer ' + token },
      payload:        JSON.stringify({
        contents:         [{ role: 'user', parts: [{ text: p.prompt }] }],
        generationConfig: { temperature: 0.0 }
      }),
      muteHttpExceptions: true
    };
  });

  let respuestas = [];
  const TAMANO_LOTE = 40;
  for (let i = 0; i < requests.length; i += TAMANO_LOTE) {
    try {
      respuestas = respuestas.concat(UrlFetchApp.fetchAll(requests.slice(i, i + TAMANO_LOTE)));
      Utilities.sleep(1500);
    } catch(e) {
      Logger.log('⚠️ Error en lote ' + i + ': ' + e.toString());
      respuestas = respuestas.concat(new Array(Math.min(TAMANO_LOTE, requests.length - i)).fill(null));
    }
  }
  return respuestas;
}
