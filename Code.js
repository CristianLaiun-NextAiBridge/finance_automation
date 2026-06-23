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

// Transacciones internas de Mercury que nunca tienen comprobante físico
const TX_SIN_COMPROBANTE = [
  'mercury', 'cashback', 'io autopay', 'autopay', 'bank reward', 'checking •'
];

// Filtra candidatos con criterios estrictos y devuelve máximo 3, ordenados por score.
// Requisitos excluyentes: monto Y fecha deben coincidir.
// La similitud de proveedor suma puntos pero no es obligatoria.
function preFiltrarCandidatosAmplio(txAmount, txDate, txDesc, jsonFiles) {
  const txDescLower = txDesc.toLowerCase();

  // Saltar transacciones internas de Mercury
  if (TX_SIN_COMPROBANTE.some(function(p) { return txDescLower.includes(p); })) return [];

  const fechaBanco    = new Date(txDate);
  const palabrasClave = txDescLower.replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(function(w) { return w.length > 2; });
  const scored        = [];

  for (let i = 0; i < jsonFiles.length; i++) {
    const json = jsonFiles[i];
    let score  = 0;

    // 1. MONTO — requerido. Primero intenta datos estructurados (más rápido)
    let montoMatch = false;
    if (json.data && json.data.total) {
      montoMatch = Math.abs(parseFloat(json.data.total.toString().replace(/,/g, ''))) === txAmount;
    }
    if (!montoMatch) {
      const s = json.contentStr.toLowerCase();
      const a = txAmount.toString();
      montoMatch = s.includes(a) || s.includes(a.replace('.', ','));
    }
    if (!montoMatch) continue;
    score += 3;

    // 2. FECHA — requerida, ventana -5 a +25 días
    const fechaDoc = _parseFechaJSON(json.data && json.data.fecha);
    if (!fechaDoc) continue;
    const diasDif = (fechaBanco - fechaDoc) / 86400000;
    if (diasDif < -5 || diasDif > 25) continue;
    score += diasDif <= 7 ? 3 : (diasDif <= 14 ? 2 : 1);

    // 3. PROVEEDOR — opcional, suma puntos
    const strLower = json.contentStr.toLowerCase();
    score += palabrasClave.filter(function(p) { return strLower.includes(p); }).length * 2;

    scored.push({ json: json, score: score });
  }

  // Top 3 candidatos por score
  scored.sort(function(a, b) { return b.score - a.score; });
  return scored.slice(0, 3).map(function(c) { return c.json; });
}

function _parseFechaJSON(fechaStr) {
  if (!fechaStr) return null;
  const d = new Date(fechaStr);
  if (!isNaN(d.getTime())) return d;
  const s = fechaStr.toString().replace(/[^0-9]/g, '');
  if (s.length === 8) {
    return new Date(s.substring(0,4), parseInt(s.substring(4,6)) - 1, s.substring(6,8));
  }
  return null;
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
