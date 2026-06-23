// ==========================================
// SETUP DE CREDENCIALES — EJECUTAR UNA SOLA VEZ
// ==========================================
// INSTRUCCIONES:
// 1. Abrí este archivo SOLO desde el editor de Apps Script (nunca subas la versión con
//    credenciales reales a GitHub — este archivo está en .gitignore).
// 2. Reemplazá los placeholders con los valores reales de tu Service Account JSON.
// 3. Ejecutá la función setupCredentials() desde el editor de Apps Script.
// 4. Verificá con verifyCredentials() que todo quedó guardado.
// 5. Borrá los valores reales de este archivo antes de hacer git push (o no lo toques,
//    ya que .gitignore lo protege localmente).
// ==========================================

function setupCredentials() {
  const props = PropertiesService.getScriptProperties();

  props.setProperties({
    // Reemplazá estos valores con los del JSON de tu Service Account:
    'SA_PROJECT_ID':   'REEMPLAZAR_CON_project_id',
    'SA_CLIENT_EMAIL': 'REEMPLAZAR_CON_client_email',

    // Pegá la private_key completa incluyendo los saltos de línea (\n):
    'SA_PRIVATE_KEY':  '-----BEGIN PRIVATE KEY-----\nREEMPLAZAR_CON_private_key\n-----END PRIVATE KEY-----\n'
  });

  Logger.log('✅ Credenciales guardadas en Script Properties correctamente.');
  Logger.log('   Podés verificar en: Apps Script > Configuración > Propiedades de secuencia de comandos');
}

function verifyCredentials() {
  const props     = PropertiesService.getScriptProperties();
  const projectId = props.getProperty('SA_PROJECT_ID');
  const email     = props.getProperty('SA_CLIENT_EMAIL');
  const key       = props.getProperty('SA_PRIVATE_KEY');

  Logger.log('SA_PROJECT_ID:   ' + (projectId   ? '✅ ' + projectId                       : '❌ No configurado'));
  Logger.log('SA_CLIENT_EMAIL: ' + (email        ? '✅ ' + email                           : '❌ No configurado'));
  Logger.log('SA_PRIVATE_KEY:  ' + (key          ? '✅ Cargada (' + key.length + ' chars)' : '❌ No configurada'));
}

function clearCredentials() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('🗑️ Todas las Script Properties fueron eliminadas.');
}
