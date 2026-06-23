// ==========================================
// CONFIGURACIÓN DE TRIGGERS — EJECUTAR UNA SOLA VEZ
// ==========================================
// Horarios:
//   5 AM  → actualizarTablaMercury  (bajada de Mercury)
//   6 AM  → procesarLedger          (formateo + categorías + matching)
//   3 PM  → actualizarTablaMercury  (segunda bajada)
//   4 PM  → procesarLedger          (segundo procesamiento)
//
// Nota: Apps Script ejecuta cada trigger dentro de una ventana de ~1 hora
// a partir de la hora indicada (ej. atHour(5) corre entre 5:00 y 6:00 AM).
// ==========================================

function setupTriggers() {
  // Eliminar todos los triggers existentes del proyecto
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // ── Bajada de Mercury: 5 AM y 3 PM ───────────────────────────────────────────
  ScriptApp.newTrigger('actualizarTablaMercury')
    .timeBased().everyDays(1).atHour(5).create();

  ScriptApp.newTrigger('actualizarTablaMercury')
    .timeBased().everyDays(1).atHour(15).create();

  // ── Pipeline del Ledger: 6 AM y 4 PM ────────────────────────────────────────
  ScriptApp.newTrigger('procesarLedger')
    .timeBased().everyDays(1).atHour(6).create();

  ScriptApp.newTrigger('procesarLedger')
    .timeBased().everyDays(1).atHour(16).create();

  Logger.log('✅ 4 triggers configurados:');
  Logger.log('   • actualizarTablaMercury → 5 AM y 3 PM');
  Logger.log('   • procesarLedger         → 6 AM y 4 PM');
}

function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) { Logger.log('No hay triggers configurados.'); return; }
  triggers.forEach(function(t) {
    Logger.log('• ' + t.getHandlerFunction() + ' — ' + t.getEventType() + ' — ID: ' + t.getUniqueId());
  });
}
