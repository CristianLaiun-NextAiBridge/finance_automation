// ==========================================
// TRIGGERS — EJECUTAR UNA SOLA VEZ
// ==========================================

function setupTriggers() {
  // Eliminar triggers previos de actualizarTablaMercury para evitar duplicados
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarTablaMercury')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Trigger diario a las 5 AM (zona horaria del script)
  ScriptApp.newTrigger('actualizarTablaMercury')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();

  Logger.log('✅ Trigger creado: actualizarTablaMercury todos los días entre 5:00 y 6:00 AM.');
}

function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log('No hay triggers configurados.');
    return;
  }
  triggers.forEach(t => {
    Logger.log(`• ${t.getHandlerFunction()} — tipo: ${t.getEventType()} — ID: ${t.getUniqueId()}`);
  });
}
