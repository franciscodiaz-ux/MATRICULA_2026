/**
 * ============================================================================
 * ARCHIVO: 99_Test.gs
 * ----------------------------------------------------------------------------
 * Archivo de pruebas para la hoja "Test".
 * Permite ejecutar pruebas controladas antes de integrar funciones al sistema.
 * ============================================================================
 */


/**
 * Crea el menú mínimo de pruebas para la hoja "99_Test".
 */
function crearMenuTest_() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("🧪 99_Test")
    .addItem("Probar carga del módulo", "test_confirmarCargaModulo")
    .addItem("Configurar hoja Test", "setupHojaTest")
    .addToUi();
}


/**
 * Confirma que el archivo 99_Test.gs está cargado correctamente.
 */
function test_confirmarCargaModulo() {
  SpreadsheetApp.getUi().alert('El módulo 99_Test está cargado correctamente.');
}

