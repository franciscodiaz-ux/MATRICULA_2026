/**
 * Archivo: 00_Setup.gs
 * ----------------------------------------------------
 * Configuración inicial del sistema MATRICULA_2026.
 *
 * Este archivo contiene funciones de configuración general
 * del documento y funciones reutilizables para preparar hojas
 * bajo criterios de formato chileno.
 *
 * ============================================================
 * ESTRUCTURA GENERAL DEL ARCHIVO
 * ============================================================
 *
 * 1. Configuración general Chile
 * ------------------------------------------------------------
 * Funciones encargadas de configurar el documento completo:
 *
 * - locale Chile
 * - timezone America/Santiago
 * - formatos base
 * - estilos visuales generales
 *
 *
 * 2. Configuración por columnas
 * ------------------------------------------------------------
 * Funciones reutilizables para configurar columnas específicas
 * según su tipo de dato.
 *
 * Ejemplos:
 *
 * - configurarColumnaRut()
 * - configurarColumnaDV()
 * - configurarColumnaTelefonoChile()
 * - configurarColumnaCelularChile()
 * - configurarColumnaFechaChile()
 * - configurarColumnaMonedaCLP()
 * - configurarColumnaCorreo()
 * - etc
 *
 *
 * 3. Setup por hoja
 * ------------------------------------------------------------
 * Cada hoja del sistema tendrá su propia función de setup.
 *
 * Estas funciones:
 *
 * - aplican el formato base de hoja;
 * - llaman únicamente a las configuraciones necesarias
 *   para esa hoja;
 * - jamás deben modificar hojas no definidas explícitamente.
 *
 * Ejemplos:
 *
 * setupHojaAlumnos()
 *    ├── configurarHojaBaseChile_()
 *    ├── configurarColumnaRut()
 *    ├── configurarColumnaDV()
 *    ├── configurarColumnaFechaChile()
 *    └── etc
 *
 * setupHojaCursos()
 *    ├── configurarHojaBaseChile_()
 *    └── columnas específicas
 *
 *
 * ============================================================
 * REGLAS IMPORTANTES DEL PROYECTO
 * ============================================================
 *
 * - Ninguna función debe recorrer todas las hojas automáticamente.
 *
 * - Una hoja solo puede ser configurada si fue indicada
 *   explícitamente.
 *
 * - Las funciones de columnas son herramientas reutilizables.
 *
 * - Las funciones de setup por hoja deciden qué columnas
 *   configurar.
 *
 * - No utilizar SpreadsheetApp.getUi() en funciones internas
 *   o reutilizables.
 *
 * - Toda nueva función debe incluir encabezado descriptivo.
 *
 */





 /*
 * 1. Configuración general Chile
 * ------------------------------------------------------------
 * Funciones encargadas de configurar el documento completo:
 *
 * - locale Chile
 * - timezone America/Santiago
 * - formatos base
 * - estilos visuales generales

/**
 * Ejecuta la configuración general del sistema bajo formato chileno.
 *
 * Esta función puede ejecutarse manualmente desde el editor
 * o desde el menú del sistema.
 */
function configurarSistemaFormatoChile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.setSpreadsheetLocale("es_CL");
  ss.setSpreadsheetTimeZone("America/Santiago");

  const hojas = ss.getSheets();

  hojas.forEach(hoja => {
    configurarHojaBaseChile_(hoja);
  });

  Logger.log("Configuración chilena aplicada al documento completo.");
}

/**
 * Aplica formato base chileno a una hoja específica.
 *
 * Si la hoja está vacía, crea una estructura mínima visible
 * para que la configuración aplicada pueda verificarse.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja a configurar.
 */
function configurarHojaBaseChile_(hoja) {
  const ESTILO_CHILE = {
    FUENTE: "Calibri",
    TAMANO_FUENTE: 10,
    TAMANO_ENCABEZADO: 10,
    ALTO_ENCABEZADO: 35
  };

  const valores = hoja.getDataRange().getValues();
  const hojaVacia = valores.flat().every(valor => valor === "");

  if (hojaVacia) {
    hoja.getRange(1, 1).setValue("ENCABEZADO");
  }

  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  const rangoCompleto = hoja.getRange(1, 1, ultimaFila, ultimaColumna);
  const encabezado = hoja.getRange(1, 1, 1, ultimaColumna);

  rangoCompleto
    .setFontFamily(ESTILO_CHILE.FUENTE)
    .setFontSize(ESTILO_CHILE.TAMANO_FUENTE)
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  encabezado
    .setFontFamily(ESTILO_CHILE.FUENTE)
    .setFontSize(ESTILO_CHILE.TAMANO_ENCABEZADO)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, ESTILO_CHILE.ALTO_ENCABEZADO);

  hoja.autoResizeColumns(1, ultimaColumna);

  Logger.log(`Hoja configurada: ${hoja.getName()}`);
}

/**
 * Devuelve los formatos estándar usados por el sistema
 * para columnas configuradas bajo criterio chileno.
 *
 * Estos formatos serán utilizados por funciones específicas
 * de columnas: fechas, moneda, porcentajes, números, etc.
 *
 * @return {Object} Objeto con formatos estándar Chile.
 */
function obtenerFormatosChile_() {
  return {
    FECHA: "dd/mm/yyyy",
    FECHA_HORA: "dd/mm/yyyy hh:mm",
    HORA: "hh:mm",
    NUMERO_ENTERO: "#.##0",
    NUMERO_DECIMAL: "#.##0,00",
    MONEDA_CLP: '"$"#.##0',
    PORCENTAJE: "0,00%"
  };
}

/**
 * Devuelve la configuración visual estándar del sistema.
 *
 * Centraliza fuente, tamaño y alto de encabezado para evitar
 * repetir valores dentro de varias funciones de setup.
 *
 * @return {Object} Objeto con estilos visuales estándar.
 */
function obtenerEstiloChile_() {
  return {
    FUENTE: "Calibri",
    TAMANO_FUENTE: 12,
    TAMANO_ENCABEZADO: 14,
    ALTO_ENCABEZADO: 35
  };
}