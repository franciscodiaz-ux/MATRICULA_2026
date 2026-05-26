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
 * 3. Autodetección de columnas por encabezado
 * ------------------------------------------------------------
 * Funciones que detectan tipos de columnas según el nombre
 * del encabezado.
 *
 * La autodetección solo debe ejecutarse dentro del setup
 * explícito de una hoja. No debe recorrer todas las hojas.
 *
 *
 * 4. Setup por hoja
 * ------------------------------------------------------------
 * Cada hoja del sistema tendrá su propia función de setup.
 *
 * Estas funciones:
 *
 * - aplican el formato base de hoja;
 * - llaman únicamente a las configuraciones necesarias
 *   para esa hoja;
 * - pueden usar autodetección de encabezados;
 * - jamás deben modificar hojas no definidas explícitamente.
 *
 * Ejemplo:
 *
 * setupHojaAlumnos()
 *    ├── configurarDocumentoFormatoChile()
 *    ├── configurarHojaBaseChile_()
 *    ├── configurarColumnasAutomaticasPorEncabezado()
 *    └── configuraciones manuales especiales si se requieren
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
 * - Los teléfonos y celulares solo se configuran como texto
 *   en este archivo. Su limpieza o normalización corresponde
 *   a 02_Formatos.gs.
 *
 * - No utilizar SpreadsheetApp.getUi() en funciones internas
 *   o reutilizables.
 *
 * - Toda nueva función debe incluir encabezado descriptivo.
 *
 */


// ============================================================
// 1. CONFIGURACIÓN GENERAL CHILE
// ============================================================

/**
 * Configura el documento completo bajo criterios generales chilenos.
 *
 * Aplica configuración regional chilena y zona horaria de Santiago.
 * Esta función afecta al documento completo, pero no modifica hojas,
 * columnas ni datos.
 */
function configurarDocumentoFormatoChile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.setSpreadsheetLocale("es_CL");
  ss.setSpreadsheetTimeZone("America/Santiago");

  Logger.log("Documento configurado con locale es_CL y timezone America/Santiago.");
}

/**
 * Devuelve los formatos estándar usados por el sistema bajo criterios chilenos.
 *
 * Estos formatos son usados por las funciones de configuración de columnas.
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
 * Centraliza fuente, tamaño y alto de encabezado para mantener
 * consistencia visual en las hojas configuradas.
 *
 * @return {Object} Objeto con estilos visuales estándar.
 */
function obtenerEstiloChile_() {
  return {
    FUENTE: "Calibri",
    TAMANO_FUENTE: 10,
    TAMANO_ENCABEZADO: 10,
    ALTO_ENCABEZADO: 35
  };
}

/**
 * Aplica formato visual base a una hoja específica.
 *
 * Esta función solo modifica la hoja recibida como parámetro.
 * Si la hoja está vacía, crea un encabezado mínimo para que la
 * configuración aplicada sea visible y verificable.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja a configurar.
 */
function configurarHojaBaseChile_(hoja) {
  const estilo = obtenerEstiloChile_();

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
    .setFontFamily(estilo.FUENTE)
    .setFontSize(estilo.TAMANO_FUENTE)
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  encabezado
    .setFontFamily(estilo.FUENTE)
    .setFontSize(estilo.TAMANO_ENCABEZADO)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, estilo.ALTO_ENCABEZADO);
  hoja.autoResizeColumns(1, ultimaColumna);

  Logger.log(`Formato base aplicado a la hoja: ${hoja.getName()}`);
}

/**
 * Obtiene una hoja por nombre y valida que exista.
 *
 * Se usa para asegurar que ninguna función trabaje sobre una hoja
 * inexistente o asumida accidentalmente.
 *
 * @param {string} nombreHoja Nombre exacto de la hoja.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Hoja encontrada.
 */
function obtenerHojaObligatoria_(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);

  if (!hoja) {
    throw new Error(`No existe la hoja "${nombreHoja}".`);
  }

  return hoja;
}


// ============================================================
// 2. CONFIGURACIÓN POR COLUMNAS
// ============================================================

/**
 * Busca el número de columna según el nombre exacto del encabezado.
 *
 * Lee la fila 1 de la hoja y devuelve la posición de la columna
 * cuyo encabezado coincide con el nombre indicado.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja donde buscar.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 * @return {number} Número de columna encontrado.
 */
function obtenerNumeroColumnaPorEncabezado_(hoja, nombreColumna) {
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaColumna < 1) {
    throw new Error(`La hoja "${hoja.getName()}" no tiene columnas.`);
  }

  const encabezados = hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getValues()[0]
    .map(valor => String(valor).trim());

  const indice = encabezados.indexOf(nombreColumna);

  if (indice === -1) {
    throw new Error(
      `No se encontró la columna "${nombreColumna}" en la hoja "${hoja.getName()}".`
    );
  }

  return indice + 1;
}

/**
 * Configura una columna recibida por número como texto.
 *
 * Se usa internamente para RUT, DV, teléfonos, celulares, correos
 * y otros datos que no deben ser tratados como números.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja a configurar.
 * @param {number} columna Número de columna.
 */
function configurarColumnaTextoPorNumero_(hoja, columna) {
  const ultimaFila = Math.max(hoja.getMaxRows(), 2);

  hoja
    .getRange(2, columna, ultimaFila - 1, 1)
    .setNumberFormat("@")
    .setHorizontalAlignment("left");
}

/**
 * Configura una columna recibida por número como fecha chilena.
 *
 * Aplica formato dd/mm/yyyy y centra el contenido.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja a configurar.
 * @param {number} columna Número de columna.
 */
function configurarColumnaFechaPorNumero_(hoja, columna) {
  const formatos = obtenerFormatosChile_();
  const ultimaFila = Math.max(hoja.getMaxRows(), 2);

  hoja
    .getRange(2, columna, ultimaFila - 1, 1)
    .setNumberFormat(formatos.FECHA)
    .setHorizontalAlignment("center");
}

/**
 * Configura una columna recibida por número como moneda chilena.
 *
 * Aplica formato pesos chilenos sin decimales y alinea a la derecha.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja a configurar.
 * @param {number} columna Número de columna.
 */
function configurarColumnaMonedaPorNumero_(hoja, columna) {
  const formatos = obtenerFormatosChile_();
  const ultimaFila = Math.max(hoja.getMaxRows(), 2);

  hoja
    .getRange(2, columna, ultimaFila - 1, 1)
    .setNumberFormat(formatos.MONEDA_CLP)
    .setHorizontalAlignment("right");
}

/**
 * Configura una columna como texto, buscándola por encabezado.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaTexto(nombreHoja, nombreColumna) {
  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const columna = obtenerNumeroColumnaPorEncabezado_(hoja, nombreColumna);

  configurarColumnaTextoPorNumero_(hoja, columna);

  Logger.log(`Columna texto configurada: ${nombreHoja}.${nombreColumna}`);
}

/**
 * Configura una columna como RUT chileno.
 *
 * El cuerpo del RUT se mantiene como texto para evitar cambios
 * automáticos de Google Sheets.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaRut(nombreHoja, nombreColumna) {
  configurarColumnaTexto(nombreHoja, nombreColumna);
}

/**
 * Configura una columna como dígito verificador chileno.
 *
 * Se mantiene como texto porque puede contener números o K.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaDV(nombreHoja, nombreColumna) {
  configurarColumnaTexto(nombreHoja, nombreColumna);
}

/**
 * Configura una columna como teléfono chileno.
 *
 * Se mantiene como texto para evitar pérdida de dígitos.
 * La limpieza y normalización del número corresponde a 02_Formatos.gs.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaTelefonoChile(nombreHoja, nombreColumna) {
  configurarColumnaTexto(nombreHoja, nombreColumna);
}

/**
 * Configura una columna como celular chileno.
 *
 * Se mantiene como texto para evitar pérdida de dígitos.
 * La limpieza y normalización del número corresponde a 02_Formatos.gs.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaCelularChile(nombreHoja, nombreColumna) {
  configurarColumnaTexto(nombreHoja, nombreColumna);
}

/**
 * Configura una columna como fecha chilena.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaFechaChile(nombreHoja, nombreColumna) {
  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const columna = obtenerNumeroColumnaPorEncabezado_(hoja, nombreColumna);

  configurarColumnaFechaPorNumero_(hoja, columna);

  Logger.log(`Columna fecha configurada: ${nombreHoja}.${nombreColumna}`);
}

/**
 * Configura una columna como moneda chilena sin decimales.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaMonedaCLP(nombreHoja, nombreColumna) {
  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const columna = obtenerNumeroColumnaPorEncabezado_(hoja, nombreColumna);

  configurarColumnaMonedaPorNumero_(hoja, columna);

  Logger.log(`Columna moneda CLP configurada: ${nombreHoja}.${nombreColumna}`);
}

/**
 * Configura una columna como correo electrónico.
 *
 * Se mantiene como texto.
 *
 * @param {string} nombreHoja Nombre de la hoja.
 * @param {string} nombreColumna Nombre exacto del encabezado.
 */
function configurarColumnaCorreo(nombreHoja, nombreColumna) {
  configurarColumnaTexto(nombreHoja, nombreColumna);
}

// ============================================================
// 3. AUTODETECCIÓN DE COLUMNAS POR ENCABEZADO
// ============================================================

/**
 * Normaliza un encabezado para facilitar su comparación.
 *
 * Convierte el texto a mayúsculas, elimina espacios extremos
 * y reemplaza espacios internos por guion bajo.
 *
 * @param {string} encabezado Encabezado original.
 * @return {string} Encabezado normalizado.
 */
function normalizarEncabezado_(encabezado) {
  return String(encabezado)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "_");
}

/**
 * Determina si un encabezado corresponde a una columna RUT.
 *
 * Detecta encabezados como:
 * - RUT
 * - RUT_AL
 * - RUT_APT
 * - RUT_APS
 * - RUT_PADRE
 * - RUT_MADRE
 * - cualquier encabezado que comience con RUT_
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a RUT.
 */
function esEncabezadoRut_(encabezado) {
  return encabezado === "RUT" || encabezado.startsWith("RUT_");
}

/**
 * Determina si un encabezado corresponde a una columna DV.
 *
 * Detecta encabezados como:
 * - DV
 * - DV_AL
 * - DV_APT
 * - DV_APS
 * - DV_PADRE
 * - DV_MADRE
 * - cualquier encabezado que comience con DV_
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a DV.
 */
function esEncabezadoDV_(encabezado) {
  return encabezado === "DV" || encabezado.startsWith("DV_");
}

/**
 * Determina si un encabezado corresponde a una columna de teléfono.
 *
 * Detecta encabezados que comiencen con:
 * - FONO
 * - TELEFONO
 * - TEL
 *
 * Nota:
 * En 00_Setup.gs solo se configura como texto.
 * La limpieza del número queda para 02_Formatos.gs.
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a teléfono.
 */
function esEncabezadoTelefono_(encabezado) {
  return (
    encabezado.startsWith("FONO") ||
    encabezado.startsWith("TELEFONO") ||
    encabezado.startsWith("TEL")
  );
}

/**
 * Determina si un encabezado corresponde a una columna de celular.
 *
 * Detecta encabezados que comiencen con:
 * - CEL
 * - CELULAR
 * - MOVIL
 *
 * Nota:
 * En 00_Setup.gs solo se configura como texto.
 * La limpieza del número queda para 02_Formatos.gs.
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a celular.
 */
function esEncabezadoCelular_(encabezado) {
  return (
    encabezado.startsWith("CEL") ||
    encabezado.startsWith("CELULAR") ||
    encabezado.startsWith("MOVIL")
  );
}

/**
 * Determina si un encabezado corresponde a una columna de correo.
 *
 * Detecta encabezados que comiencen con:
 * - CORREO
 * - EMAIL
 * - MAIL
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a correo.
 */
function esEncabezadoCorreo_(encabezado) {
  return (
    encabezado.startsWith("CORREO") ||
    encabezado.startsWith("EMAIL") ||
    encabezado.startsWith("MAIL")
  );
}

/**
 * Determina si un encabezado corresponde a una columna de fecha.
 *
 * Detecta encabezados exactos o prefijos frecuentes en matrícula:
 * - FECHA
 * - FNAC
 * - FMAT
 * - FRET
 * - FEC_
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a fecha.
 */
function esEncabezadoFecha_(encabezado) {
  return (
    encabezado === "FECHA" ||
    encabezado === "FNAC" ||
    encabezado === "FMAT" ||
    encabezado === "FRET" ||
    encabezado.startsWith("FECHA_") ||
    encabezado.startsWith("FEC_")
  );
}

/**
 * Determina si un encabezado corresponde a moneda chilena.
 *
 * Detecta encabezados asociados a valores monetarios:
 * - MONTO
 * - VALOR
 * - PRECIO
 * - TOTAL
 * - CLP
 *
 * @param {string} encabezado Encabezado normalizado.
 * @return {boolean} Verdadero si corresponde a moneda.
 */
function esEncabezadoMoneda_(encabezado) {
  return (
    encabezado.startsWith("MONTO") ||
    encabezado.startsWith("VALOR") ||
    encabezado.startsWith("PRECIO") ||
    encabezado.startsWith("TOTAL") ||
    encabezado.includes("CLP")
  );
}

/**
 * Configura automáticamente una columna según su encabezado.
 *
 * Esta función recibe una hoja, un encabezado ya normalizado y el número
 * de columna correspondiente. Evalúa el encabezado contra los detectores
 * definidos en esta sección y aplica el formato adecuado.
 *
 * Devuelve true si aplicó alguna configuración, o false si no reconoció
 * el encabezado.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja a configurar.
 * @param {string} nombreHoja Nombre exacto de la hoja.
 * @param {string} encabezado Encabezado normalizado.
 * @param {number} columna Número de columna.
 * @return {boolean} Verdadero si configuró la columna.
 */
function configurarColumnaDetectadaPorEncabezado_(hoja, nombreHoja, encabezado, columna) {
  if (esEncabezadoRut_(encabezado)) {
    configurarColumnaTextoPorNumero_(hoja, columna);
    Logger.log(`Autodetectado RUT: ${nombreHoja}.${encabezado}`);
    return true;
  }

  if (esEncabezadoDV_(encabezado)) {
    configurarColumnaTextoPorNumero_(hoja, columna);
    Logger.log(`Autodetectado DV: ${nombreHoja}.${encabezado}`);
    return true;
  }

  if (esEncabezadoCelular_(encabezado)) {
    configurarColumnaTextoPorNumero_(hoja, columna);
    Logger.log(`Autodetectado CELULAR: ${nombreHoja}.${encabezado}`);
    return true;
  }

  if (esEncabezadoTelefono_(encabezado)) {
    configurarColumnaTextoPorNumero_(hoja, columna);
    Logger.log(`Autodetectado TELEFONO: ${nombreHoja}.${encabezado}`);
    return true;
  }

  if (esEncabezadoCorreo_(encabezado)) {
    configurarColumnaTextoPorNumero_(hoja, columna);
    Logger.log(`Autodetectado CORREO: ${nombreHoja}.${encabezado}`);
    return true;
  }

  if (esEncabezadoFecha_(encabezado)) {
    configurarColumnaFechaPorNumero_(hoja, columna);
    Logger.log(`Autodetectado FECHA: ${nombreHoja}.${encabezado}`);
    return true;
  }

  if (esEncabezadoMoneda_(encabezado)) {
    configurarColumnaMonedaPorNumero_(hoja, columna);
    Logger.log(`Autodetectado MONEDA: ${nombreHoja}.${encabezado}`);
    return true;
  }

  return false;
}

/**
 * Configura automáticamente columnas según el nombre de sus encabezados.
 *
 * Esta función recorre solo las columnas de la hoja indicada explícitamente.
 * No recorre todas las hojas del documento.
 *
 * La función detecta tipos comunes:
 * - RUT
 * - DV
 * - teléfonos
 * - celulares
 * - correos
 * - fechas
 * - moneda CLP
 *
 * Debe llamarse únicamente desde una función setupHojaX().
 *
 * @param {string} nombreHoja Nombre exacto de la hoja a configurar.
 */
function configurarColumnasAutomaticasPorEncabezado(nombreHoja) {
  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaColumna < 1) {
    Logger.log(`La hoja "${nombreHoja}" no tiene columnas para detectar.`);
    return;
  }

  const encabezados = hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getValues()[0];

  encabezados.forEach((valorEncabezado, indice) => {
    const encabezado = normalizarEncabezado_(valorEncabezado);
    const columna = indice + 1;

    if (!encabezado) return;

    configurarColumnaDetectadaPorEncabezado_(
      hoja,
      nombreHoja,
      encabezado,
      columna
    );
  });

  Logger.log(`Autodetección de columnas completada para hoja: ${nombreHoja}`);
}


// ============================================================
// 4. SETUP POR HOJA---ver archivos con nombre de hoja
// ============================================================






