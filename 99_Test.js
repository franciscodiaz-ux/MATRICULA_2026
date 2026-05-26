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
    .addItem("Crear nuevo hoja Test", "recrearHojaTest")
    .addItem("Configurar hoja Test", "setupHojaTest")
    .addToUi();
}

/**
 * Recrea completamente la hoja "Test".
 *
 * Si la hoja existe, solicita confirmación al usuario y luego la elimina.
 * Después crea una nueva hoja vacía llamada "Test".
 *
 * Finalmente aplica la configuración base de la hoja de pruebas.
 */
function recrearHojaTest() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const respuesta = ui.alert(
    "Recrear hoja Test",
    "La hoja Test será eliminada y creada nuevamente desde cero.\n\n¿Desea continuar?",
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert("Operación cancelada.");
    return;
  }

  const hojaExistente = ss.getSheetByName("Test");

  if (hojaExistente) {
    ss.deleteSheet(hojaExistente);
  }

  const hojaNueva = ss.insertSheet("Test");
  ss.setActiveSheet(hojaNueva);
  SpreadsheetApp.flush();

  setupHojaTest();
  ui.alert('La hoja "Test" fue recreada y configurada correctamente.'); 
}


/**
 * Configura la hoja "Test" para pruebas controladas del proyecto.
 *
 * Aplica la configuración general del documento, obtiene la hoja Test,
 * aplica el formato base definido para hojas del sistema y ejecuta la
 * autodetección de columnas según los encabezados existentes.
 *
 * Esta función permite probar configuraciones reales sin afectar hojas
 * productivas como Alumnos.
 */
function setupHojaTest() {
  configurarDocumentoFormatoChile();

  const nombreHoja = "Test";
  const hoja = obtenerHojaObligatoria_(nombreHoja);

  configurarHojaBaseChile_(hoja);
  configurarColumnasAutomaticasPorEncabezado(nombreHoja);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Setup completado para hoja: ${nombreHoja}`,
    "Test",
    5
  );

  Logger.log(`Setup completado para hoja: ${nombreHoja}`);
}

/**
 * Crea encabezados estándar en la hoja "Test" y aplica formato automático.
 *
 * Esta función escribe una fila de encabezados representativos para probar
 * la autodetección de columnas por nombre. Luego aplica la configuración
 * base de la hoja y el formato automático según encabezados.
 */
function crearEncabezadosHojaTest() {
  const nombreHoja = "Test";
  const hoja = obtenerHojaObligatoria_(nombreHoja);

  const encabezados = [
    "RUN_ALM",
    "DV_ALM",
    "FNAC_ALM",
    "PATERNO_ALM",
    "MATERNO_ALM",
    "NOMBRES_ALM",
    "CEL_ALM",
    "CORREO_ALM"
  ];

  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);

  configurarHojaBaseChile_(hoja);
  configurarColumnasAutomaticasPorEncabezado(nombreHoja);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Encabezados creados y formato aplicado.",
    "Test",
    5
  );

  Logger.log("Encabezados creados y formato automático aplicado en hoja Test.");
}

/**
 * Genera datos de prueba en la hoja "Test".
 *
 * Inserta registros con valores correctos e incorrectos para probar
 * posteriormente normalización, validación y formato.
 */
function generarDatosPruebaTest() {
  const hoja = obtenerHojaObligatoria_("Test");

  const datos = [
    ["13272823", "k", "12/03/2010", "diaz", "perez", "juan carlos", "912345678", "juan@test.cl"],
    ["23761825", "5", "01-01-2009", "gonzalez", "soto", "maria jose", "+56987654321", "maria.correo.cl"],
    ["49132263", "1", "2011/05/22", "muñoz", "rojas", "pedro andres", "87654321", "pedro@test.cl"],
    ["12345678", "9", "31/02/2010", "silva", "vera", "ana", "56912345678", "ana@test"],
    ["", "", "", "", "", "", "", ""]
  ];

  hoja.getRange(2, 1, datos.length, datos[0].length).setValues(datos);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Datos de prueba generados en hoja Test.",
    "Test",
    5
  );
}


/**
 * Normaliza los datos existentes en la hoja "Test".
 *
 * Recorre las filas con datos y aplica normalización básica según el
 * encabezado de cada columna. Usa funciones de formato si existen en
 * 02_Formatos.gs.
 */
function normalizarDatosHojaTest() {
  const hoja = obtenerHojaObligatoria_("Test");
  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaFila < 2) {
    SpreadsheetApp.getUi().alert("No hay datos para normalizar.");
    return;
  }

  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];
  const rangoDatos = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna);
  const datos = rangoDatos.getValues();

  const datosNormalizados = datos.map(fila => {
    return fila.map((valor, indice) => {
      const encabezado = String(encabezados[indice]).toUpperCase();

      if (encabezado.includes("RUN") || encabezado.includes("RUT")) {
        return normalizarRunTest_(valor);
      }

      if (encabezado.includes("DV")) {
        return normalizarDvTest_(valor);
      }

      if (encabezado.includes("CEL") || encabezado.includes("FONO") || encabezado.includes("TELEF")) {
        return normalizarTelefonoTest_(valor);
      }

      if (encabezado.includes("CORREO") || encabezado.includes("EMAIL")) {
        return normalizarCorreoTest_(valor);
      }

      if (
        encabezado.includes("NOMBRE") ||
        encabezado.includes("PATERNO") ||
        encabezado.includes("MATERNO")
      ) {
        return normalizarTextoNombreTest_(valor);
      }

      return valor;
    });
  });

  rangoDatos.setValues(datosNormalizados);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Datos normalizados en hoja Test.",
    "Test",
    5
  );
}


/**
 * Diagnostica los encabezados de la hoja "Test".
 *
 * Revisa la fila 1 y muestra qué tipo de dato parece representar cada
 * columna según su nombre. Sirve para comprobar la autodetección antes
 * de aplicar procesos reales.
 */
function diagnosticarEncabezadosTest() {
  const hoja = obtenerHojaObligatoria_("Test");
  const ultimaColumna = hoja.getLastColumn();

  if (ultimaColumna === 0) {
    SpreadsheetApp.getUi().alert("La hoja Test no tiene encabezados.");
    return;
  }

  const encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];

  const diagnostico = encabezados.map((encabezado, indice) => {
    return `Columna ${indice + 1}: ${encabezado} → ${detectarTipoEncabezadoTest_(encabezado)}`;
  });

  SpreadsheetApp.getUi().alert(
    "Diagnóstico de encabezados",
    diagnostico.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Detecta el tipo probable de dato a partir del encabezado recibido.
 */
function detectarTipoEncabezadoTest_(encabezado) {
  const texto = String(encabezado).toUpperCase();

  if (texto.includes("RUN") || texto.includes("RUT")) return "RUN/RUT";
  if (texto.includes("DV")) return "DV";
  if (texto.includes("FNAC") || texto.includes("FECHA")) return "FECHA";
  if (texto.includes("CEL") || texto.includes("FONO") || texto.includes("TELEF")) return "TELÉFONO";
  if (texto.includes("CORREO") || texto.includes("EMAIL")) return "CORREO";
  if (texto.includes("NOMBRE") || texto.includes("PATERNO") || texto.includes("MATERNO")) return "TEXTO/NOMBRE";

  return "NO DETECTADO";
}


/**
 * Normaliza RUN/RUT dejando solo números, sin puntos ni guion.
 */
function normalizarRunTest_(valor) {
  return String(valor || "").replace(/\D/g, "");
}


/**
 * Normaliza dígito verificador chileno.
 */
function normalizarDvTest_(valor) {
  return String(valor || "").trim().toUpperCase();
}


/**
 * Normaliza teléfono chileno a una forma básica de texto.
 */
function normalizarTelefonoTest_(valor) {
  let texto = String(valor || "").replace(/\D/g, "");

  if (texto.length === 8) {
    texto = "569" + texto;
  }

  if (texto.length === 9 && texto.startsWith("9")) {
    texto = "56" + texto;
  }

  return texto;
}


/**
 * Normaliza correo electrónico en minúsculas y sin espacios.
 */
function normalizarCorreoTest_(valor) {
  return String(valor || "").trim().toLowerCase();
}


/**
 * Normaliza nombres y apellidos en mayúsculas limpias.
 */
function normalizarTextoNombreTest_(valor) {
  return String(valor || "").trim().toUpperCase();
}
