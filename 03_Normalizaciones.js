/**
 * ============================================================
 * 03_Normalizaciones.gs
 * Proyecto: MATRICULA_2026
 * ------------------------------------------------------------
 * Este archivo reúne funciones reutilizables para normalizar
 * datos ingresados en Google Sheets antes de validarlos.
 *
 * La normalización transforma valores escritos de distintas
 * formas por el operador en formatos consistentes para trabajo,
 * búsqueda, comparación y almacenamiento.
 *
 * Incluye normalización de:
 * - Texto general
 * - RUT chileno
 * - Dígito verificador
 * - RUT completo
 * - Teléfonos chilenos
 * - Correos electrónicos
 * - Fechas
 * ============================================================
 */


/**
 * Normaliza un valor según un tipo declarado.
 *
 * Esta función sirve como punto de entrada general para llamar
 * normalizaciones desde otros archivos, menús, pruebas o eventos.
 *
 * Tipos soportados:
 * - TEXTO
 * - RUT
 * - DV
 * - RUT_COMPLETO
 * - TELEFONO_CHILE
 * - CORREO
 * - FECHA
 *
 * @param {*} valor Valor original ingresado.
 * @param {string} tipo Tipo de normalización solicitada.
 * @return {*} Valor normalizado.
 */
function normalizarValorPorTipo(valor, tipo) {
  const tipoNormalizado = String(tipo || "").trim().toUpperCase();

  switch (tipoNormalizado) {
    case "TEXTO":
      return normalizarTextoMayusculas_(valor);

    case "RUT":
      return normalizarRutCuerpo_(valor);

    case "DV":
      return normalizarDvRut_(valor);

    case "RUT_COMPLETO":
      return normalizarRutCompleto_(valor);

    case "TELEFONO_CHILE":
      return normalizarTelefonoChile_(valor);

    case "CORREO":
      return normalizarCorreo_(valor);

    case "FECHA":
      return normalizarFecha_(valor);

    default:
      return valor;
  }
}


/**
 * Normaliza texto general.
 *
 * Elimina espacios sobrantes, convierte múltiples espacios internos
 * en uno solo y transforma el texto a mayúsculas.
 *
 * @param {*} valor Texto original.
 * @return {string} Texto normalizado.
 */
function normalizarTextoMayusculas_(valor) {
  if (valor === null || valor === undefined) return "";

  return String(valor)
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}


/**
 * Normaliza el cuerpo de un RUT chileno.
 *
 * Elimina puntos, guiones, espacios y cualquier carácter no numérico.
 * Devuelve solo el cuerpo numérico del RUT, sin dígito verificador.
 *
 * Ejemplo:
 * "23.761.825-4" -> "23761825"
 *
 * @param {*} valor RUT o cuerpo de RUT.
 * @return {string} Cuerpo del RUT normalizado.
 */
function normalizarRutCuerpo_(valor) {
  if (valor === null || valor === undefined) return "";

  let texto = String(valor).trim().toUpperCase();

  if (texto.includes("-")) {
    texto = texto.split("-")[0];
  }

  return texto.replace(/\D/g, "");
}


/**
 * Normaliza el dígito verificador de un RUT chileno.
 *
 * Elimina espacios, puntos y guiones. Convierte K minúscula a K
 * mayúscula. Si recibe un RUT completo con guion, toma lo que está
 * después del guion.
 *
 * Ejemplo:
 * "23.761.825-k" -> "K"
 *
 * @param {*} valor Dígito verificador o RUT completo.
 * @return {string} DV normalizado.
 */
function normalizarDvRut_(valor) {
  if (valor === null || valor === undefined) return "";

  let texto = String(valor).trim().toUpperCase();

  if (texto.includes("-")) {
    const partes = texto.split("-");
    texto = partes[partes.length - 1];
  }

  texto = texto.replace(/[.\-\s]/g, "");

  if (texto.length === 0) return "";

  return texto.slice(-1);
}


/**
 * Normaliza un RUT completo chileno.
 *
 * Elimina puntos y espacios, separa cuerpo y DV, y devuelve el formato:
 * cuerpo-DV
 *
 * No valida matemáticamente el RUT. Solo normaliza estructura.
 *
 * Ejemplo:
 * "23.761.825-k" -> "23761825-K"
 *
 * @param {*} valor RUT completo.
 * @return {string} RUT normalizado.
 */
function normalizarRutCompleto_(valor) {
  if (valor === null || valor === undefined) return "";

  let texto = String(valor)
    .trim()
    .toUpperCase()
    .replace(/\./g, "")
    .replace(/\s+/g, "");

  if (texto === "") return "";

  let cuerpo = "";
  let dv = "";

  if (texto.includes("-")) {
    const partes = texto.split("-");
    cuerpo = partes[0].replace(/\D/g, "");
    dv = partes[1] ? partes[1].replace(/[^0-9K]/g, "") : "";
  } else {
    texto = texto.replace(/[^0-9K]/g, "");

    if (texto.length < 2) return texto;

    cuerpo = texto.slice(0, -1).replace(/\D/g, "");
    dv = texto.slice(-1);
  }

  if (!cuerpo || !dv) return "";

  return cuerpo + "-" + dv;
}

/**
 * Normaliza teléfonos chilenos.
 *
 * Deja el teléfono en formato interno limpio:
 * 56XXXXXXXXX
 *
 * No agrega espacios ni formato visual.
 *
 * @param {*} valor Teléfono original.
 * @return {string} Teléfono normalizado o vacío si no corresponde.
 */
function normalizarTelefonoChile_(valor) {
  if (valor === null || valor === undefined || valor === "") return "";

  let telefono = String(valor).replace(/\D/g, "");

  if (telefono.startsWith("0056")) {
    telefono = telefono.substring(2);
  }

  if (!telefono.startsWith("56") && telefono.length === 9) {
    telefono = "56" + telefono;
  }

  if (telefono.startsWith("56") && telefono.length === 11) {
    return telefono;
  }

  return "";
}


/**
 * Normaliza correos electrónicos.
 *
 * Elimina espacios al inicio y al final, convierte a minúsculas
 * y remueve espacios internos accidentales.
 *
 * No valida si el correo existe. Solo normaliza escritura.
 *
 * @param {*} valor Correo original.
 * @return {string} Correo normalizado.
 */
function normalizarCorreo_(valor) {
  if (valor === null || valor === undefined) return "";

  return String(valor)
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}


/**
 * Normaliza fechas.
 *
 * Si recibe una fecha válida, la devuelve como objeto Date.
 * Si recibe texto interpretable como fecha, intenta convertirlo.
 * Si no puede convertirlo, devuelve vacío.
 *
 * Pensada para que Google Sheets maneje visualmente el formato
 * de fecha mediante formato de celda, no mediante texto fijo.
 *
 * @param {*} valor Fecha original.
 * @return {Date|string} Fecha normalizada o vacío.
 */
function normalizarFecha_(valor) {
  if (valor === null || valor === undefined || valor === "") return "";

  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor)) {
    return valor;
  }

  const texto = String(valor).trim();

  if (texto === "") return "";

  const fecha = new Date(texto);

  if (!isNaN(fecha.getTime())) {
    return fecha;
  }

  return "";
}
