/**
 * 05_Comportamientos.js
 *
 * Coordina los comportamientos automáticos del sistema ante ediciones
 * realizadas por el usuario en la hoja "Alumnos".
 *
 * Este archivo detecta la hoja, fila y encabezado editado, y luego deriva
 * la acción a funciones específicas. La lógica se basa en nombres de
 * encabezados, no en números ni letras de columna.
 *
 * Implementación actual:
 * - Normalización de teléfonos de apoderados.
 * - Normalización de correos de alumno y apoderados.
 * - Marcado visual de alumnos retirados al editar FRET.
 */


/**
 * Se ejecuta automáticamente cuando el usuario edita una celda del documento.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Evento de edición.
 */
function onEdit(e) {
  if (!e) return;

  const rango = e.range;
  const hoja = rango.getSheet();
  const nombreHoja = hoja.getName();

  switch (nombreHoja) {
    case "Alumnos":
      manejarEdicionAlumnos_(e);
      break;

    default:
      return;
  }
}


/**
 * Maneja las ediciones realizadas en la hoja "Alumnos".
 *
 * Detecta la fila y columna editada, ignora la fila de encabezados,
 * obtiene el encabezado de la columna y ejecuta el comportamiento
 * correspondiente.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Evento de edición.
 */
function manejarEdicionAlumnos_(e) {
  const rango = e.range;
  const hoja = rango.getSheet();
  const fila = rango.getRow();
  const columna = rango.getColumn();

  if (fila === 1) return;

  const encabezado = String(hoja.getRange(1, columna).getValue()).trim();

  switch (encabezado) {
    case "FONO1APT":
    case "FONO2APT":
    case "FONO1APS":
    case "FONO2APS":
      comportamientoNormalizarTelefono_(rango);
      break;

    case "CORREOAL":
    case "CORREOAPT":
    case "CORREOAPS":
      comportamientoNormalizarCorreo_(rango);
      break;

    case "FRET":
      comportamientoFechaRetiro_(e);
      break;

    default:
      return;
  }
}


/**
 * Aplica o quita el formato visual de alumno retirado según la columna FRET.
 *
 * Si FRET contiene una fecha válida, marca toda la fila con letra roja
 * y texto tachado. Si FRET queda vacía, restaura toda la fila con letra
 * negra y sin tachado.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Evento de edición.
 */
function comportamientoFechaRetiro_(e) {
  const rango = e.range;
  const hoja = rango.getSheet();
  const fila = rango.getRow();
  const valor = rango.getValue();

  const ultimaColumna = hoja.getLastColumn();
  const rangoFila = hoja.getRange(fila, 1, 1, ultimaColumna);

  const esFechaValida =
    Object.prototype.toString.call(valor) === "[object Date]" &&
    !isNaN(valor.getTime());

  if (esFechaValida) {
    rangoFila
      .setFontColor("red")
      .setFontLine("line-through");
    return;
  }

  if (valor === "") {
    rangoFila
      .setFontColor("black")
      .setFontLine("none");
  }
}


/**
 * Normaliza el valor telefónico escrito en una celda.
 *
 * @param {GoogleAppsScript.Range} rango Celda editada.
 */
function comportamientoNormalizarTelefono_(rango) {
  const valorOriginal = rango.getValue();

  if (valorOriginal === "" || valorOriginal === null) return;

  const telefonoNormalizado = normalizarTelefonoChile_(valorOriginal);

  if (!telefonoNormalizado) return;

  rango.setValue(telefonoNormalizado);
  rango.setHorizontalAlignment("right");
}


/**
 * Normaliza un número telefónico chileno a formato visual estándar.
 *
 * @param {*} valor Valor ingresado por el usuario.
 * @return {string|null} Teléfono normalizado o null si no se pudo normalizar.
 */
function normalizarTelefonoChile_(valor) {
  let digitos = String(valor).replace(/\D/g, "");

  if (!digitos) return null;

  if (digitos.length === 9) {
    digitos = "56" + digitos;
  }

  if (digitos.length !== 11) return null;
  if (!digitos.startsWith("56")) return null;

  const pais = digitos.substring(0, 2);
  const tipo = digitos.substring(2, 3);
  const bloque1 = digitos.substring(3, 7);
  const bloque2 = digitos.substring(7, 11);

  return `${pais} ${tipo} ${bloque1} ${bloque2}`;
}


/**
 * Normaliza el correo escrito en una celda.
 *
 * @param {GoogleAppsScript.Range} rango Celda editada.
 */
function comportamientoNormalizarCorreo_(rango) {
  const valorOriginal = rango.getValue();

  if (valorOriginal === "" || valorOriginal === null) return;

  const correoNormalizado = normalizarCorreo_(valorOriginal);

  if (!correoNormalizado) return;

  rango.setValue(correoNormalizado);
}


/**
 * Normaliza una dirección de correo electrónico.
 *
 * @param {*} valor Valor ingresado por el usuario.
 * @return {string|null} Correo normalizado o null si queda vacío.
 */
function normalizarCorreo_(valor) {
  const correo = String(valor)
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();

  return correo || null;
}