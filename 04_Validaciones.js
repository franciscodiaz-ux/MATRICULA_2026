/**
 * ============================================================
 * 04_Validaciones.gs
 * Proyecto: MATRICULA_2026
 * ------------------------------------------------------------
 * Este archivo contiene funciones reutilizables para validar
 * datos ya normalizados o ingresados por el operador.
 *
 * La validación no debe formatear visualmente ni modificar la hoja.
 * Solo debe responder si un dato cumple o no una regla definida.
 *
 * Incluye validación de:
 * - RUT chileno
 * - Teléfono chileno normalizado
 * - Correo electrónico básico
 * ============================================================
 */


/**
 * Calcula el dígito verificador de un RUT chileno.
 *
 * Recibe el cuerpo del RUT, sin puntos ni guion, y retorna el
 * dígito verificador correspondiente usando módulo 11.
 *
 * @param {string|number} rut Cuerpo del RUT.
 * @return {string} Dígito verificador calculado.
 */
function calcularDV(rut) {
  rut = String(rut || "").replace(/\D/g, "");

  if (rut === "") return "";

  let suma = 0;
  let multiplicador = 2;

  for (let i = rut.length - 1; i >= 0; i--) {
    suma += parseInt(rut[i], 10) * multiplicador;
    multiplicador++;

    if (multiplicador > 7) {
      multiplicador = 2;
    }
  }

  const resto = 11 - (suma % 11);

  if (resto === 11) return "0";
  if (resto === 10) return "K";

  return String(resto);
}


/**
 * Valida un RUT chileno completo.
 *
 * Acepta formatos como:
 * - 23761825-K
 * - 23.761.825-K
 * - 23761825K
 *
 * Usa normalizarRutCompleto_() desde 03_Normalizaciones.gs.
 *
 * @param {*} valor RUT completo ingresado.
 * @return {boolean} true si el RUT es válido.
 */
function validarRutCompletoChile(valor) {
  const rutNormalizado = normalizarRutCompleto_(valor);

  if (rutNormalizado === "") return false;
  if (!rutNormalizado.includes("-")) return false;

  const partes = rutNormalizado.split("-");
  const cuerpo = partes[0];
  const dv = partes[1];

  if (!/^\d{7,8}$/.test(cuerpo)) return false;
  if (!/^[0-9K]$/.test(dv)) return false;

  return calcularDV(cuerpo) === dv;
}


/**
 * Valida un teléfono chileno según el criterio interno del proyecto.
 *
 * El teléfono válido debe poder normalizarse como:
 * 56XXXXXXXXX
 *
 * No intenta determinar si es celular, Santiago o región.
 *
 * Usa normalizarTelefonoChile_() desde 03_Normalizaciones.gs.
 *
 * @param {*} valor Teléfono ingresado.
 * @return {boolean} true si el teléfono queda válido.
 */
function validarTelefonoChile(valor) {
  const telefono = normalizarTelefonoChile_(valor);

  return /^56\d{9}$/.test(telefono);
}


/**
 * Valida un correo electrónico con una regla básica.
 *
 * Usa normalizarCorreo_() desde 03_Normalizaciones.gs.
 * No verifica si el correo existe realmente.
 *
 * @param {*} valor Correo ingresado.
 * @return {boolean} true si tiene estructura básica válida.
 */
function validarCorreo(valor) {
  const correo = normalizarCorreo_(valor);

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo);
}