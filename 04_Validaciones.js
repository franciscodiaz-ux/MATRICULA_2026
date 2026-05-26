/**
 * Archivo: 02_Validaciones.gs
 * Contiene funciones auxiliares de validación usadas por el sistema de matrícula 2026.
 * Estas funciones pueden ser llamadas desde formularios web, menús, gatillos onEdit
 * o procesos internos de escritura en la hoja principal.
 */


/**
 * Calcula el dígito verificador (DV) de un RUT chileno.
 * La función recibe el cuerpo del RUT (sin DV) y retorna
 * el dígito verificador correspondiente según el algoritmo módulo 11.
 *
 * Ejemplo:
 * calcularDV("12345678") → "5"
 *
 * @param {string|number} rut Cuerpo del RUT sin puntos ni guión.
 * @returns {string} Dígito verificador calculado ("0"-"9" o "K").
 */
function calcularDV(rut) {

  // Convertir a texto y eliminar caracteres no numéricos
  rut = String(rut).replace(/\D/g, '');

  let suma = 0;
  let multiplicador = 2;

  // Recorrer el RUT desde derecha a izquierda
  for (let i = rut.length - 1; i >= 0; i--) {

    suma += parseInt(rut[i], 10) * multiplicador;

    multiplicador++;

    // Los multiplicadores van de 2 a 7
    if (multiplicador > 7) {
      multiplicador = 2;
    }
  }

  const resto = 11 - (suma % 11);

  // Casos especiales
  if (resto === 11) {
    return '0';
  }

  if (resto === 10) {
    return 'K';
  }

  return String(resto);
}

