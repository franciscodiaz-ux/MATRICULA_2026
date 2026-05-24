/**
 * -------------------------------------------------------------------
 * FUNCION: configurarColumnaRut
 * -------------------------------------------------------------------
 * Configura una columna de una hoja Google Sheets para que acepte
 * únicamente números correspondientes al cuerpo del RUT chileno.
 *
 * REGLAS:
 * - Solo números.
 * - Sin puntos.
 * - Sin guión.
 * - Sin DV.
 * - Largo máximo: 9 dígitos.
 *
 * La validación se aplica directamente sobre la columna indicada.
 *
 * EJEMPLO:
 * configurarColumnaRut("Alumnos", "B");
 * -------------------------------------------------------------------
 */
function configurarColumnaRut(nombreHoja, columna) {

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);

  if (!hoja) {
    throw new Error(`No existe la hoja: ${nombreHoja}`);
  }

  // Desde fila 2 para evitar encabezados
  const rango = hoja.getRange(`${columna}2:${columna}`);

  const regla = SpreadsheetApp.newDataValidation()
    .requireTextMatchesPattern('^[0-9]{1,9}$')
    .setAllowInvalid(false)
    .setHelpText('Ingrese solo números del RUT, sin puntos, sin guión y sin DV.')
    .build();

  rango.setDataValidation(regla);
}
