/**
 * Archivo: 02_Formatos.gs
 * ----------------------------------------------------
 * Formato y limpieza de datos para MATRICULA_2026.
 *
 * Este archivo contiene funciones reutilizables para normalizar
 * visualmente datos ingresados manualmente o provenientes de sistemas.
 *
 * Actualmente normaliza:
 * - RUT: formato xx.xxx.xxx
 * - DV: mayúscula, sin espacios
 * - Teléfonos: formato 56 X XXXX XXXX
 *
 * Este archivo NO valida reglas de negocio.
 * No revisa duplicados, no valida DV matemático y no bloquea matrícula.
 */


/**
 * Normaliza el cuerpo de un RUT chileno para dejarlo como texto
 * con separadores de miles usando puntos.
 *
 * Ejemplos:
 * 12345678     -> 12.345.678
 * 12.345.678   -> 12.345.678
 * 12345678-9   -> 12.345.678
 *
 * @param {*} valor Valor ingresado en la celda.
 * @return {string} RUT formateado o vacío.
 */
function formatearRutChile(valor) {
  if (valor === null || valor === undefined || valor === "") {
    return "";
  }

  let rut = String(valor).trim();

  // Si viene con guion, se toma solo el cuerpo del RUT.
  if (rut.includes("-")) {
    rut = rut.split("-")[0];
  }

  rut = rut.replace(/\D/g, "");

  if (rut === "") {
    return "";
  }

  return rut.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}


/**
 * Normaliza el dígito verificador chileno.
 *
 * Limpia espacios, elimina caracteres extraños y deja la K en mayúscula.
 * No valida si el DV corresponde matemáticamente al RUT.
 *
 * @param {*} valor Valor ingresado en la celda.
 * @return {string} DV normalizado.
 */
function formatearDvChile(valor) {
  if (valor === null || valor === undefined || valor === "") {
    return "";
  }

  return String(valor)
    .trim()
    .toUpperCase()
    .replace(/[^0-9K]/g, "")
    .substring(0, 1);
}

/**
 * Formatea visualmente un teléfono chileno ya normalizado o ingresado
 * directamente por el operador.
 *
 * Muestra el teléfono como:
 * +56 X XXXX XXXX
 *
 * Esta función NO debe usarse para guardar datos limpios, solo para
 * presentación visual.
 *
 * @param {*} valor Teléfono original o normalizado.
 * @return {string} Teléfono formateado o vacío si no corresponde.
 */
function formatearTelefonoChile(valor) {
  const telefono = normalizarTelefonoChile_(valor);

  if (telefono === "") return "";

  return (
    "+" +
    telefono.substring(0, 2) + " " +
    telefono.substring(2, 3) + " " +
    telefono.substring(3, 7) + " " +
    telefono.substring(7, 11)
  );
}

/**
 * Aplica formato RUT + DV a dos columnas consecutivas.
 *
 * La primera columna corresponde al cuerpo del RUT.
 * La segunda columna, inmediatamente a la derecha, corresponde al DV.
 *
 * Ejemplo:
 * formatearColumnasRutDv("Alumnos", "RUT");
 *
 * @param {string} nombreHoja Nombre exacto de la hoja.
 * @param {string} encabezadoRut Encabezado exacto de la columna RUT.
 */
function formatearColumnasRutDv(nombreHoja, encabezadoRut) {
  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const columnaRut = obtenerNumeroColumnaPorEncabezado_(hoja, encabezadoRut);
  const columnaDv = columnaRut + 1;
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return;
  }

  const rangoRut = hoja.getRange(2, columnaRut, ultimaFila - 1, 1);
  const rangoDv = hoja.getRange(2, columnaDv, ultimaFila - 1, 1);

  const valoresRut = rangoRut.getValues();
  const valoresDv = rangoDv.getValues();

  const rutFormateado = valoresRut.map(fila => [formatearRutChile(fila[0])]);
  const dvFormateado = valoresDv.map(fila => [formatearDvChile(fila[0])]);

  rangoRut.setNumberFormat("@");
  rangoDv.setNumberFormat("@");

  rangoRut.setValues(rutFormateado);
  rangoDv.setValues(dvFormateado);

  rangoRut.setHorizontalAlignment("right");
  rangoDv.setHorizontalAlignment("center");
}


/**
 * Aplica formato telefónico chileno a una columna completa.
 *
 * El formato visual final será:
 * 56 X XXXX XXXX
 *
 * Ejemplo:
 * formatearColumnaTelefonosChile("Alumnos", "FONO1APT");
 *
 * @param {string} nombreHoja Nombre exacto de la hoja.
 * @param {string} encabezadoTelefono Encabezado exacto de la columna teléfono.
 */
function formatearColumnaTelefonosChile(nombreHoja, encabezadoTelefono) {
  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const columna = obtenerNumeroColumnaPorEncabezado_(hoja, encabezadoTelefono);
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    return;
  }

  const rango = hoja.getRange(2, columna, ultimaFila - 1, 1);
  const valores = rango.getValues();

  const telefonosFormateados = valores.map(fila => [
    formatearTelefonoChile(fila[0])
  ]);

  rango.setNumberFormat("@");
  rango.setValues(telefonosFormateados);
  rango.setHorizontalAlignment("right");
}