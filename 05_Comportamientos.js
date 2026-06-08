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
 * - Marcado visual de alumnos retirados al editar FECHA_RET.
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
    case "RUT_AL":
      comportamientoRutAlumno_(e);
      break;

    case "DV_AL":
      comportamientoDvAlumno_(e);
      break;
       
    case "FONO1_APT":
    case "FONO2_APT":
    case "FONO1_APS":
    case "FONO2_APS":
      comportamientoNormalizarTelefono_(rango);
      break;

    case "CORREO_AL":
    case "CORREO_APT":
    case "CORREO_APS":
      comportamientoNormalizarCorreo_(rango);
      break;

    case "FECHA_RET":
      comportamientoFechaRetiro_(e);
      break;

    default:
      return;
  }
}


/**
 * Aplica o quita el formato visual de alumno retirado según la columna FECHA_RET.
 *
 * Si FECHA_RET contiene una fecha válida:
 * - Marca toda la fila con letra roja.
 * - Marca toda la fila como tachada.
 * - Quita los hipervínculos de URL_MAT y URL_AR.
 *
 * Si FECHA_RET queda vacía:
 * - Restaura letra negra.
 * - Quita el tachado.
 * - Restaura los hipervínculos de URL_MAT y URL_AR.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Evento de edición.
 */
function comportamientoFechaRetiro_(e) {
  const rango = e.range;
  const hoja = rango.getSheet();
  const fila = rango.getRow();
  const valor = rango.getValue();

  const mapa = obtenerMapaColumnas_(hoja);

  const colUrlMat = mapa["URL_MAT"]?.numero;
  const colUrlAr = mapa["URL_AR"]?.numero;

  const ultimaColumna = hoja.getLastColumn();
  const rangoFila = hoja.getRange(fila, 1, 1, ultimaColumna);

  const esFechaValida =
    Object.prototype.toString.call(valor) === "[object Date]" &&
    !isNaN(valor.getTime());

  if (esFechaValida) {
    rangoFila
      .setFontColor("red")
      .setFontLine("line-through");

    if (colUrlMat) {
      quitarHipervinculoCelda_(hoja.getRange(fila, colUrlMat));
    }

    if (colUrlAr) {
      quitarHipervinculoCelda_(hoja.getRange(fila, colUrlAr));
    }

    return;
  }

  if (valor === "") {
    rangoFila
      .setFontColor("black")
      .setFontLine("none");

    if (colUrlMat) {
      restaurarHipervinculoCelda_(hoja.getRange(fila, colUrlMat));
    }

    if (colUrlAr) {
      restaurarHipervinculoCelda_(hoja.getRange(fila, colUrlAr));
    }
  }
}


/**
 * Quita el hipervínculo de una celda manteniendo visible el texto.
 *
 * Fuerza un RichTextValue sin LinkUrl, porque setValue() puede conservar
 * el comportamiento clickeable en celdas con URL reconocidas por Sheets.
 *
 * @param {GoogleAppsScript.Range} celda Celda que contiene una URL.
 */
function quitarHipervinculoCelda_(celda) {
  const texto = celda.getDisplayValue();
  if (!texto) return;

  const textoBloqueado = texto.replace(/^https?:\/\//i, "BLOQUEADO://");

  celda.setValue(textoBloqueado);
}


/**
 * Agrega nuevamente el hipervínculo usando el texto contenido en la celda.
 *
 * Si el contenido no parece una URL válida, no hace nada.
 *
 * @param {GoogleAppsScript.Range} celda Celda que contiene una URL.
 */
function restaurarHipervinculoCelda_(celda) {
  let texto = String(celda.getValue()).trim();
  if (!texto) return;

  texto = texto.replace(/^BLOQUEADO:\/\//i, "https://");

  if (!/^https?:\/\//i.test(texto)) return;

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(texto)
    .setLinkUrl(texto)
    .build();

  celda.setRichTextValue(richText);
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


/**
 * Revisa si el RUT_AL ingresado ya existe con matrícula vigente.
 *
 * Al editar RUT_AL:
 * - Normaliza el cuerpo del RUT solo para comparar.
 * - Formatea visualmente RUT_AL.
 * - Busca duplicado vigente.
 * - Si existe duplicado vigente, avisa y limpia RUT_AL y DV_AL.
 *
 * No calcula ni escribe DV_AL. El DV debe ser ingresado por el usuario
 * y validado por comportamientoDvAlumno_().
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Evento de edición.
 */
function comportamientoRutAlumno_(e) {
  const rango = e.range;
  const hoja = rango.getSheet();
  const filaEditada = rango.getRow();

  const rutIngresado = normalizarRutCuerpo_(rango.getValue());

  if (!rutIngresado) return;

  const mapa = obtenerMapaColumnas_(hoja);
  const colDv = mapa["DV_AL"]?.numero;

  rango.setValue(formatearRutChile(rutIngresado));

  const duplicadoVigente = buscarRutAlumnoVigenteDuplicado_(
    hoja,
    filaEditada,
    rutIngresado
  );

  if (!duplicadoVigente) return;

  rango.clearContent();

  if (colDv) {
    hoja.getRange(filaEditada, colDv).clearContent();
  }

  SpreadsheetApp.getUi().alert(
    "RUT duplicado con matrícula vigente",
    `El RUT ${rutIngresado} corresponde a ${duplicadoVigente.nombre}, actualmente matriculado en el curso ${duplicadoVigente.curso}.\n\nSe borraron RUT_AL y DV_AL para evitar un registro duplicado.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Busca si existe otro registro vigente con el mismo RUT_AL.
 *
 * Recorre la hoja Alumnos usando encabezados como referencia.
 * Ignora la fila editada. Considera matrícula vigente cuando FECHA_RET
 * está vacía.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja Alumnos.
 * @param {number} filaEditada Fila donde se acaba de editar RUT_AL.
 * @param {string} rutBuscado Cuerpo de RUT normalizado.
 * @return {Object|null} Datos del duplicado vigente o null si no existe.
 */
function buscarRutAlumnoVigenteDuplicado_(hoja, filaEditada, rutBuscado) {
  const mapa = obtenerMapaColumnas_(hoja);

  const colRut = mapa["RUT_AL"]?.numero;
  const colFechaRet = mapa["FECHA_RET"]?.numero;
  const colNombres = mapa["NOMBRES"]?.numero;
  const colApaterno = mapa["APATERNO"]?.numero;
  const colAmaterno = mapa["AMATERNO"]?.numero;
  const colCurso = mapa["CURSO"]?.numero;
  const colLetra = mapa["LETRA"]?.numero;

  if (!colRut || !colFechaRet || !colNombres || !colApaterno || !colAmaterno || !colCurso || !colLetra) {
    throw new Error("Faltan columnas requeridas para validar RUT_AL duplicado.");
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) return null;

  const ultimaColumna = hoja.getLastColumn();
  const datos = hoja.getRange(2, 1, ultimaFila - 1, ultimaColumna).getValues();

  for (let i = 0; i < datos.length; i++) {
    const filaReal = i + 2;

    if (filaReal === filaEditada) continue;

    const fila = datos[i];
    const rutFila = normalizarRutCuerpo_(fila[colRut - 1]);

    if (rutFila !== rutBuscado) continue;

    const fechaRet = fila[colFechaRet - 1];
    const tieneFechaRet = fechaRet !== "" && fechaRet !== null;

    if (tieneFechaRet) continue;

    const nombreCompleto = [
      fila[colNombres - 1],
      fila[colApaterno - 1],
      fila[colAmaterno - 1]
    ]
      .filter(valor => valor !== "" && valor !== null)
      .join(" ");

    const cursoCompleto = [
      fila[colCurso - 1],
      fila[colLetra - 1]
    ]
      .filter(valor => valor !== "" && valor !== null)
      .join(" ");

    return {
      nombre: nombreCompleto || "(sin nombre registrado)",
      curso: cursoCompleto || "(sin curso registrado)",
      fila: filaReal
    };
  }

  return null;
}

/**
 * Valida el DV_AL ingresado contra el RUT_AL de la misma fila.
 *
 * Si el dígito verificador escrito no coincide con el calculado desde RUT_AL,
 * muestra una alerta y limpia RUT_AL y DV_AL.
 *
 * Por regla de seguridad del proyecto, esta función NO elimina filas completas.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Evento de edición.
 */
function comportamientoDvAlumno_(e) {
  const hoja = e.range.getSheet();
  const fila = e.range.getRow();

  const mapa = obtenerMapaColumnas_(hoja);

  const colRut = mapa["RUT_AL"]?.numero;
  const colDv = mapa["DV_AL"]?.numero;

  if (!colRut || !colDv) {
    throw new Error("Faltan columnas RUT_AL o DV_AL para validar dígito verificador.");
  }

  const celdaRut = hoja.getRange(fila, colRut);
  const celdaDv = hoja.getRange(fila, colDv);

  const rutCuerpo = normalizarRutCuerpo_(celdaRut.getValue());
  const dvIngresado = String(celdaDv.getValue()).trim().toUpperCase();

  if (!rutCuerpo || !dvIngresado) return;

  const dvCalculado = calcularDV(rutCuerpo);

  if (dvIngresado === dvCalculado) {
    celdaRut.setValue(formatearRutChile(rutCuerpo));
    celdaDv.setValue(dvIngresado);
    return;
  }

  celdaRut.clearContent();
  celdaDv.clearContent();

  SpreadsheetApp.getUi().alert(
    "DV inválido",
    `El dígito verificador ingresado (${dvIngresado}) no corresponde al RUT ${rutCuerpo}.\n\nEl DV correcto sería ${dvCalculado}.\n\nSe borraron RUT_AL y DV_AL para evitar un registro inconsistente.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}