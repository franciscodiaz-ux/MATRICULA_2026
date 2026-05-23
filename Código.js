/**
 * syncAeLtoAlumnos()
 * Copia nuevas filas desde la hoja "03 - AeL_2026" hacia la hoja "Alumnos".
 *
 * Solo copia filas cuyo valor en columna A no exista en "Alumnos".
 *
 * Mapeo:
 * Origen:  A,B,D,E,F,G,H
 * Destino: A,B,C,D,G,H,I
 *
 * Además escribe "AeL_2026" en la columna E del destino.
 *
 * Usa un marcador interno para recordar la última fila revisada.
 */
function syncAeLtoAlumnos() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSrc = ss.getSheetByName("03 - AeL_2026");
  const shDst = ss.getSheetByName("Alumnos");

  const props = PropertiesService.getDocumentProperties();
  const KEY = "AEL_LAST_ROW";

  const lastProcessed = Number(props.getProperty(KEY) || "1");
  const lastRowSrc = shSrc.getLastRow();

  if (lastRowSrc <= lastProcessed) return;

  const startRow = Math.max(2, lastProcessed + 1);
  const numRows = lastRowSrc - startRow + 1;

  const srcValues = shSrc.getRange(startRow, 1, numRows, 8).getValues();

  const lastRowDst = shDst.getLastRow();
  const dstAValues = lastRowDst >= 2
    ? shDst.getRange(2, 1, lastRowDst - 1, 1).getValues().flat()
    : [];

  const existing = new Set(dstAValues.filter(v => v !== "" && v !== null));

  const rowsToAppend = [];

  for (const r of srcValues) {

    const A = r[0];
    if (!A) continue;

    if (!existing.has(A)) {

      rowsToAppend.push([
        r[0],        // A -> A
        r[1],        // B -> B
        r[2],        // C -> C
        r[3],        // D -> D
        "AeL_2026",  // E
        "",          // F
        r[5],        // F -> G
        r[6],        // G -> H
        r[7]         // H -> I
      ]);

      existing.add(A);
    }
  }

  if (rowsToAppend.length) {
    const insertRow = shDst.getLastRow() + 1;
    shDst.getRange(insertRow, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
  }

  props.setProperty(KEY, String(lastRowSrc));
}


/**
 * resetAeLMarker()
 * Permite reiniciar el marcador de filas procesadas.
 * Útil para volver a procesar datos si algo falló.
 */
function resetAeLMarker() {
  PropertiesService.getDocumentProperties().deleteProperty("AEL_LAST_ROW");
}


/**
 * createTimeTrigger_AeL()
 * Crea un trigger cada 1 minuto para syncAeLtoAlumnos().
 * Versión simple: no usa getProjectTriggers() (evita permiso script.scriptapp).
 * IMPORTANTE: si se ejecuta varias veces, puede crear triggers duplicados.
 */
function createTimeTrigger_AeL() {
  ScriptApp.newTrigger("syncAeLtoAlumnos")
    .timeBased()
    .everyMinutes(1)
    .create();
}


/**
 * auditarDatosCertificados(mostrarResultado)
 * ------------------------------------------------------------
 * Valida datos de la hoja "Alumnos".
 *
 * Si mostrarResultado = true → muestra alert con resultados.
 * Si mostrarResultado = false → solo devuelve el resultado.
 *
 * Devuelve:
 * { ok:true/false, errores:[...] }
 */
function auditarDatosCertificados(mostrarResultado = true) {

  const ID_MATRICULA_2026 = "1xJ-HeDRDbF4CUGI5ZZ4g3JIkAkJMI0GwQldKnDkIRyc";
  const NOMBRE_HOJA = "Alumnos";

  const ss = SpreadsheetApp.openById(ID_MATRICULA_2026);
  const sh = ss.getSheetByName(NOMBRE_HOJA);

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());

  const idxRut      = headers.indexOf("RUTAL");
  const idxDv       = headers.indexOf("DVAL");
  const idxCurso    = headers.indexOf("CURSO");
  const idxLetra    = headers.indexOf("LETRA");
  const idxApaterno = headers.indexOf("APATERNO");
  const idxNombres  = headers.indexOf("NOMBRES");

  const errores = [];
  const rutMap = {};

  for (let i = 1; i < data.length; i++) {

    const fila = data[i];
    const nroFila = i + 1;

    const rut      = String(fila[idxRut] ?? "").trim();
    const dv       = String(fila[idxDv] ?? "").trim().toUpperCase();
    const curso    = String(fila[idxCurso] ?? "").trim();
    const letra    = String(fila[idxLetra] ?? "").trim();
    const apaterno = String(fila[idxApaterno] ?? "").trim();
    const nombres  = String(fila[idxNombres] ?? "").trim();

    if (!rut)      errores.push(`Fila ${nroFila}: falta RUTAL.`);
    if (!dv)       errores.push(`Fila ${nroFila}: falta DVAL.`);
    if (!curso)    errores.push(`Fila ${nroFila}: falta CURSO.`);
    if (!letra)    errores.push(`Fila ${nroFila}: falta LETRA.`);
    if (!apaterno) errores.push(`Fila ${nroFila}: falta APATERNO.`);
    if (!nombres)  errores.push(`Fila ${nroFila}: falta NOMBRES.`);

    if (rut && dv) {

      const rutCompleto = `${rut}-${dv}`;

      if (rutMap[rutCompleto]) {
        errores.push(`Fila ${nroFila}: RUT repetido (${rutCompleto}) con fila ${rutMap[rutCompleto]}.`);
      } else {
        rutMap[rutCompleto] = nroFila;
      }
    }
  }

  const resultado = {
    ok: errores.length === 0,
    errores: errores
  };

  if (mostrarResultado) {

    if (resultado.ok) {

      SpreadsheetApp.getUi().alert(
        "Auditoría completada.\n\nNo se encontraron errores en los datos."
      );

    } else {

      SpreadsheetApp.getUi().alert(
        "Se encontraron errores en la auditoría:\n\n" +
        errores.join("\n")
      );
    }
  }

  return resultado;
}


const LIMITE_CERTIFICADOS_POR_LOTE = 100;

function obtenerConfigCertificados() {
  return {
    MAT: {
      nombre: "Certificado de Matrícula",
      prefijo: "MAT-2026-",
      plantillaId: "1-c8AeY1RR5RlSk14tGFpVuO2sidj10QOdxD_ABdkQwE",
      carpetaId: "1sBUNdZv9Kouf7Nwa305W0gpzJcUsirbR",
      colFolio: "FOLIO_MAT",
      colCodigo: "CODIGO_MAT",
      colUrl: "URL_MAT",
      colFecha: "FECHA_MAT",
      fechaModo: "FMAT_O_FIJA"
    },

    AR: {
      nombre: "Certificado de Alumno Regular",
      prefijo: "AR-2026-",
      plantillaId: "14D2QpV5jdGICvAv-sbndWHFsPgFvG-FFhWe4np3EzeI",
      carpetaId: "1Q21lPfp3btZgQxxEfOnTyW2cpOR9dfmu",
      colFolio: "FOLIO_AR",
      colCodigo: "CODIGO_AR",
      colUrl: "URL_AR",
      colFecha: "FECHA_AR",
      fechaModo: "HOY"
    }
  };
}


/**
 * formatearFechaDDMMAAAA(fecha)
 * ------------------------------------------------------------
 * Devuelve una fecha en formato dd-MM-AAAA.
 */
function formatearFechaDDMMAAAA(fecha) {
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd-MM-yyyy");
}

/**
 * generarCodigoValidacion()
 * ------------------------------------------------------------
 * Genera un código de validación de 8 caracteres hexadecimales.
 */
function generarCodigoValidacion() {
  const raw = Utilities.getUuid() + Date.now();
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    raw
  );

  return hash
    .slice(0, 4)
    .map(b => ("0" + (b & 0xff).toString(16)).slice(-2))
    .join("")
    .toUpperCase();
}

/**
 * extraerNumeroFolio(folio, prefijo)
 * ------------------------------------------------------------
 * Extrae la parte numérica de un folio tipo MAT-2026-000123.
 */
function extraerNumeroFolio(folio, prefijo) {
  if (!folio || typeof folio !== "string") return 0;
  if (!folio.startsWith(prefijo)) return 0;

  const n = parseInt(folio.replace(prefijo, ""), 10);
  return isNaN(n) ? 0 : n;
}

/**
 * construirNombreCompleto(filaObj)
 * ------------------------------------------------------------
 * Construye NOMBRES + APATERNO + AMATERNO si existe.
 */
function construirNombreCompleto(filaObj) {
  const partes = [
    filaObj.NOMBRES || "",
    filaObj.APATERNO || "",
    filaObj.AMATERNO || ""
  ].map(v => String(v).trim()).filter(Boolean);

  return partes.join(" ");
}

/**
 * construirRut(filaObj)
 * ------------------------------------------------------------
 * Construye RUTAL-DVAL.
 */
function construirRut(filaObj) {
  const rut = String(filaObj.RUTAL ?? "").trim();
  const dv  = String(filaObj.DVAL ?? "").trim().toUpperCase();

  if (!rut) return "";
  if (dv === "") return rut;

  return `${rut}-${dv}`;
}


function normalizarFechaCertificado(valor, fechaDefault = "") {
  if (valor === null || valor === undefined || valor === "") {
    return fechaDefault;
  }

  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd-MM-yyyy");
  }

  const texto = String(valor).trim();

  if (/^\d{2}-\d{2}-\d{4}$/.test(texto)) {
    return texto;
  }

  const fecha = new Date(texto);
  if (!isNaN(fecha.getTime())) {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd-MM-yyyy");
  }

  return fechaDefault || texto;
}


/**
 * construirCurso(filaObj)
 * ------------------------------------------------------------
 * Devuelve la glosa de curso para Certificado de Alumno Regular.
 *
 * Ejemplos:
 * 1 + A  -> 1° Medio A
 * 2 + B  -> 2° Medio B
 * 3 + C  -> 3° Medio C
 * 4 + A  -> 4° Medio A
 * 0      -> Laboral
 */
function construirCurso(filaObj) {
  const curso = String(filaObj.CURSO ?? "").trim();
  const letra = String(filaObj.LETRA ?? "").trim();

  if (!curso) return "";

  if (curso === "0") {
    return "Laboral";
  }

  let glosa = `${curso}° Medio`;

  if (letra) {
    glosa += ` ${letra}`;
  }

  return glosa;
}


/**
 * construirNivel(filaObj)
 * ------------------------------------------------------------
 * Si CURSO = 0 → "Laboral"
 */
function construirNivel(filaObj) {
  const curso = String(filaObj.CURSO ?? "").trim();

  const mapa = {
    "0": "Laboral",
    "1": "Primero Medio",
    "2": "Segundo Medio",
    "3": "Tercero Medio",
    "4": "Cuarto Medio"
  };

  if (mapa[curso]) {
    return mapa[curso];
  }

  // Si por alguna razón viene otro valor
  return curso;
}


/**
 * leerHojaAlumnos()
 * ------------------------------------------------------------
 * Lee la hoja Alumnos y devuelve:
 * - ss
 * - sh
 * - headers
 * - rows
 */
function leerHojaAlumnos() {
  const ID_MATRICULA_2026 = "1xJ-HeDRDbF4CUGI5ZZ4g3JIkAkJMI0GwQldKnDkIRyc";
  const ss = SpreadsheetApp.openById(ID_MATRICULA_2026);
  const sh = ss.getSheetByName("Alumnos");

  if (!sh) {
    throw new Error('No existe la hoja "Alumnos".');
  }

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('La hoja "Alumnos" no contiene filas para procesar.');
  }

  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const rows = data.slice(1);

  return { ss, sh, headers, rows };
}

/**
 * filaAObjeto(headers, fila)
 * ------------------------------------------------------------
 * Convierte una fila a objeto según encabezados.
 */
function filaAObjeto(headers, fila) {
  const obj = {};
  headers.forEach((h, i) => {
    obj[h] = fila[i];
  });
  return obj;
}

/**
 * obtenerMaximoFolioUsado(rows, headers, colFolio, prefijo)
 * ------------------------------------------------------------
 * Recorre toda la hoja una sola vez y devuelve el mayor correlativo usado.
 */
function obtenerMaximoFolioUsado(rows, headers, colFolio, prefijo) {
  const idxFolio = headers.indexOf(colFolio);
  if (idxFolio < 0) {
    throw new Error(`No existe la columna ${colFolio}.`);
  }

  let max = 99; // para que el primero nuevo sea 000100

  rows.forEach(fila => {
    const folio = String(fila[idxFolio] || "").trim();
    const n = extraerNumeroFolio(folio, prefijo);
    if (n > max) max = n;
  });

  return max;
}

/**
 * crearPdfDesdePlantilla(config, datos)
 * ------------------------------------------------------------
 * Copia una plantilla Google Docs, reemplaza marcadores,
 * exporta a PDF, lo guarda en Drive y devuelve su ID y URL.
 */
function crearPdfDesdePlantilla(config, datos) {
  Logger.log("crearPdfDesdePlantilla: inicio para RUT " + datos.RUT);

  const carpeta = DriveApp.getFolderById(config.carpetaId);
  const plantillaFile = DriveApp.getFileById(config.plantillaId);
  Logger.log("crearPdfDesdePlantilla: carpeta y plantilla obtenidas");

  let nombreArchivo = "";
  if (config.prefijo === "MAT-2026-") {
    nombreArchivo = `Certificado_Matricula_${datos.RUT}.pdf`;
  } else if (config.prefijo === "AR-2026-") {
    nombreArchivo = `Certificado_Alumno_Regular_${datos.RUT}.pdf`;
  } else {
    nombreArchivo = `${config.nombre} - ${datos.RUT}.pdf`;
  }

  Logger.log("crearPdfDesdePlantilla: nombre archivo = " + nombreArchivo);

  const copia = plantillaFile.makeCopy(`TMP - ${nombreArchivo}`);
  const copiaId = copia.getId();
  Logger.log("crearPdfDesdePlantilla: makeCopy OK, copiaId = " + copiaId);

  const doc = DocumentApp.openById(copiaId);
  const body = doc.getBody();
  Logger.log("crearPdfDesdePlantilla: documento abierto");

  body.replaceText("\\{\\{NOMBRE_COMPLETO\\}\\}", String(datos.NOMBRE_COMPLETO || ""));
  body.replaceText("\\{\\{RUT\\}\\}", String(datos.RUT || ""));
  body.replaceText("\\{\\{NIVEL\\}\\}", String(datos.NIVEL || ""));
  body.replaceText("\\{\\{CURSO\\}\\}", String(datos.CURSO || ""));
  body.replaceText("\\{\\{FOLIO\\}\\}", String(datos.FOLIO || ""));
  body.replaceText("\\{\\{CODIGO\\}\\}", String(datos.CODIGO || ""));
  body.replaceText("\\{\\{FECHA\\}\\}", String(datos.FECHA || ""));
  Logger.log("crearPdfDesdePlantilla: replaceText OK");

  doc.saveAndClose();
  Logger.log("crearPdfDesdePlantilla: saveAndClose OK");

  const pdfBlob = DriveApp.getFileById(copiaId).getBlob().getAs(MimeType.PDF);
  pdfBlob.setName(nombreArchivo);
  Logger.log("crearPdfDesdePlantilla: PDF blob OK");

  const pdfFile = carpeta.createFile(pdfBlob);
  Logger.log("crearPdfDesdePlantilla: createFile OK, fileId = " + pdfFile.getId());

  DriveApp.getFileById(copiaId).setTrashed(true);
  Logger.log("crearPdfDesdePlantilla: copia temporal enviada a papelera");

  return {
    fileId: pdfFile.getId(),
    url: pdfFile.getUrl()
  };
}

/**
 * generarCertificados(tipo)
 * ------------------------------------------------------------
 * Genera certificados masivamente.
 * tipo = "MAT" o "AR"
 */
function generarCertificados(tipo) {
  Logger.log("=== INICIO generarCertificados ===");
  Logger.log("Tipo solicitado: " + tipo);

  const audit = auditarDatosCertificados(false);
  Logger.log("Auditoría ejecutada. OK = " + audit.ok);

  if (!audit.ok) {
    SpreadsheetApp.getUi().alert(
      "No se pueden generar certificados.\n\n" +
      audit.errores.join("\n")
    );
    return;
  }

  const configs = obtenerConfigCertificados();
  const config = configs[tipo];

  if (!config) {
    throw new Error(`Tipo de certificado no válido: ${tipo}`);
  }

  const { sh, headers, rows } = leerHojaAlumnos();

  const idxFolio  = headers.indexOf(config.colFolio);
  const idxCodigo = headers.indexOf(config.colCodigo);
  const idxUrl    = headers.indexOf(config.colUrl);
  const idxFecha  = headers.indexOf(config.colFecha);

  Logger.log(
    "Índices columnas control => " +
    "Folio: " + idxFolio + ", " +
    "Codigo: " + idxCodigo + ", " +
    "URL: " + idxUrl + ", " +
    "Fecha: " + idxFecha
  );

  if ([idxFolio, idxCodigo, idxUrl, idxFecha].includes(-1)) {
    throw new Error(`Faltan columnas de control para ${config.nombre}.`);
  }

  let maxFolio = obtenerMaximoFolioUsado(rows, headers, config.colFolio, config.prefijo);
  Logger.log("Máximo folio detectado: " + maxFolio);

  let generados = 0;
  let omitidos = 0;
  let pendientes = 0;

  for (let i = 0; i < rows.length; i++) {
    const folioExistente = String(rows[i][idxFolio] || "").trim();
    if (!folioExistente) pendientes++;
  }

  Logger.log("Pendientes detectados: " + pendientes);

  if (pendientes === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "No hay certificados pendientes por generar.",
      config.nombre,
      10
    );
    return;
  }

  for (let i = 0; i < rows.length; i++) {
    const nroFila = i + 2;
    const fila = rows[i];
    const filaObj = filaAObjeto(headers, fila);

    const folioExistente = String(fila[idxFolio] || "").trim();

    if (folioExistente) {
      omitidos++;
      continue;
    }

    if (generados >= LIMITE_CERTIFICADOS_POR_LOTE) {
      Logger.log("Se alcanzó el límite del lote: " + LIMITE_CERTIFICADOS_POR_LOTE);
      break;
    }

    maxFolio++;
    const folio = `${config.prefijo}${Utilities.formatString("%06d", maxFolio)}`;
    const codigo = generarCodigoValidacion();

    let fecha = "";
    if (config.fechaModo === "FMAT_O_FIJA") {
      fecha = normalizarFechaCertificado(filaObj.FMAT, "06-03-2025");
    } else if (config.fechaModo === "HOY") {
      fecha = formatearFechaDDMMAAAA(new Date());
    }

    const datosPlantilla = {
      NOMBRE_COMPLETO: construirNombreCompleto(filaObj),
      RUT: construirRut(filaObj),
      NIVEL: construirNivel(filaObj),   // para Matrícula
      CURSO: construirCurso(filaObj),   // para Alumno Regular
      FOLIO: folio,
      CODIGO: codigo,
      FECHA: fecha
    };

    Logger.log(`Fila ${nroFila}: antes de crearPdfDesdePlantilla`);
    const archivo = crearPdfDesdePlantilla(config, datosPlantilla);
    Logger.log(`Fila ${nroFila}: PDF creado = ${archivo.url}`);

    // ✅ Escritura inmediata de las 4 columnas de control
    const rangoSalida = sh.getRange(nroFila, idxFolio + 1, 1, 4);

    rangoSalida.setValues([[
      folio,
      codigo,
      archivo.url,
      String(fecha)
    ]]);

    // forzar la celda de fecha como texto
    sh.getRange(nroFila, idxFecha + 1).setNumberFormat("@");
    sh.getRange(nroFila, idxFecha + 1).setValue(String(fecha));

    Logger.log(`Fila ${nroFila}: datos de validación escritos en hoja`);

    generados++;
  }

  const restantes = Math.max(pendientes - generados, 0);

  Logger.log(
    "FIN generarCertificados => " +
    "Generados: " + generados +
    ", Omitidos: " + omitidos +
    ", Restantes: " + restantes
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Generados en este lote: ${generados}\n` +
    `Ya existentes omitidos: ${omitidos}\n` +
    `Pendientes restantes: ${restantes}`,
    config.nombre,
    10
  );
}


/**
 * generarCertificadosMatricula()
 * ------------------------------------------------------------
 * Genera certificados de matrícula para todas las filas.
 */
function generarCertificadosMatricula() {
  generarCertificados("MAT");
}

/**
 * generarCertificadosAlumnoRegular()
 * ------------------------------------------------------------
 * Genera certificados de alumno regular para todas las filas.
 */
function generarCertificadosAlumnoRegular() {
  generarCertificados("AR");
}

/**
 * generarAmbosCertificados()
 * ------------------------------------------------------------
 * Genera ambos certificados para todas las filas.
 */
function generarAmbosCertificados() {
  const audit = auditarDatosCertificados(false);
  if (!audit.ok) {
    SpreadsheetApp.getUi().alert(
      "No se pueden generar certificados.\n\n" +
      audit.errores.join("\n")
    );
    return;
  }

  generarCertificados("MAT");
  generarCertificados("AR");
}

function pruebaPermisosDrive() {
  DriveApp.getFolderById("1sBUNdZv9Kouf7Nwa305W0gpzJcUsirbR").getName();
}

function pruebaUnCertificadoMatricula() {
  const configs = obtenerConfigCertificados();
  const config = configs.MAT;

  const { sh, headers, rows } = leerHojaAlumnos();
  const fila = rows[0]; // primera fila de datos
  const filaObj = filaAObjeto(headers, fila);

  const datosPlantilla = {
    NOMBRE_COMPLETO: construirNombreCompleto(filaObj),
    RUT: construirRut(filaObj),
    CURSO: construirCurso(filaObj),
    FOLIO: "MAT-2026-000100",
    CODIGO: "PRUEBA001",
    FECHA: normalizarFechaCertificado(filaObj.FMAT, "06-03-2025")
  };

  const archivo = crearPdfDesdePlantilla(config, datosPlantilla);

  SpreadsheetApp.getUi().alert(
    "Prueba completada.\n\nURL:\n" + archivo.url
  );
}

function diagnosticarUnCertificadoMatricula() {
  const inicio = new Date();
  Logger.log("Inicio total: " + inicio);

  const configs = obtenerConfigCertificados();
  const config = configs.MAT;

  const { headers, rows } = leerHojaAlumnos();
  const fila = rows[0];
  const filaObj = filaAObjeto(headers, fila);

  const datosPlantilla = {
    NOMBRE_COMPLETO: construirNombreCompleto(filaObj),
    RUT: construirRut(filaObj),
    NIVEL: construirNivel(filaObj),
    FOLIO: "MAT-2026-000100",
    CODIGO: "PRUEBA001",
    FECHA: normalizarFechaCertificado(filaObj.FMAT, "06-03-2025")
  };

  const t1 = new Date();
  Logger.log("Antes de obtener carpeta/plantilla: " + t1);

  const carpeta = DriveApp.getFolderById(config.carpetaId);
  const plantillaFile = DriveApp.getFileById(config.plantillaId);

  const t2 = new Date();
  Logger.log("Después de obtener carpeta/plantilla: " + t2);

  const nombreArchivo = `Diagnostico_${datosPlantilla.RUT}.pdf`;

  const copia = plantillaFile.makeCopy(`TMP - ${nombreArchivo}`);
  const copiaId = copia.getId();

  const t3 = new Date();
  Logger.log("Después de makeCopy: " + t3);

  const doc = DocumentApp.openById(copiaId);
  const body = doc.getBody();

  const t4 = new Date();
  Logger.log("Después de openById/getBody: " + t4);

  body.replaceText("\\{\\{NOMBRE_COMPLETO\\}\\}", String(datosPlantilla.NOMBRE_COMPLETO || ""));
  body.replaceText("\\{\\{RUT\\}\\}", String(datosPlantilla.RUT || ""));
  body.replaceText("\\{\\{NIVEL\\}\\}", String(datosPlantilla.NIVEL || ""));
  body.replaceText("\\{\\{FOLIO\\}\\}", String(datosPlantilla.FOLIO || ""));
  body.replaceText("\\{\\{CODIGO\\}\\}", String(datosPlantilla.CODIGO || ""));
  body.replaceText("\\{\\{FECHA\\}\\}", String(datosPlantilla.FECHA || ""));

  const t5 = new Date();
  Logger.log("Después de replaceText: " + t5);

  doc.saveAndClose();

  const t6 = new Date();
  Logger.log("Después de saveAndClose: " + t6);

  const pdfBlob = DriveApp.getFileById(copiaId).getBlob().getAs(MimeType.PDF);
  pdfBlob.setName(nombreArchivo);

  const t7 = new Date();
  Logger.log("Después de getBlob/getAs(PDF): " + t7);

  const pdfFile = carpeta.createFile(pdfBlob);

  const t8 = new Date();
  Logger.log("Después de createFile: " + t8);

  DriveApp.getFileById(copiaId).setTrashed(true);

  const fin = new Date();
  Logger.log("Fin total: " + fin);

  Logger.log("Diagnóstico completado. Archivo creado: " + pdfFile.getUrl());
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Diagnóstico completado. Revisa el log para la URL.",
    "Diagnóstico",
    10
  );
}

/**
 * resumenContactabilidadAlumnos()
 * ------------------------------------------------------------------
 * Revisa la hoja "Alumnos" y muestra:
 * - total de alumnos con datos
 * - cuántos tienen al menos un correo (y porcentaje)
 * - cuántos tienen al menos un teléfono entre apoderado titular y suplente (y porcentaje)
 * - cuántos no tienen ningún medio de contacto
 * - cuáles no tienen ningún medio de contacto
 *
 * Salida:
 * - Toast: resumen general
 * - Alert: listado de alumnos sin ningún medio de contacto
 */
function resumenContactabilidadAlumnos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const hoja = ss.getSheetByName("Alumnos");

  if (!hoja) {
    ui.alert('No existe la hoja "Alumnos".');
    return;
  }

  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    ss.toast("La hoja Alumnos no tiene registros.", "Matrícula 2026", 5);
    return;
  }

  const encabezados = datos[0];
  const idx = construirIndiceEncabezadosContacto_(encabezados);

  validarEncabezadosContacto_(idx, [
    "RUTAL",
    "DVAL",
    "NOMBRES",
    "APATERNO",
    "AMATERNO",
    "CORREOAL",
    "CORREOAPT",
    "CORREOAPS",
    "FONO1APT",
    "FONO2APT",
    "FONO1APS",
    "FONO2APS"
  ]);

  let totalAlumnos = 0;
  let conCorreo = 0;
  let conTelefono = 0;
  let sinNada = 0;

  const listadoSinNada = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];

    // Consideramos alumno válido si tiene al menos RUT o nombre
    const rut = limpiarTextoContacto_(fila[idx["RUTAL"]]);
    const dv = limpiarTextoContacto_(fila[idx["DVAL"]]);
    const nombres = limpiarTextoContacto_(fila[idx["NOMBRES"]]);
    const apaterno = limpiarTextoContacto_(fila[idx["APATERNO"]]);
    const amaterno = limpiarTextoContacto_(fila[idx["AMATERNO"]]);

    const hayAlumno = rut || nombres || apaterno || amaterno;
    if (!hayAlumno) continue;

    totalAlumnos++;

    const correos = [
      fila[idx["CORREOAL"]],
      fila[idx["CORREOAPT"]],
      fila[idx["CORREOAPS"]]
    ]
      .map(limpiarTextoContacto_)
      .filter(v => v !== "");

    const telefonos = [
      fila[idx["FONO1APT"]],
      fila[idx["FONO2APT"]],
      fila[idx["FONO1APS"]],
      fila[idx["FONO2APS"]]
    ]
      .map(limpiarTextoContacto_)
      .filter(v => v !== "");

    const tieneCorreo = correos.length > 0;
    const tieneTelefono = telefonos.length > 0;

    if (tieneCorreo) conCorreo++;
    if (tieneTelefono) conTelefono++;

    if (!tieneCorreo && !tieneTelefono) {
      sinNada++;

      const nombreCompleto = [nombres, apaterno, amaterno].filter(Boolean).join(" ").trim();
      const rutCompleto = [rut, dv].filter(Boolean).join("-");

      listadoSinNada.push(
        `${sinNada}. ${rutCompleto || "(sin RUT)"} | ${nombreCompleto || "(sin nombre)"}`
      );
    }
  }

  const pctCorreo = totalAlumnos ? ((conCorreo / totalAlumnos) * 100).toFixed(1) : "0.0";
  const pctTelefono = totalAlumnos ? ((conTelefono / totalAlumnos) * 100).toFixed(1) : "0.0";
  const pctSinNada = totalAlumnos ? ((sinNada / totalAlumnos) * 100).toFixed(1) : "0.0";

  const resumen =
    `Alumnos: ${totalAlumnos} | ` +
    `Con correo: ${conCorreo} (${pctCorreo}%) | ` +
    `Con teléfono: ${conTelefono} (${pctTelefono}%) | ` +
    `Sin contacto: ${sinNada} (${pctSinNada}%)`;

  ss.toast(resumen, "Resumen contactabilidad", 10);

  if (sinNada > 0) {
    ui.alert(
      "Alumnos sin ningún medio de contacto",
      `Total: ${sinNada}\n\n${listadoSinNada.join("\n")}`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert("Todos los alumnos tienen al menos un medio de contacto.");
  }
}


/**
 * construirIndiceEncabezadosContacto_(encabezados)
 * ------------------------------------------------------------------
 * Construye un mapa { encabezado: índice } a partir de la fila 1.
 */
function construirIndiceEncabezadosContacto_(encabezados) {
  const idx = {};
  encabezados.forEach((valor, i) => {
    idx[String(valor).trim()] = i;
  });
  return idx;
}


/**
 * validarEncabezadosContacto_(idx, requeridos)
 * ------------------------------------------------------------------
 * Verifica que existan los encabezados necesarios para el resumen.
 */
function validarEncabezadosContacto_(idx, requeridos) {
  const faltantes = requeridos.filter(nombre => idx[nombre] === undefined);
  if (faltantes.length > 0) {
    throw new Error("Faltan encabezados obligatorios en hoja Alumnos: " + faltantes.join(", "));
  }
}


/**
 * limpiarTextoContacto_(valor)
 * ------------------------------------------------------------------
 * Convierte un valor a texto y elimina espacios sobrantes.
 */
function limpiarTextoContacto_(valor) {
  return String(valor == null ? "" : valor).trim();
}



/* ====== MENÚ UNIFICADO  ====== */

/**
 * Se ejecuta automáticamente al abrir el archivo de Google Sheets.
 * Crea un menú principal, con submenús con funciones especificas.
 */
function onOpen() {
  S
  const ui = SpreadsheetApp.getUi();

   // Submenú Herramientas
  const menuHerramientas = ui.createMenu("Herramientas del sistema")
    .addItem("📞 Fomatear numeros de celular", "limpiarTelefonosHoja")
   
   // Submenú ingreso de datos
  const menuMatricula = ui.createMenu("📋 Obtener datos de ingreso")
    .addItem("Auditar RUT duplicados y alumnos sin curso", "auditarDatosCertificados")
    .addItem("Auditar medios de contacto con alumnos", "resumenContactabilidadAlumnos")
    .addItem("Completar correos institucionales alumnos", "completarCorreosAlumnosDesdeGoogle")
    .addSeparator()
    .addItem("Sincronizar AeL", "syncAeLtoAlumnos")
    .addItem("Reiniciar marcador AeL", "resetAeLMarker")
    .addSeparator()
    .addItem("Sincronizar SAE", "placeholder")
    .addSeparator()
    .addItem("Sincronizar Resultado 2025", "placeholder");

  // Submenú envíos
  const menuEnvios = ui.createMenu("Envío de certificados")
    .addItem("Enviar certificados de matrícula por correo", "enviarMasivoCorreoMatricula")
    .addItem("Enviar certificados de alumno regular por correo", "enviarMasivoCorreoAlumnoRegular")
    .addItem("Enviar ambos certificados por correo", "enviarMasivoCorreoCertificadosDisponibles")
    .addSeparator()
    .addItem("Generar enlaces WhatsApp matrícula", "generarLinksWhatsAppMatricula")
    .addItem("Generar enlaces WhatsApp alumno regular", "generarLinksWhatsAppAlumnoRegular");

  // Submenú certificados
  const menuCertificados = ui.createMenu("Certificados de alumnos")
    .addItem("Generar certificados de matrícula", "generarCertificadosMatricula")
    .addItem("Generar certificados de alumno regular", "generarCertificadosAlumnoRegular")
    .addItem("Generar ambos certificados", "generarAmbosCertificados")
    .addSeparator()
    .addSubMenu(menuEnvios);

   // Menú principal
  ui.createMenu("🎯 Funciones del sistema")
    .addSubMenu(menuHerramientas)
    .addSubMenu(menuMatricula)
    .addSubMenu(menuCertificados)
    .addToUi();
}

function placeholder() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Función aún no implementada.",
    "Sistema",
    5
  );
}


/**
 * CONFIG_ENVIOS_CERTIFICADOS()
 * ------------------------------------------------------------------
 * Devuelve la configuración general del módulo de envío masivo
 * de certificados por correo.
 */
function CONFIG_ENVIOS_CERTIFICADOS() {
  return {
    HOJA_ALUMNOS: "Alumnos",

    REMITENTE_NOMBRE: "Matrícula Liceo América",
    REMITENTE_REPLYTO: "matricula@liceoamerica.cl",

    LOTE_CORREO_MAT: 1,
    LOTE_CORREO_AR: 1,

    CAMPOS: {
      APATERNO: "APATERNO",
      AMATERNO: "AMATERNO",
      NOMBRES: "NOMBRES",
      CURSO: "CURSO",
      LETRA: "LETRA",

      CORREOAL: "CORREOAL",
      CORREOAPT: "CORREOAPT",
      CORREOAPS: "CORREOAPS",

      URL_MAT: "URL_MAT",
      URL_AR: "URL_AR",

      CORREOMAT: "CORREOMAT",
      CORREOAR: "CORREOAR"
    }
  };
}


/**
 * enviarMasivoCorreoMatricula()
 * ------------------------------------------------------------------
 * Ejecuta el envío masivo por lote de certificados de matrícula.
 */
function enviarMasivoCorreoMatricula() {
  procesarLoteCorreoCertificados_("MAT");
}


/**
 * enviarMasivoCorreoAlumnoRegular()
 * ------------------------------------------------------------------
 * Ejecuta el envío masivo por lote de certificados de alumno regular.
 */
function enviarMasivoCorreoAlumnoRegular() {
  procesarLoteCorreoCertificados_("AR");
}


/**
 * enviarMasivoCorreoCertificadosDisponibles()
 * ------------------------------------------------------------------
 * Envía por lote un solo correo por alumno, adjuntando todos los
 * certificados disponibles y aún no enviados:
 * - Certificado de matrícula
 * - Certificado de alumno regular
 *
 * Reglas:
 * - Si ambos ya fueron enviados, se omite la fila.
 * - Si no hay correos válidos, se omite la fila.
 * - Si no hay certificados disponibles para enviar, se omite la fila.
 * - Si el envío se realiza sin error, marca "SI" en cada columna
 *   correspondiente al certificado efectivamente enviado.
 */
function enviarMasivoCorreoCertificadosDisponibles() {
  const cfg = CONFIG_ENVIOS_CERTIFICADOS();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(cfg.HOJA_ALUMNOS);

  if (!hoja) {
    ui.alert(`No existe la hoja "${cfg.HOJA_ALUMNOS}".`);
    return;
  }

  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    ui.alert("La hoja Alumnos no tiene registros para procesar.");
    return;
  }

  const encabezados = datos[0];
  const idx = construirIndiceEncabezadosEnvios_(encabezados);

  validarEncabezadosEnvios_(idx, [
    cfg.CAMPOS.NOMBRES,
    cfg.CAMPOS.APATERNO,
    cfg.CAMPOS.AMATERNO,
    cfg.CAMPOS.CURSO,
    cfg.CAMPOS.LETRA,
    cfg.CAMPOS.CORREOAL,
    cfg.CAMPOS.CORREOAPT,
    cfg.CAMPOS.CORREOAPS,
    cfg.CAMPOS.URL_MAT,
    cfg.CAMPOS.URL_AR,
    cfg.CAMPOS.CORREOMAT,
    cfg.CAMPOS.CORREOAR
  ]);

  const colUrlMat = idx[cfg.CAMPOS.URL_MAT];
  const colUrlAr = idx[cfg.CAMPOS.URL_AR];
  const colCheckMat = idx[cfg.CAMPOS.CORREOMAT];
  const colCheckAr = idx[cfg.CAMPOS.CORREOAR];

  // Usa el lote de matrícula como límite del combinado.
  // Si después quieres, creamos un parámetro separado.
  const limiteLote = cfg.LOTE_CORREO_MAT;

  ss.toast("Iniciando envío combinado de certificados", "Matrícula 2026", 5);

  let enviados = 0;
  let omitidosYaEnviados = 0;
  let omitidosSinCorreo = 0;
  let omitidosSinCertificados = 0;
  let errores = 0;

  for (let i = 1; i < datos.length; i++) {
    if (enviados >= limiteLote) break;

    const fila = datos[i];
    const nroFila = i + 1;

    try {
      const yaEnviadoMat = esEnviadoCorreoEnvios_(fila[colCheckMat]);
      const yaEnviadoAr = esEnviadoCorreoEnvios_(fila[colCheckAr]);

      if (yaEnviadoMat && yaEnviadoAr) {
        omitidosYaEnviados++;
        continue;
      }

      const destinatarios = obtenerCorreosFilaEnvios_(fila, idx, cfg);
      if (destinatarios.length === 0) {
        omitidosSinCorreo++;
        continue;
      }

      const adjuntos = [];
      let envioMat = false;
      let envioAr = false;

      const urlMat = limpiarTextoEnvios_(fila[colUrlMat]);
      if (!yaEnviadoMat && urlMat) {
        const fileIdMat = extraerIdArchivoDriveEnvios_(urlMat);
        const archivoMat = DriveApp.getFileById(fileIdMat);
        adjuntos.push(archivoMat.getBlob());
        envioMat = true;
      }

      const urlAr = limpiarTextoEnvios_(fila[colUrlAr]);
      if (!yaEnviadoAr && urlAr) {
        const fileIdAr = extraerIdArchivoDriveEnvios_(urlAr);
        const archivoAr = DriveApp.getFileById(fileIdAr);
        adjuntos.push(archivoAr.getBlob());
        envioAr = true;
      }

      if (adjuntos.length === 0) {
        omitidosSinCertificados++;
        continue;
      }

      const asunto = construirAsuntoCorreoCertificadosDisponiblesEnvios_(fila, idx, envioMat, envioAr);
      const htmlBody = construirHtmlCorreoCertificadosDisponiblesEnvios_(fila, idx, envioMat, envioAr);
      const plainBody = convertirHtmlATextoPlanoEnvios_(htmlBody);

      GmailApp.sendEmail(destinatarios.join(","), asunto, plainBody, {
        htmlBody: htmlBody,
        attachments: adjuntos,
        name: cfg.REMITENTE_NOMBRE,
        replyTo: cfg.REMITENTE_REPLYTO
      });

      if (envioMat) {
        hoja.getRange(nroFila, colCheckMat + 1).setValue("SI");
      }

      if (envioAr) {
        hoja.getRange(nroFila, colCheckAr + 1).setValue("SI");
      }

      enviados++;

      // pausa aleatoria entre 0.8 y 2.2 segundos
      Utilities.sleep(Math.floor(Math.random() * 1400) + 800);

    } catch (error) {
      errores++;
      Logger.log(`Error en fila ${nroFila} (COMBINADO): ${error && error.message ? error.message : error}`);
    }
  }

  ui.alert(
    `Envío combinado finalizado\n\n` +
    `Límite del lote: ${limiteLote}\n` +
    `Correos enviados: ${enviados}\n` +
    `Omitidos por ya enviados: ${omitidosYaEnviados}\n` +
    `Omitidos sin correo válido: ${omitidosSinCorreo}\n` +
    `Omitidos sin certificados disponibles: ${omitidosSinCertificados}\n` +
    `Errores: ${errores}`
  );
}


/**
 * construirAsuntoCorreoCertificadosDisponiblesEnvios_(fila, idx, envioMat, envioAr)
 * ------------------------------------------------------------------
 * Construye el asunto del correo combinado según los certificados
 * adjuntos.
 */
function construirAsuntoCorreoCertificadosDisponiblesEnvios_(fila, idx, envioMat, envioAr) {
  const nombreAlumno = construirNombreAlumnoEnvios_(fila, idx);

  if (envioMat && envioAr) {
    return `Envío de certificados 2026 - ${nombreAlumno}`;
  }

  if (envioMat) {
    return `Envío de certificado de matrícula 2026 - ${nombreAlumno}`;
  }

  return `Envío de certificado de alumno regular 2026 - ${nombreAlumno}`;
}


/**
 * construirHtmlCorreoCertificadosDisponiblesEnvios_(fila, idx, envioMat, envioAr)
 * ------------------------------------------------------------------
 * Construye el cuerpo HTML del correo combinado.
 */
function construirHtmlCorreoCertificadosDisponiblesEnvios_(fila, idx, envioMat, envioAr) {
  const nombreAlumno = construirNombreAlumnoEnvios_(fila, idx);
  const curso = construirCursoEnvios_(fila, idx);

  let textoAdjuntos = "";
  if (envioMat && envioAr) {
    textoAdjuntos = "los certificados de matrícula y de alumno regular";
  } else if (envioMat) {
    textoAdjuntos = "el certificado de matrícula";
  } else {
    textoAdjuntos = "el certificado de alumno regular";
  }

  return `
    <p>Estimado/a:</p>
    <p>Junto con saludar, se adjunta ${textoAdjuntos} correspondiente al estudiante <strong>${escapeHtmlEnvios_(nombreAlumno)}</strong>${curso ? `, curso <strong>${escapeHtmlEnvios_(curso)}</strong>` : ""}.</p>
    <p>Este mensaje ha sido enviado automáticamente por el sistema de Matrícula 2026.</p>
    <p>Saludos cordiales.</p>
    <p><strong>Liceo Bicentenario Politécnico América</strong></p>
  `;
}



/**
 * procesarLoteCorreoCertificados_(tipo)
 * ------------------------------------------------------------------
 * Procesa un lote fijo de envíos por correo para certificados.
 *
 * tipo:
 * - "MAT" = certificado de matrícula
 * - "AR"  = certificado de alumno regular
 *
 * Reglas:
 * - Si el checkbox ya está marcado, se omite.
 * - Si no existe URL del certificado, se omite.
 * - Si no hay correos válidos, se omite.
 * - Solo cuenta como parte del lote a las filas realmente enviadas.
 * - Si el envío no falla, marca el checkbox correspondiente.
 */
function procesarLoteCorreoCertificados_(tipo) {
  const cfg = CONFIG_ENVIOS_CERTIFICADOS();
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(cfg.HOJA_ALUMNOS);

  if (!hoja) {
    ui.alert(`No existe la hoja "${cfg.HOJA_ALUMNOS}".`);
    return;
  }

  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    ui.alert("La hoja Alumnos no tiene registros para procesar.");
    return;
  }

  const encabezados = datos[0];
  const idx = construirIndiceEncabezadosEnvios_(encabezados);

  const campoUrl = tipo === "MAT" ? cfg.CAMPOS.URL_MAT : cfg.CAMPOS.URL_AR;
  const campoCheck = tipo === "MAT" ? cfg.CAMPOS.CORREOMAT : cfg.CAMPOS.CORREOAR;
  const limiteLote = tipo === "MAT" ? cfg.LOTE_CORREO_MAT : cfg.LOTE_CORREO_AR;

  validarEncabezadosEnvios_(idx, [
    cfg.CAMPOS.NOMBRES,
    cfg.CAMPOS.APATERNO,
    cfg.CAMPOS.AMATERNO,
    cfg.CAMPOS.CURSO,
    cfg.CAMPOS.LETRA,
    cfg.CAMPOS.CORREOAL,
    cfg.CAMPOS.CORREOAPT,
    cfg.CAMPOS.CORREOAPS,
    campoUrl,
    campoCheck
  ]);

  const colUrl = idx[campoUrl];
  const colCheck = idx[campoCheck];
  const nombreTipo = tipo === "MAT" ? "Matrícula" : "Alumno Regular";
  ss.toast(`Iniciando envío por lote: ${nombreTipo}`, "Matrícula 2026", 5);

  let enviados = 0;
  let omitidosMarcados = 0;
  let omitidosSinCertificado = 0;
  let omitidosSinCorreo = 0;
  let errores = 0;

  for (let i = 1; i < datos.length; i++) {
    if (enviados >= limiteLote) break;

    const fila = datos[i];
    const nroFila = i + 1;

    try {
      if (esEnviadoCorreoEnvios_(fila[colCheck])) {
        omitidosMarcados++;
        continue;
      }

      const urlCertificado = limpiarTextoEnvios_(fila[colUrl]);
      if (!urlCertificado) {
        omitidosSinCertificado++;
        continue;
      }

      const destinatarios = obtenerCorreosFilaEnvios_(fila, idx, cfg);
      if (destinatarios.length === 0) {
        omitidosSinCorreo++;
        continue;
      }

      const fileId = extraerIdArchivoDriveEnvios_(urlCertificado);
      const archivo = DriveApp.getFileById(fileId);
      const blobPdf = archivo.getBlob();

      const asunto = construirAsuntoCorreoCertificadoEnvios_(fila, idx, tipo);
      const htmlBody = construirHtmlCorreoCertificadoEnvios_(fila, idx, tipo);
      const plainBody = convertirHtmlATextoPlanoEnvios_(htmlBody);

      GmailApp.sendEmail(destinatarios.join(","), asunto, plainBody, {
        htmlBody: htmlBody,
        attachments: [blobPdf],
        name: cfg.REMITENTE_NOMBRE,
        replyTo: cfg.REMITENTE_REPLYTO
      });

      hoja.getRange(nroFila, colCheck + 1).setValue("SI");
      enviados++;

      // pausa aleatoria entre 0.8 y 2.2 segundos
      Utilities.sleep(Math.floor(Math.random() * 1400) + 800);

    } catch (error) {
      errores++;
      Logger.log(`Error en fila ${nroFila} (${tipo}): ${error && error.message ? error.message : error}`);
    }
  }

  
  ui.alert(
    `Envío por lote finalizado: ${nombreTipo}\n\n` +
    `Límite del lote: ${limiteLote}\n` +
    `Enviados: ${enviados}\n` +
    `Omitidos por ya enviados: ${omitidosMarcados}\n` +
    `Omitidos sin certificado: ${omitidosSinCertificado}\n` +
    `Omitidos sin correo válido: ${omitidosSinCorreo}\n` +
    `Errores: ${errores}`
  );
}


/**
 * construirIndiceEncabezadosEnvios_(encabezados)
 * ------------------------------------------------------------------
 * Crea un mapa con los encabezados de la hoja:
 * { NOMBRE_ENCABEZADO: indiceColumna }
 */
function construirIndiceEncabezadosEnvios_(encabezados) {
  const indice = {};
  encabezados.forEach((valor, i) => {
    indice[String(valor).trim()] = i;
  });
  return indice;
}


/**
 * validarEncabezadosEnvios_(idx, requeridos)
 * ------------------------------------------------------------------
 * Verifica que existan todos los encabezados requeridos en la hoja.
 */
function validarEncabezadosEnvios_(idx, requeridos) {
  const faltantes = requeridos.filter(nombre => idx[nombre] === undefined);
  if (faltantes.length > 0) {
    throw new Error("Faltan encabezados obligatorios en hoja Alumnos: " + faltantes.join(", "));
  }
}


/**
 * obtenerCorreosFilaEnvios_(fila, idx, cfg)
 * ------------------------------------------------------------------
 * Obtiene todos los correos válidos no vacíos de la fila:
 * CORREOAL, CORREOAPT y CORREOAPS.
 *
 * Devuelve una lista sin repetidos.
 */
function obtenerCorreosFilaEnvios_(fila, idx, cfg) {
  const candidatos = [
    fila[idx[cfg.CAMPOS.CORREOAL]],
    fila[idx[cfg.CAMPOS.CORREOAPT]],
    fila[idx[cfg.CAMPOS.CORREOAPS]]
  ];

  const unicos = new Set();

  candidatos.forEach(valor => {
    const correo = limpiarTextoEnvios_(valor).toLowerCase();
    if (correo && esCorreoValidoEnvios_(correo)) {
      unicos.add(correo);
    }
  });

  return Array.from(unicos);
}


/**
 * esCorreoValidoEnvios_(correo)
 * ------------------------------------------------------------------
 * Valida formato básico de correo electrónico.
 */
function esCorreoValidoEnvios_(correo) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo);
}


/**
 * extraerIdArchivoDriveEnvios_(url)
 * ------------------------------------------------------------------
 * Extrae el ID de un archivo de Google Drive desde una URL.
 */
function extraerIdArchivoDriveEnvios_(url) {
  const texto = limpiarTextoEnvios_(url);

  let match = texto.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  match = texto.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  throw new Error("No fue posible extraer el ID del archivo desde la URL.");
}


/**
 * construirAsuntoCorreoCertificadoEnvios_(fila, idx, tipo)
 * ------------------------------------------------------------------
 * Construye el asunto del correo según el tipo de certificado.
 */
function construirAsuntoCorreoCertificadoEnvios_(fila, idx, tipo) {
  const nombreAlumno = construirNombreAlumnoEnvios_(fila, idx);

  if (tipo === "MAT") {
    return `Envío de certificado de matrícula 2026 - ${nombreAlumno}`;
  }

  return `Envío de certificado de alumno regular 2026 - ${nombreAlumno}`;
}


/**
 * construirHtmlCorreoCertificadoEnvios_(fila, idx, tipo)
 * ------------------------------------------------------------------
 * Construye el cuerpo HTML del correo.
 */
function construirHtmlCorreoCertificadoEnvios_(fila, idx, tipo) {
  const nombreAlumno = construirNombreAlumnoEnvios_(fila, idx);
  const curso = construirCursoEnvios_(fila, idx);
  const nombreCertificado = tipo === "MAT"
    ? "certificado de matrícula"
    : "certificado de alumno regular";

  return `
    <p>Estimado/a:</p>
    <p>Junto con saludar, se adjunta el <strong>${escapeHtmlEnvios_(nombreCertificado)}</strong> correspondiente al estudiante <strong>${escapeHtmlEnvios_(nombreAlumno)}</strong>${curso ? `, curso <strong>${escapeHtmlEnvios_(curso)}</strong>` : ""}.</p>
    <p>Este mensaje ha sido enviado automáticamente por el sistema de Matrícula 2026.</p>
    <p>Saludos cordiales.</p>
    <p><strong>Liceo Bicentenario Politécnico América</strong></p>
  `;
}


/**
 * construirNombreAlumnoEnvios_(fila, idx)
 * ------------------------------------------------------------------
 * Devuelve el nombre completo del estudiante en formato:
 * NOMBRES APATERNO AMATERNO
 */
function construirNombreAlumnoEnvios_(fila, idx) {
  return [
    limpiarTextoEnvios_(fila[idx["NOMBRES"]]),
    limpiarTextoEnvios_(fila[idx["APATERNO"]]),
    limpiarTextoEnvios_(fila[idx["AMATERNO"]])
  ].filter(Boolean).join(" ").trim();
}


/**
 * construirCursoEnvios_(fila, idx)
 * ------------------------------------------------------------------
 * Construye el texto del curso en formato "CURSO LETRA".
 * Si el curso es 0, devuelve "Laboral".
 */
function construirCursoEnvios_(fila, idx) {
  const curso = limpiarTextoEnvios_(fila[idx["CURSO"]]);
  const letra = limpiarTextoEnvios_(fila[idx["LETRA"]]);

  if (curso === "0") return "Laboral";
  return `${curso} ${letra}`.trim();
}


/**
 * convertirHtmlATextoPlanoEnvios_(html)
 * ------------------------------------------------------------------
 * Convierte un cuerpo HTML simple a texto plano.
 */
function convertirHtmlATextoPlanoEnvios_(html) {
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}


/**
 * limpiarTextoEnvios_(valor)
 * ------------------------------------------------------------------
 * Convierte un valor a texto y elimina espacios sobrantes.
 */
function limpiarTextoEnvios_(valor) {
  return String(valor == null ? "" : valor).trim();
}


/**
 * esEnviadoCorreoEnvios_(valor)
 * ------------------------------------------------------------------
 * Determina si una celda indica que el correo ya fue enviado.
 * Se considera enviado solo cuando la celda contiene "SI".
 */
function esEnviadoCorreoEnvios_(valor) {
  return limpiarTextoEnvios_(valor).toUpperCase() === "SI";
}


/**
 * escapeHtmlEnvios_(texto)
 * ------------------------------------------------------------------
 * Escapa caracteres HTML básicos para evitar errores.
 */
function escapeHtmlEnvios_(texto) {
  return String(texto == null ? "" : texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * completarCorreosAlumnosDesdeGoogle()
 * ----------------------------------------------------------------
 * Recorre la hoja "Alumnos", toma APATERNO, AMATERNO y NOMBRES,
 * consulta una sola vez todos los usuarios del dominio liceoamerica.cl
 * desde Google Workspace Directory, y completa la columna CORREOAL.
 *
 * Lógica de búsqueda:
 * 1) Busca por apellidos exactos: APATERNO + AMATERNO
 * 2) Dentro de esos candidatos:
 *    - intenta nombres exactos
 *    - si no hay, intenta coincidencia parcial por nombres
 *
 * Resultado por fila:
 * - correo encontrado
 * - NO_ENCONTRADO
 * - MULTIPLE
 */
function completarCorreosAlumnosDesdeGoogle() {
  const NOMBRE_HOJA = "Alumnos";
  const DOMINIO = "liceoamerica.cl";

  Logger.log("=== INICIO completarCorreosAlumnosDesdeGoogle ===");
  Logger.log("Descargando usuarios del dominio: " + DOMINIO);

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    throw new Error(`No existe la hoja "${NOMBRE_HOJA}".`);
  }

  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    Logger.log("No hay datos para procesar.");
    Logger.log("=== FIN completarCorreosAlumnosDesdeGoogle ===");
    return;
  }

  const encabezados = datos[0];
  const idxApaterno = encabezados.indexOf("APATERNO");
  const idxAmaterno = encabezados.indexOf("AMATERNO");
  const idxNombres = encabezados.indexOf("NOMBRES");
  const idxCorreo = encabezados.indexOf("CORREOAL");

  if ([idxApaterno, idxAmaterno, idxNombres, idxCorreo].includes(-1)) {
    throw new Error("Faltan columnas requeridas: APATERNO, AMATERNO, NOMBRES, CORREOAL.");
  }

  Logger.log("Leyendo directorio completo...");
  const usuarios = obtenerTodosLosUsuariosDirectorio_(DOMINIO);
  Logger.log("Usuarios descargados: " + usuarios.length);

  Logger.log("Construyendo índice por apellidos...");
  const indiceApellidos = construirIndicePorApellidos_(usuarios);
  Logger.log("Índice construido.");

  const salida = [];

  for (let i = 1; i < datos.length; i++) {
    const filaReal = i + 1;

    const apaterno = String(datos[i][idxApaterno] || "").trim();
    const amaterno = String(datos[i][idxAmaterno] || "").trim();
    const nombres = String(datos[i][idxNombres] || "").trim();

    Logger.log("----------------------------------------");
    Logger.log("Fila " + filaReal);
    Logger.log("APATERNO: " + apaterno);
    Logger.log("AMATERNO: " + amaterno);
    Logger.log("NOMBRES: " + nombres);

    if (!apaterno && !amaterno && !nombres) {
      Logger.log("Fila vacía. Se deja correo en blanco.");
      salida.push([""]);
      continue;
    }

    const resultado = buscarCorreoEnIndice_(indiceApellidos, apaterno, amaterno, nombres);
    Logger.log("Resultado fila " + filaReal + ": " + resultado);

    salida.push([resultado]);
  }

  Logger.log("Escribiendo resultados en CORREOAL...");
  hoja.getRange(2, idxCorreo + 1, salida.length, 1).setValues(salida);

  Logger.log("Filas escritas: " + salida.length);
  Logger.log("=== FIN completarCorreosAlumnosDesdeGoogle ===");
}


/**
 * obtenerTodosLosUsuariosDirectorio_(dominio)
 * ----------------------------------------------------------------
 * Descarga todos los usuarios del dominio desde Google Workspace.
 */
function obtenerTodosLosUsuariosDirectorio_(dominio) {
  const usuarios = [];
  let pageToken = null;

  do {
    const resp = AdminDirectory.Users.list({
      domain: dominio,
      maxResults: 500,
      orderBy: "familyName",
      pageToken: pageToken
    });

    const lote = resp.users || [];
    Logger.log("Usuarios recibidos en página: " + lote.length);

    for (let i = 0; i < lote.length; i++) {
      const u = lote[i];

      usuarios.push({
        givenName: normalizarTexto_((u.name && u.name.givenName) || ""),
        familyName: normalizarTexto_((u.name && u.name.familyName) || ""),
        fullName: normalizarTexto_((u.name && u.name.fullName) || ""),
        email: String(u.primaryEmail || "").trim()
      });
    }

    pageToken = resp.nextPageToken || null;

  } while (pageToken);

  return usuarios;
}


/**
 * construirIndicePorApellidos_(usuarios)
 * ----------------------------------------------------------------
 * Crea un índice:
 *   "APATERNO AMATERNO" -> [usuarios...]
 */
function construirIndicePorApellidos_(usuarios) {
  const indice = {};

  for (let i = 0; i < usuarios.length; i++) {
    const u = usuarios[i];
    const clave = u.familyName;

    if (!clave) continue;

    if (!indice[clave]) {
      indice[clave] = [];
    }

    indice[clave].push(u);
  }

  return indice;
}


/**
 * buscarCorreoEnIndice_(indiceApellidos, apaterno, amaterno, nombres)
 * ----------------------------------------------------------------
 * Busca por capas:
 * 1) apellidos exactos
 * 2) nombres exactos
 * 3) coincidencia parcial de nombres
 */
function buscarCorreoEnIndice_(indiceApellidos, apaterno, amaterno, nombres) {
  const apellidosAlumno = normalizarTexto_(`${apaterno} ${amaterno}`);
  const nombresAlumno = normalizarTexto_(nombres);

  const candidatos = indiceApellidos[apellidosAlumno] || [];
  Logger.log("Candidatos por apellidos exactos: " + candidatos.length);

  if (candidatos.length === 0) {
    return "NO_ENCONTRADO";
  }

  // Capa 1: nombres exactos
  const exactos = candidatos.filter(u => u.givenName === nombresAlumno);

  if (exactos.length === 1) {
    Logger.log("Coincidencia exacta de nombres encontrada.");
    return exactos[0].email;
  }

  if (exactos.length > 1) {
    Logger.log("Múltiples coincidencias exactas.");
    return "MULTIPLE";
  }

  // Capa 2: coincidencia parcial por nombres
  const tokensAlumno = tokenizarTexto_(nombresAlumno);

  const parciales = candidatos.filter(u => {
    const tokensGoogle = tokenizarTexto_(u.givenName);

    // Coincidencia si todos los tokens del alumno están en Google
    return tokensAlumno.every(t => tokensGoogle.includes(t));
  });

  Logger.log("Coincidencias parciales por nombres: " + parciales.length);

  if (parciales.length === 1) {
    Logger.log("Coincidencia parcial única encontrada.");
    return parciales[0].email;
  }

  if (parciales.length > 1) {
    Logger.log("Múltiples coincidencias parciales.");
    return "MULTIPLE";
  }

  return "NO_ENCONTRADO";
}


/**
 * tokenizarTexto_(texto)
 * ----------------------------------------------------------------
 * Divide un texto normalizado en palabras no vacías.
 */
function tokenizarTexto_(texto) {
  return String(texto || "")
    .split(" ")
    .map(s => s.trim())
    .filter(Boolean);
}


/**
 * normalizarTexto_(texto)
 * ----------------------------------------------------------------
 * Convierte texto a mayúsculas, elimina tildes y espacios repetidos.
 */
function normalizarTexto_(texto) {
  return String(texto || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}


/**
 * ------------------------------------------------------------
 * FUNCIÓN: formatearCelularChile(numero)
 * ------------------------------------------------------------
 * Normaliza celulares chilenos al formato:
 *
 * +56 9 12345678
 * ------------------------------------------------------------
 */

function formatearCelularChile(numero) {

  if (!numero) return "";

  numero = String(numero);

  // Eliminar todo lo que no sea número
  let limpio = numero.replace(/\D/g, "");

  // 56912345678
  if (limpio.length === 11 && limpio.startsWith("569")) {
    limpio = limpio.substring(3);
  }

  // 912345678
  else if (limpio.length === 9 && limpio.startsWith("9")) {
    limpio = limpio.substring(1);
  }

  // 12345678
  else if (limpio.length === 8) {
    // queda igual
  }

  else {
    return "";
  }

  if (limpio.length !== 8) return "";

  return "+56 9 " + limpio;
}



/**
 * ------------------------------------------------------------
 * FUNCIÓN: limpiarTelefonosHoja()
 * ------------------------------------------------------------
 * Busca automáticamente columnas de teléfonos
 * según encabezados y normaliza los datos.
 * ------------------------------------------------------------
 */

function limpiarTelefonosHoja() {

  const hoja = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Alumnos");

  const encabezadosTelefonos = [
    "FONO1APT",
    "FONO2APT",
    "FONO1APS",
    "FONO2APS"
  ];

  const ultimaFila = hoja.getLastRow();
  const ultimaColumna = hoja.getLastColumn();

  const encabezados = hoja
    .getRange(1, 1, 1, ultimaColumna)
    .getValues()[0];

  encabezadosTelefonos.forEach(nombreColumna => {

    const indice = encabezados.indexOf(nombreColumna);

    if (indice === -1) return;

    const columna = indice + 1;

    const rango = hoja.getRange(2, columna, ultimaFila - 1, 1);

    const datos = rango.getValues();

    for (let i = 0; i < datos.length; i++) {

      datos[i][0] = formatearCelularChile(datos[i][0]);
    }

    rango.setValues(datos);

    // Alinear a la derecha
    rango.setHorizontalAlignment("right");
  });

  SpreadsheetApp.getUi().alert("Teléfonos normalizados.");
}

// Script para consultar nombrerutyfirma.com por cada RUT y actualizar la columna NOMBRES
// si el nombre devuelto tiene más información que el registrado en la hoja.
//
// INSTRUCCIONES DE USO:
// 1. Abre tu Google Sheet
// 2. Ve a Extensiones > Apps Script
// 3. Pega este código (reemplaza todo lo que haya)
// 4. Guarda (Ctrl+S)
// 5. Haz clic en "Ejecutar" con la función "actualizarNombres" seleccionada
// 6. Acepta los permisos cuando se pida
//
// CONFIGURACIÓN: Ajusta estos valores según tu hoja
var NOMBRE_HOJA = "MATRICULA_2026";  // Nombre de la pestaña con los datos
var COL_RUTAL   = 1;  // Columna A (número del RUT sin dígito verificador)
var COL_DVAL    = 2;  // Columna B (dígito verificador)
var COL_NOMBRES = 9;  // Columna I (NOMBRES del alumno)
var FILA_INICIO = 2;  // Fila donde empiezan los datos (2 = después del encabezado)

function actualizarNombres() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var hoja  = ss.getSheetByName(NOMBRE_HOJA);

  if (!hoja) {
    SpreadsheetApp.getUi().alert("No se encontró la hoja: " + NOMBRE_HOJA);
    return;
  }

  var ultimaFila = hoja.getLastRow();
  var actualizados = 0;
  var sinResultado = 0;
  var errores = 0;
  var log = [];

  for (var fila = FILA_INICIO; fila <= ultimaFila; fila++) {
    var rut    = hoja.getRange(fila, COL_RUTAL).getValue();
    var dv     = hoja.getRange(fila, COL_DVAL).getValue();
    var nombre = hoja.getRange(fila, COL_NOMBRES).getValue();

    // Saltar filas sin RUT o sin DV
    if (!rut || !dv) continue;

    // Limpiar el RUT: quitar puntos y espacios, convertir a string
    var rutLimpio = String(rut).replace(/\./g, "").replace(/\s/g, "").trim();
    var dvLimpio  = String(dv).trim().toUpperCase();

    // Construir la URL de consulta
    var url = "https://www.nombrerutyfirma.com/buscar?term=" + rutLimpio + "-" + dvLimpio;

    try {
      var respuesta = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true
      });

      var html    = respuesta.getContentText();
      var codigo  = respuesta.getResponseCode();

      if (codigo !== 200) {
        errores++;
        log.push("Fila " + fila + " RUT " + rutLimpio + "-" + dvLimpio + ": HTTP " + codigo);
        Utilities.sleep(300);
        continue;
      }

      // Extraer el nombre del HTML de respuesta
      // El sitio muestra el nombre en una etiqueta con clase o en la tabla de resultados
      var nombreWeb = extraerNombreDeHTML(html);

      if (!nombreWeb) {
        sinResultado++;
        Utilities.sleep(300);
        continue;
      }

      // Comparar: si el nombre web tiene más palabras (más información), actualizar
      var nombreActual = String(nombre).trim().toUpperCase();
      var nombreWebUC  = nombreWeb.trim().toUpperCase();

      if (debeActualizar(nombreActual, nombreWebUC)) {
        hoja.getRange(fila, COL_NOMBRES).setValue(nombreWebUC);
        actualizados++;
        log.push("Fila " + fila + ": '" + nombreActual + "' → '" + nombreWebUC + "'");
      }

    } catch (e) {
      errores++;
      log.push("Fila " + fila + " RUT " + rutLimpio + ": ERROR - " + e.message);
    }

    // Pausa entre consultas para no sobrecargar el servidor
    Utilities.sleep(500);
  }

  // Resumen final
  var resumen = "✅ Proceso completado:\n" +
    "• Actualizados: " + actualizados + "\n" +
    "• Sin resultado web: " + sinResultado + "\n" +
    "• Errores: " + errores + "\n\n";

  if (log.length > 0) {
    resumen += "Cambios realizados:\n" + log.slice(0, 30).join("\n");
    if (log.length > 30) resumen += "\n... y " + (log.length - 30) + " más.";
  }

  SpreadsheetApp.getUi().alert(resumen);
  Logger.log(resumen);
}

// Extrae el nombre completo del HTML devuelto por nombrerutyfirma.com
function extraerNombreDeHTML(html) {
  // El sitio usa diferentes estructuras; intentamos varias estrategias

  // Estrategia 1: buscar dentro de <td> o <strong> en resultado de búsqueda
  var patrones = [
    /<td[^>]*>\s*([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ\s]{5,60})\s*<\/td>/i,
    /<strong[^>]*>\s*([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ\s]{5,60})\s*<\/strong>/i,
    /class="nombre"[^>]*>\s*([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ\s]{5,60})\s*</i,
    /<h[1-4][^>]*>\s*([A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ\s]{5,60})\s*<\/h[1-4]>/i
  ];

  for (var i = 0; i < patrones.length; i++) {
    var match = html.match(patrones[i]);
    if (match && match[1]) {
      var candidato = match[1].trim();
      // Validar: debe tener al menos 2 palabras y solo letras/espacios
      if (/^[A-ZÁÉÍÓÚÑÜ][A-ZÁÉÍÓÚÑÜ\s]{5,}$/.test(candidato) && candidato.split(/\s+/).length >= 2) {
        return candidato;
      }
    }
  }

  return null;
}


