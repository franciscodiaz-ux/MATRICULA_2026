/**
 * Configura la hoja Alumnos.
 *
 * Esta función es el punto de entrada explícito para preparar
 * la hoja Alumnos. Solo modifica esta hoja.
 *
 * Aplica:
 * - configuración general del documento;
 * - formato base de la hoja;
 * - autodetección de columnas según encabezado.
 *
 * Además, genera logs detallados para verificar qué columnas fueron
 * encontradas y qué acciones de configuración fueron aplicadas.
 *
 * Esta función no normaliza datos, no valida datos ingresados,
 * no calcula DV, no asigna números de registro y no modifica valores
 * existentes. Solo prepara formato y tipos de columnas.
 */
function setupHojaAlumnos() {
  const nombreHoja = "Alumnos";

  Logger.log("============================================================");
  Logger.log(`INICIO setupHojaAlumnos()`);
  Logger.log(`Hoja objetivo: ${nombreHoja}`);
  Logger.log("Acciones esperadas:");
  Logger.log("- Configurar documento con locale es_CL y timezone America/Santiago.");
  Logger.log("- Aplicar formato visual base a la hoja.");
  Logger.log("- Autodetectar columnas por encabezado.");
  Logger.log("- Configurar RUT, DV, teléfonos, celulares y correos como texto.");
  Logger.log("- Configurar fechas con formato chileno dd/mm/yyyy.");
  Logger.log("- Configurar monedas CLP si existen encabezados monetarios.");
  Logger.log("Acciones NO realizadas:");
  Logger.log("- No normaliza datos.");
  Logger.log("- No valida datos.");
  Logger.log("- No calcula dígitos verificadores.");
  Logger.log("- No asigna números de registro.");
  Logger.log("- No borra ni modifica valores existentes.");
  Logger.log("------------------------------------------------------------");

  configurarDocumentoFormatoChile();

  const hoja = obtenerHojaObligatoria_(nombreHoja);
  const mapaColumnas = obtenerMapaColumnas_(hoja);
  const columnas = Object.values(mapaColumnas);

  Logger.log(`Hoja encontrada: ${hoja.getName()}`);
  Logger.log(`Filas actuales con datos: ${hoja.getLastRow()}`);
  Logger.log(`Columnas actuales: ${hoja.getLastColumn()}`);
  Logger.log(`Encabezados detectados: ${columnas.length}`);

  columnas.forEach(columna => {
    Logger.log(
      `Encabezado detectado: Columna ${columna.letra} (${columna.numero}) | ${columna.encabezado}`
    );
  });

  configurarHojaBaseChile_(hoja);
  configurarColumnasAutomaticasPorEncabezado(nombreHoja);

  Logger.log("------------------------------------------------------------");
  Logger.log(`FIN setupHojaAlumnos(). Setup completado para hoja: ${nombreHoja}`);
  Logger.log("============================================================");
}

