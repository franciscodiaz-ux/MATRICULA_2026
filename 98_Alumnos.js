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
 * El comportamiento vivo de la hoja Alumnos, como respuestas a edición,
 * validaciones al escribir y normalizaciones de datos ingresados, debe
 * implementarse en 04_Alumnos.gs usando funciones reutilizables de
 * 02_Formatos.gs y 03_Validaciones.gs.
 */
function setupHojaAlumnos() {
  configurarDocumentoFormatoChile();

  const nombreHoja = "Alumnos";
  const hoja = obtenerHojaObligatoria_(nombreHoja);

  configurarHojaBaseChile_(hoja);
  configurarColumnasAutomaticasPorEncabezado(nombreHoja);

  Logger.log(`Setup completado para hoja: ${nombreHoja}`);
}

