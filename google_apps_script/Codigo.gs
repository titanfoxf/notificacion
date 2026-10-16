/**
 * Codigo.gs - Backend Google Apps Script para Visor GIS Subsector 1A-4
 * Conexión directa con Google Sheets y Google Drive para Notificaciones en Campo.
 * 
 * INSTRUCCIONES DE DESPLIEGUE EN 2 MINUTOS:
 * 1. Abre tu Google Sheet del padrón de notificaciones.
 * 2. En el menú superior, haz clic en: Extensiones -> Apps Script.
 * 3. Borra todo el código que aparezca y pega este archivo completo.
 * 4. Haz clic en "Implementar" (botón azul arriba a la derecha) -> "Nueva implementación".
 * 5. En tipo selecciona: "Aplicación web".
 * 6. Configura:
 *    - Descripción: "API Campo Notificaciones 1A-4"
 *    - Ejecutar como: "Yo" (tu cuenta)
 *    - Quién tiene acceso: "Cualquier persona" (Anyone)
 * 7. Haz clic en "Implementar", autoriza los permisos y COPIA la "URL de la aplicación web".
 * 8. Pega esa URL en el Visor Web GIS (dentro de la tuerca o en la configuración del modal).
 */

const NOMBRE_CARPETA_DRIVE = "FOTOS_NOTIFICACIONES_1A4";
const NOMBRE_HOJA_PADRON = "Sheet1"; // O el nombre de tu pestaña principal
const NOMBRE_HOJA_HISTORIAL = "HISTORIAL_CAMPO";

/**
 * Responde a peticiones GET (para verificar que el script está activo)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "OK",
    message: "Servicio Google Apps Script para Notificaciones 1A-4 Activo",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Recibe datos del Visor GIS (foto en base64, lote, estado, GPS, observaciones)
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return errorResponse("No se recibieron datos en la petición.");
    }

    const data = JSON.parse(e.postData.contents);
    const codPromuvi = String(data.cod_promuvi || '').trim().toUpperCase();
    const estadoNotif = String(data.estado_notificacion || 'NOTIFICADO').trim().toUpperCase();
    const notificador = String(data.notificador || '').trim();
    const observaciones = String(data.observaciones || '').trim();
    const gpsLat = data.gps_lat || '';
    const gpsLng = data.gps_lng || '';
    const gpsPrecision = data.gps_precision || '';
    const fechaHora = new Date();
    const fechaTexto = Utilities.formatDate(fechaHora, Session.getScriptTimeZone() || "GMT-5", "yyyy-MM-dd HH:mm:ss");

    if (!codPromuvi) {
      return errorResponse("El código COD_PROMUVI es obligatorio.");
    }

    // 1. GESTIONAR FOTO EN GOOGLE DRIVE (Si se adjuntó)
    let urlFotoDrive = "";
    if (data.foto_base64) {
      try {
        const folder = getOrCreateFolder(NOMBRE_CARPETA_DRIVE);
        const fileName = `${codPromuvi}_${Utilities.formatDate(fechaHora, "GMT-5", "yyyyMMdd_HHmmss")}.jpg`;
        const decodedBytes = Utilities.base64Decode(data.foto_base64);
        const blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
        const file = folder.createFile(blob);
        
        // Dar permisos de lectura pública mediante enlace para verla en el Visor
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urlFotoDrive = file.getUrl();
      } catch (errFoto) {
        console.error("Error al guardar foto en Google Drive:", errFoto);
        urlFotoDrive = `Error al guardar foto: ${errFoto.message}`;
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 2. ACTUALIZAR O INSERTAR EN LA HOJA PRINCIPAL DEL PADRÓN
    let sheetPrincipal = ss.getSheetByName(NOMBRE_HOJA_PADRON) || ss.getSheets()[0];
    let rowUpdated = false;

    if (sheetPrincipal) {
      const dataRange = sheetPrincipal.getDataRange();
      const values = dataRange.getValues();
      const headers = values[0].map(h => String(h).trim().toUpperCase());

      // Verificar y agregar columnas si no existen
      let colCod = headers.findIndex(h => h.includes('COD_PROMUVI') || h === 'CODIGO' || h === 'LOTE');
      if (colCod === -1) colCod = 0;

      let colEstadoNotif = getOrAddColumn(sheetPrincipal, headers, 'ESTADO_NOTIFICACION');
      let colFechaNotif = getOrAddColumn(sheetPrincipal, headers, 'FECHA_NOTIFICACION');
      let colNotificador = getOrAddColumn(sheetPrincipal, headers, 'NOTIFICADOR');
      let colObsNotif = getOrAddColumn(sheetPrincipal, headers, 'OBSERVACIONES_NOTIFICACION');
      let colFotoDrive = getOrAddColumn(sheetPrincipal, headers, 'FOTO_DRIVE');
      let colGps = getOrAddColumn(sheetPrincipal, headers, 'GPS_CAMPO');

      // Buscar fila correspondiente al lote
      for (let r = 1; r < values.length; r++) {
        const rowCod = String(values[r][colCod] || '').trim().toUpperCase();
        if (rowCod === codPromuvi || normCodeGas(rowCod) === normCodeGas(codPromuvi)) {
          sheetPrincipal.getRange(r + 1, colEstadoNotif).setValue(estadoNotif);
          sheetPrincipal.getRange(r + 1, colFechaNotif).setValue(fechaTexto);
          if (notificador) sheetPrincipal.getRange(r + 1, colNotificador).setValue(notificador);
          if (observaciones) sheetPrincipal.getRange(r + 1, colObsNotif).setValue(observaciones);
          if (urlFotoDrive) sheetPrincipal.getRange(r + 1, colFotoDrive).setValue(urlFotoDrive);
          if (gpsLat && gpsLng) sheetPrincipal.getRange(r + 1, colGps).setValue(`${gpsLat}, ${gpsLng} (±${gpsPrecision}m)`);
          rowUpdated = true;
          break;
        }
      }
    }

    // 3. REGISTRAR EN PESTAÑA DE HISTORIAL DE VISITAS (AUDITORÍA PERMANENTE)
    let sheetHistorial = ss.getSheetByName(NOMBRE_HOJA_HISTORIAL);
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet(NOMBRE_HOJA_HISTORIAL);
      sheetHistorial.appendRow([
        "FECHA_HORA", 
        "COD_PROMUVI", 
        "ESTADO_NOTIFICACION", 
        "NOTIFICADOR", 
        "OBSERVACIONES", 
        "GPS_LAT", 
        "GPS_LNG", 
        "GPS_PRECISION_M", 
        "URL_FOTO_DRIVE"
      ]);
      sheetHistorial.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#0f172a").setFontColor("#38bdf8");
    }

    sheetHistorial.appendRow([
      fechaTexto,
      codPromuvi,
      estadoNotif,
      notificador,
      observaciones,
      gpsLat,
      gpsLng,
      gpsPrecision,
      urlFotoDrive
    ]);

    return successResponse({
      success: true,
      cod_promuvi: codPromuvi,
      estado_notificacion: estadoNotif,
      fecha_notificacion: fechaTexto,
      foto_url: urlFotoDrive,
      actualizado_en_padron: rowUpdated,
      mensaje: `Notificación registrada correctamente para ${codPromuvi}.`
    });

  } catch (error) {
    console.error("Error general en doPost:", error);
    return errorResponse(`Error interno del servidor: ${error.message}`);
  }
}

/**
 * Obtiene o crea la carpeta de Google Drive
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Busca una columna por nombre o la crea al final
 */
function getOrAddColumn(sheet, headers, columnName) {
  let idx = headers.indexOf(columnName);
  if (idx !== -1) {
    return idx + 1;
  }
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue(columnName).setFontWeight("bold");
  headers.push(columnName);
  return newCol;
}

/**
 * Normaliza códigos tipo 1A-4_I_07 a 1A-4_I_7
 */
function normCodeGas(code) {
  if (!code) return '';
  const clean = String(code).trim().toUpperCase();
  const parts = clean.split('_');
  if (parts.length === 3) {
    const ltNum = parseInt(parts[2], 10);
    return isNaN(ltNum) ? clean : `${parts[0]}_${parts[1]}_${ltNum}`;
  }
  return clean;
}

function successResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(msg) {
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    error: msg
  })).setMimeType(ContentService.MimeType.JSON);
}
