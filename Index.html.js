// === CONFIGURACIÓN GLOBAL ===
const ATHLETE_ID = "i777901";
const API_KEY = "52r09k430jqyp2y3ordbbj686";
const ID_CARPETA_DRIVE = "1TTqMSp_l1xp-9S0ymC4gP8T1q6J0NtFW";

// PEGA AQUÍ EL ID DE TU GOOGLE SHEET (el código entre /d/ y /edit):
const SPREADSHEET_ID = "PEGA_AQUÍ_EL_ID_DE_TU_SHEET"; 
const NOMBRE_PESTANA = "Hoja 1";
// =============================

const AUTH_HEADER = "Basic " + Utilities.base64Encode("API_KEY:" + API_KEY);

/**
 * 1. DESCARGA DEL ARCHIVO .FIT (Se ejecuta cada hora)
 */
function descargarUltimaActividadADrive() {
  const hoy = new Date();
  const hace30Dias = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  const formatDate = d => d.toISOString().split("T")[0];

  const urlList = `https://intervals.icu/api/v1/athlete/${ATHLETE_ID}/activities?oldest=${formatDate(hace30Dias)}&newest=${formatDate(hoy)}`;

  const resList = UrlFetchApp.fetch(urlList, {
    headers: { "Authorization": AUTH_HEADER },
    muteHttpExceptions: true
  });

  if (resList.getResponseCode() !== 200) {
    Logger.log(`Error al consultar actividades: ${resList.getContentText()}`);
    return;
  }

  const actividades = JSON.parse(resList.getContentText());
  if (!Array.isArray(actividades) || actividades.length === 0) return;

  const actividadesReales = actividades.filter(a => a.type && a.type !== "Note" && a.moving_time > 0);
  if (actividadesReales.length === 0) return;

  const ultima = actividadesReales[actividadesReales.length - 1];
  const activityId = ultima.id;
  const fechaStr = (ultima.start_date_local || ultima.start_date || "").substring(0, 10);
  const nombreArchivo = `${fechaStr}_${activityId}.fit`;

  const carpeta = DriveApp.getFolderById(ID_CARPETA_DRIVE);
  if (carpeta.getFilesByName(nombreArchivo).hasNext()) {
    Logger.log(`El archivo ${nombreArchivo} ya existe.`);
    return;
  }

  let urlFile = `https://intervals.icu/api/v1/athlete/${ATHLETE_ID}/activities/${activityId}/file`;
  let resFile = UrlFetchApp.fetch(urlFile, {
    headers: { "Authorization": AUTH_HEADER },
    muteHttpExceptions: true
  });

  if (resFile.getResponseCode() === 404) {
    urlFile = `https://intervals.icu/api/v1/activity/${activityId}/file`;
    resFile = UrlFetchApp.fetch(urlFile, {
      headers: { "Authorization": AUTH_HEADER },
      muteHttpExceptions: true
    });
  }

  if (resFile.getResponseCode() === 200) {
    const blob = resFile.getBlob().setName(nombreArchivo);
    carpeta.createFile(blob);
    Logger.log(`¡Archivo .FIT guardado en Drive!: ${nombreArchivo}`);
  } else {
    Logger.log(`Error al descargar FIT: ${resFile.getContentText()}`);
  }
}

/**
 * 2. DESCARGA DE DATOS DE SUEÑO Y SALUD (Se ejecuta una vez al día)
 */
function syncWellnessDataASheet() {
  if (SPREADSHEET_ID === "PEGA_AQUÍ_EL_ID_DE_TU_SHEET") {
    Logger.log("Debes configurar primero la variable SPREADSHEET_ID con el ID de tu Google Sheet.");
    return;
  }

  // Tomar la fecha de hoy (Garmin sube el sueño en la mañana de hoy)
  const hoy = new Date();
  const fechaStr = hoy.toISOString().split("T")[0];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(NOMBRE_PESTANA);
  const datos = sheet.getDataRange().getValues();

  // Comprobar si el día ya se guardó para evitar duplicar filas
  for (let i = 1; i < datos.length; i++) {
    const celda = datos[i][0];
    const fechaFila = celda instanceof Date ? celda.toISOString().split("T")[0] : String(celda).substring(0, 10);
    if (fechaFila === fechaStr) {
      Logger.log(`Los datos de salud para ${fechaStr} ya estaban anotados.`);
      return;
    }
  }

  // Petición al endpoint de Wellness en Intervals.icu
  const urlWellness = `https://intervals.icu/api/v1/athlete/${ATHLETE_ID}/wellness/${fechaStr}`;
  const response = UrlFetchApp.fetch(urlWellness, {
    headers: { "Authorization": AUTH_HEADER },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log(`Error al obtener wellness (${response.getResponseCode()}): ${response.getContentText()}`);
    return;
  }

  const data = JSON.parse(response.getContentText());

  // Extraer las variables principales
  const horasSueno = data.sleepSecs ? Number((data.sleepSecs / 3600).toFixed(2)) : "";
  const sleepScore = data.sleepScore || "";
  const hrv = data.hrv || "";
  const restingHR = data.restingHR || "";
  const readiness = data.readiness || "";
  const avgStress = data.avgStress || "";

  // Añadir la fila con los valores numéricos limpios
  sheet.appendRow([
    fechaStr,
    horasSueno,
    sleepScore,
    hrv,
    restingHR,
    readiness,
    avgStress
  ]);

  Logger.log(`Métricas de salud añadidas con éxito para el día ${fechaStr}: Sueño ${horasSueno}h | Score ${sleepScore} | HRV ${hrv}`);
}