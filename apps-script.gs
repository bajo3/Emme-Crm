const SPREADSHEET_ID = "1STfRTWj0oMiyo4nJu-PMfO7pxKg8r-l5m8c4bzXUGY4";
const SCRIPT_VERSION = 9;
const CALENDAR_ID = "";
const HEADERS = ["Fecha", "Categoria", "Descripcion", "Monto (ARS)", "Metodo de pago", "", "Observaciones", "Total:", "", "Cliente", "Hora", "Duracion", "Calendar Event ID", "ID Fila"];
const MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

function doGet(event) {
  try {
    const action = event.parameter.action || "list";

    if (action === "version") {
      return output_({ ok: true, version: SCRIPT_VERSION }, event.parameter.callback);
    }

    if (action === "sheets") {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheets = spreadsheet
        .getSheets()
        .filter((sheet) => !sheet.isSheetHidden())
        .map((sheet) => ({ name: sheet.getName(), gid: sheet.getSheetId() }));

      return output_({ ok: true, version: SCRIPT_VERSION, sheets }, event.parameter.callback);
    }

    if (action !== "list") {
      return output_({ ok: false, error: "Accion no soportada." }, event.parameter.callback);
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet
      .getSheets()
      .filter((sheet) => !sheet.isSheetHidden())
      .map((sheet) => ({
        name: sheet.getName(),
        values: sheet.getDataRange().getDisplayValues(),
      }));

    return output_({ ok: true, version: SCRIPT_VERSION, sheets }, event.parameter.callback);
  } catch (error) {
    return output_({ ok: false, error: error.message }, event.parameter.callback);
  }
}

function doPost(event) {
  try {
    const params = event.parameter;
    const action = params.action || "append";
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === "delete") {
      const sheet = getSheetByExactName_(spreadsheet, params.sheetName);
      const rowNumber = resolveRowNumber_(sheet, params.rowId, params.rowNumber);
      const eventId = sheet.getRange(rowNumber, 13).getDisplayValue();
      deleteCalendarEvent_(eventId);
      sheet.deleteRow(rowNumber);
      return output_({ ok: true, version: SCRIPT_VERSION, action, sheetName: sheet.getName() }, params.callback);
    }

    const date = parseInputDate_(params.date);
    const sheetName = monthSheetName_(date);
    const targetSheet = getOrCreateMonthSheet_(spreadsheet, sheetName, date);
    const amount = Number(params.amount || 0);
    const sourceSheetForUpdate = action === "update" ? getSheetByExactName_(spreadsheet, params.sheetName) : null;
    const sourceRow = action === "update" ? resolveRowNumber_(sourceSheetForUpdate, params.rowId, params.rowNumber) : null;
    const existingEventId = action === "update" ? sourceSheetForUpdate.getRange(sourceRow, 13).getDisplayValue() : "";
    const existingRowId = action === "update" ? sourceSheetForUpdate.getRange(sourceRow, HEADERS.length).getDisplayValue() : "";
    const rowId = existingRowId || Utilities.getUuid();
    const eventId = upsertCalendarEvent_(existingEventId, date, params);
    const values = [
      formatDate_(date),
      params.category || "",
      params.description || "",
      amount,
      params.method || "",
      "",
      params.notes || "",
      "",
      "",
      params.client || "",
      params.time || "",
      params.duration || "",
      eventId || "",
      rowId,
    ];

    if (action === "update") {
      const sourceSheet = sourceSheetForUpdate;

      if (sourceSheet.getName() === targetSheet.getName()) {
        sourceSheet.getRange(sourceRow, 1, 1, values.length).setValues([values]);
        sourceSheet.getRange(sourceRow, 4).setNumberFormat('"$"#,##0.00');
        return output_({ ok: true, version: SCRIPT_VERSION, action, sheetName: sourceSheet.getName(), row: sourceRow }, params.callback);
      }

      targetSheet.appendRow(values);
      const targetRow = targetSheet.getLastRow();
      targetSheet.getRange(targetRow, 4).setNumberFormat('"$"#,##0.00');
      sourceSheet.deleteRow(sourceRow);
      return output_({ ok: true, version: SCRIPT_VERSION, action, sheetName: targetSheet.getName(), row: targetRow }, params.callback);
    }

    targetSheet.appendRow(values);

    const row = targetSheet.getLastRow();
    targetSheet.getRange(row, 4).setNumberFormat('"$"#,##0.00');

    return output_({ ok: true, version: SCRIPT_VERSION, action, sheetName: targetSheet.getName(), row }, params.callback);
  } catch (error) {
    return output_({ ok: false, error: error.message }, event.parameter.callback);
  }
}

function getSheetByExactName_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`No existe la hoja ${sheetName}.`);
  }

  return sheet;
}

function getOrCreateMonthSheet_(spreadsheet, sheetName, date) {
  const existing = findMonthSheet_(spreadsheet, date);

  if (existing) {
    ensureHeaders_(existing);
    return existing;
  }

  const sheet = spreadsheet.insertSheet(sheetName);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([buildMetaRow_()]);
  sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(2);
  sheet.getRange("A2:I2").setFontWeight("bold");
  sheet.autoResizeColumns(1, HEADERS.length);
  return sheet;
}

function buildMetaRow_() {
  const row = new Array(HEADERS.length).fill("");
  row[7] = "Dolar Blue:";
  return row;
}

function resolveRowNumber_(sheet, rowId, fallbackRowNumber) {
  if (rowId) {
    const found = findRowById_(sheet, rowId);

    if (found) {
      return found;
    }

    throw new Error("No se encontro el registro (ID no coincide). Actualiza la app e intenta de nuevo.");
  }

  return Number(fallbackRowNumber);
}

function findRowById_(sheet, rowId) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    return null;
  }

  const idColumn = HEADERS.length;
  const values = sheet.getRange(3, idColumn, lastRow - 2, 1).getDisplayValues();
  const index = values.findIndex((row) => row[0] === rowId);

  return index === -1 ? null : index + 3;
}

// Ejecutar UNA vez desde el editor para asignar ID Fila a todas las filas existentes.
function migrateRowIds() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const migrated = backfillRowIds_(spreadsheet);
  Logger.log("IDs asignados: " + migrated);
  return migrated;
}

function backfillRowIds_(spreadsheet) {
  let count = 0;

  spreadsheet.getSheets().forEach((sheet) => {
    const data = sheet.getDataRange().getDisplayValues();
    const headerRow = data.findIndex((row) => normalize_(row[0]) === "fecha");

    if (headerRow === -1) {
      return;
    }

    const idColumn = HEADERS.length;
    const headerCell = sheet.getRange(headerRow + 1, idColumn);

    if (!String(headerCell.getDisplayValue()).trim()) {
      headerCell.setValue("ID Fila");
    }

    const firstDataRow = headerRow + 2;
    const lastRow = sheet.getLastRow();

    if (lastRow < firstDataRow) {
      return;
    }

    const numRows = lastRow - firstDataRow + 1;
    const idRange = sheet.getRange(firstDataRow, idColumn, numRows, 1);
    const idValues = idRange.getValues();
    const contentValues = sheet.getRange(firstDataRow, 1, numRows, 5).getDisplayValues();
    let changed = false;

    for (let i = 0; i < numRows; i += 1) {
      const hasContent = contentValues[i].some((cell) => String(cell).trim() !== "");

      if (hasContent && !String(idValues[i][0]).trim()) {
        idValues[i][0] = Utilities.getUuid();
        changed = true;
        count += 1;
      }
    }

    if (changed) {
      idRange.setValues(idValues);
    }
  });

  return count;
}

function upsertCalendarEvent_(eventId, date, params) {
  const calendar = getCalendar_();

  if (!calendar || !params.time) {
    return eventId || "";
  }

  const start = buildEventStart_(date, params.time);
  const duration = Number(params.duration || 60);
  const end = new Date(start.getTime() + duration * 60 * 1000);
  const title = buildEventTitle_(params);
  const options = {
    description: buildEventDescription_(params),
  };
  const existing = eventId ? findCalendarEvent_(calendar, eventId) : null;

  if (existing) {
    existing.setTitle(title);
    existing.setTime(start, end);
    existing.setDescription(options.description);
    return existing.getId();
  }

  return calendar.createEvent(title, start, end, options).getId();
}

function deleteCalendarEvent_(eventId) {
  const calendar = getCalendar_();
  const event = calendar && eventId ? findCalendarEvent_(calendar, eventId) : null;

  if (event) {
    event.deleteEvent();
  }
}

function getCalendar_() {
  if (!CALENDAR_ID) {
    return null;
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("No se encontro el calendario configurado.");
  }

  return calendar;
}

function findCalendarEvent_(calendar, eventId) {
  try {
    return calendar.getEventById(eventId);
  } catch (error) {
    return null;
  }
}

function buildEventStart_(date, time) {
  const parts = String(time || "09:00").split(":");
  const hours = Number(parts[0] || 9);
  const minutes = Number(parts[1] || 0);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), hours, minutes, 0, 0);
}

function buildEventTitle_(params) {
  const pieces = [params.client, params.description || params.category].filter(Boolean);
  return pieces.length ? pieces.join(" - ") : "Turno Emme";
}

function buildEventDescription_(params) {
  return [
    params.category ? `Categoria: ${params.category}` : "",
    params.description ? `Descripcion: ${params.description}` : "",
    params.method ? `Metodo: ${params.method}` : "",
    params.amount ? `Monto: $${params.amount}` : "",
    params.notes ? `Observaciones: ${params.notes}` : "",
  ].filter(Boolean).join("\n");
}

function ensureHeaders_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const headerRow = values.findIndex((row) => normalize_(row[0]) === "fecha");

  if (headerRow === -1) {
    sheet.insertRows(1, 2);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([buildMetaRow_()]);
    sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(2);
    return;
  }

  sheet.getRange(headerRow + 1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function parseInputDate_(value) {
  if (!value) {
    return new Date();
  }

  const isoMatch = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  }

  const arMatch = String(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (arMatch) {
    const year = Number(arMatch[3]) < 100 ? 2000 + Number(arMatch[3]) : Number(arMatch[3]);
    return new Date(year, Number(arMatch[2]) - 1, Number(arMatch[1]));
  }

  return new Date(value);
}

function monthSheetName_(date) {
  const month = MONTHS[date.getMonth()];
  const year = date.getFullYear();

  return year === 2025 ? month : `${month} ${String(year).slice(-2)}`;
}

function findMonthSheet_(spreadsheet, date) {
  const month = MONTHS[date.getMonth()];
  const year = date.getFullYear();
  const shortYear = String(year).slice(-2);
  const expectedNames = year === 2025
    ? [month, `${month} 25`, `${month} 2025`]
    : [`${month} ${shortYear}`, `${month} ${year}`];
  const expected = expectedNames.map(normalize_);

  return spreadsheet.getSheets().find((sheet) => expected.includes(normalize_(sheet.getName())));
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yy");
}

function normalize_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function output_(payload, callback) {
  const body = callback ? `${callback}(${JSON.stringify(payload)})` : JSON.stringify(payload);
  const mimeType = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mimeType);
}

// ---------------------------------------------------------------------------
// Importacion de turnos desde Google Calendar
// Formato del titulo: "17hs Cata Gerace (semi $21000)"
//   - hora al principio: 17hs, 15:30hs, 8:30
//   - nombre hasta el parentesis
//   - entre parentesis: servicios y precio. Varios precios = varias filas:
//     "(kap, pies $45000 esclava $10000)" -> "kap, pies" $45000 + "esclava" $10000
//   - texto despues del ultimo precio ("no pago") va a Observaciones
// Sin hora al principio (ej. "Viaje") no se importa.
// Se sincroniza cada hora: si le agregan el precio despues, la fila se actualiza;
// si el turno se borra del calendario o se le pone "cancelado", sus filas se borran.
// Solo se importan turnos hasta hoy: los de dias futuros entran cuando llega el dia.
// ---------------------------------------------------------------------------

const IMPORT_CALENDAR_ID = ""; // Vacio = calendario principal de la cuenta que corre el script.
const IMPORT_START_DATE = "04/11/26"; // Primer dia a importar (dd/mm/aa). Vacio = desde hoy.
const IMPORT_LOOKBACK_DAYS = 7; // En cada corrida automatica revisa estos dias hacia atras.
// Servicios que no aparecen en el historial del sheet. El resto se aprende de las filas cargadas.
const IMPORT_EXTRA_CATEGORIES = { esclava: "Anillos", pulsera: "Anillos", pulseras: "Anillos" };

// Ejecutar UNA vez desde el editor, logueado con la cuenta duena del calendario.
// Importa desde IMPORT_START_DATE y deja un disparador que importa cada hora.
function instalarImportacionCalendar() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === "importarTurnosCalendar")
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  const start = getImportStartDate_();
  PropertiesService.getScriptProperties().setProperty("importStart", String(start.getTime()));

  const imported = importRange_(start);
  ScriptApp.newTrigger("importarTurnosCalendar").timeBased().everyHours(1).create();
  Logger.log("Disparador instalado. Filas importadas ahora: " + imported);
}

// Ejecutar desde el editor para ver que importaria, sin escribir nada.
function probarImportacionCalendar() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const categoryMap = buildCategoryMap_(spreadsheet);
  const knownIds = collectCalendarEventIds_(spreadsheet);
  const events = getImportableEvents_(getImportStartDate_());

  events.forEach((event) => {
    const title = event.getTitle();
    const turno = parseTurno_(title);
    const day = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "dd/MM/yy");

    if (!turno) {
      Logger.log(`${day} | "${title}" -> se ignora (no es un turno)`);
      return;
    }

    const status = knownIds.has(event.getId()) ? " [ya importado]" : "";
    const items = turno.items
      .map((item) => `${resolveCategory_(item.description, categoryMap)} / ${item.description} / $${item.amount}`)
      .join(" + ");
    Logger.log(`${day} ${turno.time} | ${turno.client} | ${items}${turno.notes ? " | obs: " + turno.notes : ""}${status}`);
  });

  Logger.log("Eventos revisados: " + events.length);
}

// La llama el disparador cada hora.
function importarTurnosCalendar() {
  const lookback = new Date(Date.now() - IMPORT_LOOKBACK_DAYS * 24 * 60 * 60 * 1000);
  const importStart = Number(PropertiesService.getScriptProperties().getProperty("importStart") || 0);
  return importRange_(new Date(Math.max(lookback.getTime(), importStart)));
}

// Sincroniza el calendario con el sheet desde "from" hasta hoy:
//   - evento nuevo           -> agrega sus filas
//   - titulo cambiado        -> reemplaza sus filas (ej. le pusieron el precio al terminar),
//                               conservando el metodo de pago cargado en la app
//   - evento borrado o con "cancel" en el titulo -> borra sus filas
// Solo toca filas creadas por esta importacion (las que tienen titulo guardado).
function importRange_(from) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const properties = PropertiesService.getScriptProperties();
    const storedTitles = readStoredTitles_(properties);
    const rowsByEvent = indexImportedRows_(spreadsheet);
    const categoryMap = buildCategoryMap_(spreadsheet);
    const rowsToDelete = [];
    const seenIds = new Set();
    const titleUpdates = {};
    const stats = { nuevas: 0, actualizadas: 0, canceladas: 0 };

    getImportableEvents_(from).forEach((event) => {
      const eventId = event.getId();
      const signature = eventSignature_(event);
      const existingRows = rowsByEvent[eventId] || [];
      const turno = parseTurno_(event.getTitle());
      const tracked = eventId in storedTitles;
      seenIds.add(eventId);

      if (tracked && storedTitles[eventId] === signature) {
        // Sin cambios. Si borraron sus filas a mano en la app, no se vuelven a crear.
        return;
      }

      if (!tracked && existingRows.length) {
        // Filas importadas por una version anterior (sin titulo guardado): se registran
        // tal como estan y desde ahora se sincronizan.
        titleUpdates[eventId] = signature;
        return;
      }

      if (existingRows.length) {
        rowsToDelete.push(...existingRows);
      }

      if (turno) {
        const method = existingRows.map((row) => row.method).find(Boolean) || "";
        appendTurnoRows_(spreadsheet, event, turno, method, categoryMap);
        stats[existingRows.length ? "actualizadas" : "nuevas"] += 1;
      } else if (existingRows.length) {
        stats.canceladas += 1; // Le pusieron "cancelado" al titulo.
      }

      if (turno || tracked) {
        titleUpdates[eventId] = signature;
      }
    });

    // Turnos borrados del calendario (o movidos a otro dia que todavia no llego).
    Object.keys(rowsByEvent).forEach((eventId) => {
      if (seenIds.has(eventId) || !(eventId in storedTitles)) {
        return;
      }

      const rows = rowsByEvent[eventId].filter((row) => row.date && row.date >= from);

      if (rows.length) {
        rowsToDelete.push(...rows);
        stats.canceladas += 1;
      }
    });

    deleteRows_(rowsToDelete);
    saveStoredTitles_(properties, storedTitles, titleUpdates, seenIds);
    Logger.log(`Turnos nuevos: ${stats.nuevas} | Actualizados: ${stats.actualizadas} | Cancelados: ${stats.canceladas}`);
    return stats.nuevas;
  } finally {
    lock.releaseLock();
  }
}

function appendTurnoRows_(spreadsheet, event, turno, method, categoryMap) {
  const start = event.getStartTime();
  const date = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const sheet = getOrCreateMonthSheet_(spreadsheet, monthSheetName_(date), date);
  const duration = String(Math.round((event.getEndTime().getTime() - start.getTime()) / 60000));

  turno.items.forEach((item, index) => {
    sheet.appendRow([
      formatDate_(date),
      resolveCategory_(item.description, categoryMap),
      item.description,
      item.amount,
      method,
      "",
      index === turno.items.length - 1 ? turno.notes : "",
      "",
      "",
      turno.client,
      turno.time,
      duration,
      event.getId(),
      Utilities.getUuid(),
    ]);
    sheet.getRange(sheet.getLastRow(), 4).setNumberFormat('"$"#,##0.00');
  });

  return turno.items.length;
}

// eventId -> [{ sheet, row, method, date }] de todas las filas que tienen Calendar Event ID.
function indexImportedRows_(spreadsheet) {
  const index = {};

  spreadsheet.getSheets().forEach((sheet) => {
    const lastRow = sheet.getLastRow();

    if (lastRow < 1 || sheet.getLastColumn() < 13) {
      return;
    }

    sheet.getRange(1, 1, lastRow, 13).getDisplayValues().forEach((values, offset) => {
      const eventId = String(values[12]).trim();

      if (!eventId || normalize_(values[0]) === "fecha") {
        return;
      }

      const date = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(values[0]).trim()) ? parseInputDate_(String(values[0]).trim()) : null;
      index[eventId] = index[eventId] || [];
      index[eventId].push({ sheet, row: offset + 1, method: String(values[4]).trim(), date });
    });
  });

  return index;
}

// Borra de abajo hacia arriba en cada hoja para no correr los numeros de fila.
function deleteRows_(rows) {
  const bySheet = {};

  rows.forEach((entry) => {
    const name = entry.sheet.getName();
    bySheet[name] = bySheet[name] || { sheet: entry.sheet, rows: [] };
    bySheet[name].rows.push(entry.row);
  });

  Object.keys(bySheet).forEach((name) => {
    const group = bySheet[name];
    Array.from(new Set(group.rows))
      .sort((a, b) => b - a)
      .forEach((row) => group.sheet.deleteRow(row));
  });
}

// Titulo de cada evento importado, para detectar cambios. Se guarda solo lo que esta en la ventana revisada.
function readStoredTitles_(properties) {
  const titles = {};
  const all = properties.getProperties();

  Object.keys(all).forEach((key) => {
    if (key.indexOf("t:") === 0) {
      titles[key.slice(2)] = all[key];
    }
  });

  return titles;
}

// Titulo + dia: si cambia cualquiera de los dos, se vuelven a generar las filas.
function eventSignature_(event) {
  const day = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return `${day}|${event.getTitle()}`;
}

function saveStoredTitles_(properties, storedTitles, titleUpdates, seenIds) {
  Object.keys(storedTitles).forEach((eventId) => {
    if (!seenIds.has(eventId)) {
      properties.deleteProperty("t:" + eventId);
    }
  });

  const batch = {};
  Object.keys(titleUpdates).forEach((eventId) => {
    batch["t:" + eventId] = titleUpdates[eventId];
  });
  properties.setProperties(batch, false);
}

function getImportStartDate_() {
  const date = IMPORT_START_DATE ? parseInputDate_(IMPORT_START_DATE) : new Date();
  date.setHours(0, 0, 0, 0);
  return date;
}

function getImportableEvents_(from) {
  const calendar = IMPORT_CALENDAR_ID ? CalendarApp.getCalendarById(IMPORT_CALENDAR_ID) : CalendarApp.getDefaultCalendar();

  if (!calendar) {
    throw new Error("No se encontro el calendario a importar.");
  }

  // Hasta el final de hoy: los turnos del dia se cargan aunque el evento figure mas tarde.
  const endOfToday = new Date();
  endOfToday.setHours(23, 59, 59, 999);

  return calendar.getEvents(from, endOfToday);
}

// "17hs Cata Gerace (semi $21000)" -> { time: "17:00", client: "Cata Gerace", items: [...], notes: "" }
// "$6000 anillo" (venta sin hora) -> { time: "", client: "", items: [{ "anillo", 6000 }] }
// Devuelve null si no empieza con hora ni con precio (ej. "Viaje", "Depi").
function parseTurno_(title) {
  const raw = String(title || "").replace(/\s+/g, " ").trim();
  const timeMatch = raw.match(/^(\d{1,2})(?:[:.](\d{2}))?\s*(hs?)?(?=\s|$)/i);

  if (/cancel/i.test(raw)) {
    return null;
  }

  if (/^\$\s*\d/.test(raw)) {
    const sale = splitItems_(raw);
    return { time: "", client: "", items: sale.items, notes: sale.notes };
  }

  if (!timeMatch || (!timeMatch[2] && !timeMatch[3])) {
    return null;
  }

  const hours = Number(timeMatch[1]);
  const minutes = Number(timeMatch[2] || 0);

  if (hours > 23 || minutes > 59) {
    return null;
  }

  const rest = raw.slice(timeMatch[0].length).trim();
  const parenIndex = rest.indexOf("(");
  const client = (parenIndex === -1 ? rest : rest.slice(0, parenIndex)).trim();
  const content = parenIndex === -1 ? "" : rest.slice(parenIndex + 1).replace(/\)\s*$/, "").trim();
  const { items, notes } = splitItems_(content);

  return {
    time: `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`,
    client,
    items,
    notes,
  };
}

// "kap, pies $45000 esclava $10000" -> [{ "kap, pies", 45000 }, { "esclava", 10000 }]
// "pies $22000 $4000 anillo"        -> [{ "pies", 22000 }, { "anillo", 4000 }]
function splitItems_(content) {
  const parts = String(content || "").split(/\$\s*([\d.,]+)/);
  const texts = parts.filter((_, index) => index % 2 === 0).map(cleanItemText_);
  const amounts = parts.filter((_, index) => index % 2 === 1).map(parseAmount_);

  if (!amounts.length) {
    return { items: [{ description: texts[0] || "", amount: 0 }], notes: "" };
  }

  const items = [];
  let trailingUsed = false;

  amounts.forEach((amount, index) => {
    let description = texts[index];

    if (!description && index === amounts.length - 1 && texts[index + 1]) {
      description = texts[index + 1];
      trailingUsed = true;
    }

    items.push({ description, amount });
  });

  const notes = trailingUsed ? "" : texts[amounts.length] || "";
  return { items, notes };
}

function cleanItemText_(text) {
  return String(text || "").replace(/^[\s,;+-]+|[\s,;+-]+$/g, "").replace(/\s+/g, " ");
}

function parseAmount_(value) {
  return Number(String(value).replace(/[.,]\d{1,2}$/, "").replace(/[.,]/g, "")) || 0;
}

function collectCalendarEventIds_(spreadsheet) {
  const ids = new Set();

  spreadsheet.getSheets().forEach((sheet) => {
    const lastRow = sheet.getLastRow();

    if (lastRow < 1 || sheet.getLastColumn() < 13) {
      return;
    }

    sheet.getRange(1, 13, lastRow, 1).getDisplayValues().forEach((row) => {
      const id = String(row[0]).trim();
      if (id) {
        ids.add(id);
      }
    });
  });

  return ids;
}

// Aprende del historial: primera palabra de la descripcion -> categoria mas usada.
function buildCategoryMap_(spreadsheet) {
  const counts = {};
  const totals = {};

  spreadsheet.getSheets().forEach((sheet) => {
    const lastRow = sheet.getLastRow();

    if (lastRow < 3) {
      return;
    }

    sheet.getRange(1, 2, lastRow, 2).getDisplayValues().forEach((row) => {
      const category = String(row[0]).trim();
      const key = firstWord_(row[1]);

      if (!category || !key || normalize_(category) === "categoria") {
        return;
      }

      const label = category.charAt(0).toUpperCase() + category.slice(1);
      counts[key] = counts[key] || {};
      counts[key][label] = (counts[key][label] || 0) + 1;
      totals[label] = (totals[label] || 0) + 1;
    });
  });

  const map = {};
  Object.keys(counts).forEach((key) => {
    map[key] = mostFrequent_(counts[key]);
  });
  Object.keys(IMPORT_EXTRA_CATEGORIES).forEach((key) => {
    map[normalize_(key)] = IMPORT_EXTRA_CATEGORIES[key];
  });
  map._default = mostFrequent_(totals) || "Uñas";
  return map;
}

function resolveCategory_(description, categoryMap) {
  return categoryMap[firstWord_(description)] || categoryMap._default;
}

function firstWord_(text) {
  const words = normalize_(text).split(/[^a-z]+/).filter((word) => word.length > 1);
  return words[0] || "";
}

function mostFrequent_(counter) {
  return Object.keys(counter).sort((a, b) => counter[b] - counter[a])[0] || "";
}
