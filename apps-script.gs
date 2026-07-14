const SPREADSHEET_ID = "1STfRTWj0oMiyo4nJu-PMfO7pxKg8r-l5m8c4bzXUGY4";
const SCRIPT_VERSION = 6;
const CALENDAR_ID = "";
const HEADERS = ["Fecha", "Categoria", "Descripcion", "Monto (ARS)", "Metodo de pago", "", "Observaciones", "Total:", "", "Cliente", "Hora", "Duracion", "Calendar Event ID", "ID Fila"];
const MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

function doGet(event) {
  try {
    const action = event.parameter.action || "list";

    if (action === "version") {
      return output_({ ok: true, version: SCRIPT_VERSION }, event.parameter.callback);
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
