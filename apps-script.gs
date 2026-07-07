const SPREADSHEET_ID = "1STfRTWj0oMiyo4nJu-PMfO7pxKg8r-l5m8c4bzXUGY4";
const SCRIPT_VERSION = 3;
const HEADERS = ["Fecha", "Categoria", "Descripcion", "Monto (ARS)", "Metodo de pago", "", "Observaciones", "Total:", ""];
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
      sheet.deleteRow(Number(params.rowNumber));
      return output_({ ok: true, version: SCRIPT_VERSION, action, sheetName: sheet.getName() }, params.callback);
    }

    const date = parseInputDate_(params.date);
    const sheetName = monthSheetName_(date);
    const targetSheet = getOrCreateMonthSheet_(spreadsheet, sheetName, date);
    const amount = Number(params.amount || 0);
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
    ];

    if (action === "update") {
      const sourceSheet = getSheetByExactName_(spreadsheet, params.sheetName);
      const sourceRow = Number(params.rowNumber);

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
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([["", "", "", "", "", "", "", "Dolar Blue:", ""]]);
  sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(2);
  sheet.getRange("A2:I2").setFontWeight("bold");
  sheet.autoResizeColumns(1, HEADERS.length);
  return sheet;
}

function ensureHeaders_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const headerRow = values.findIndex((row) => normalize_(row[0]) === "fecha");

  if (headerRow === -1) {
    sheet.insertRows(1, 2);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([["", "", "", "", "", "", "", "Dolar Blue:", ""]]);
    sheet.getRange(2, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(2);
  }
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
