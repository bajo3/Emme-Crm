const CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ6JUiLN7Xll5G9KfMi981z-N6RszRJ6tBywxV-3pK_zrhPLM_6539z16ZCxJ2SL9j6CahUyNTG-XCv/pub?output=csv";
const SPREADSHEET_ID = "1STfRTWj0oMiyo4nJu-PMfO7pxKg8r-l5m8c4bzXUGY4";
// Pega aca la URL /exec de la implementacion web de Apps Script.
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyJesbBhDhsamI0VZlGh9wd2UfeRgjIvZzTudlHUPC4s8BLedJ5u8cF35m_aLRNiJOa/exec";
const REQUIRED_SCRIPT_VERSION = 2;
const FALLBACK_SHEET_NAMES = ["Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero 26", "Febrero 26", "Marzo 26", "Abril 26", "Mayo 26", "Junio 26", "Julio 26"];
const FALLBACK_SHEETS = FALLBACK_SHEET_NAMES.map((name) => ({
  name,
  url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(name)}`,
}));

const state = {
  rows: [],
  meta: {},
  source: "csv",
  backendVersion: getApiUrl() ? REQUIRED_SCRIPT_VERSION : 0,
  activeView: "entry",
  filters: {
    month: "all",
    category: "all",
    method: "all",
    search: "",
  },
};

const els = {};
let activeLoadId = 0;

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  loadData();
});

function cacheElements() {
  [
    "refreshButton",
    "statusDot",
    "statusText",
    "sheetMeta",
    "monthFilter",
    "categoryFilter",
    "methodFilter",
    "searchFilter",
    "totalRevenue",
    "jobCount",
    "avgTicket",
    "cashTotal",
    "transferTotal",
    "categoryTotal",
    "dailyTotal",
    "categoryChart",
    "dailyChart",
    "visibleRows",
    "rowsTable",
    "entryForm",
    "entryStatus",
    "entryDate",
    "entryCategory",
    "entryDescription",
    "entryAmount",
    "entryMethod",
    "entryNotes",
    "categoryOptions",
    "saveEntryButton",
  ].forEach((id) => {
    els[id] = document.getElementById(id);
  });
  els.menuButtons = [...document.querySelectorAll("[data-view-target]")];
  els.viewSections = [...document.querySelectorAll("[data-view-section]")];
  els.analyticsTools = document.querySelector(".analytics-tools");
}

function bindEvents() {
  els.entryDate.value = todayInputValue();
  updateEntryUi();

  els.refreshButton.addEventListener("click", loadData);
  els.entryForm.addEventListener("submit", handleEntrySubmit);
  els.menuButtons.forEach((button) => {
    button.addEventListener("click", () => setActiveView(button.dataset.viewTarget));
  });
  els.monthFilter.addEventListener("change", (event) => {
    state.filters.month = event.target.value;
    render();
  });
  els.categoryFilter.addEventListener("change", (event) => {
    state.filters.category = event.target.value;
    render();
  });
  els.methodFilter.addEventListener("change", (event) => {
    state.filters.method = event.target.value;
    render();
  });
  els.searchFilter.addEventListener("input", (event) => {
    state.filters.search = event.target.value.trim().toLowerCase();
    render();
  });
  setActiveView(state.activeView);
}

function setActiveView(view) {
  state.activeView = view;

  els.menuButtons.forEach((button) => {
    const active = button.dataset.viewTarget === view;
    button.classList.toggle("active", active);
    if (active) {
      button.setAttribute("aria-current", "page");
    } else {
      button.removeAttribute("aria-current");
    }
  });

  els.viewSections.forEach((section) => {
    section.classList.toggle("active", section.dataset.viewSection === view);
  });

  if (els.analyticsTools) {
    els.analyticsTools.classList.toggle("active", view !== "entry");
  }

  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function updateEntryUi() {
  const connected = Boolean(getApiUrl());
  const updatedBackend = state.backendVersion >= REQUIRED_SCRIPT_VERSION;
  const canSave = connected && updatedBackend;
  els.entryStatus.textContent = canSave ? "Listo para guardar" : connected ? "Actualiza Apps Script" : "Guardado pendiente";
  els.saveEntryButton.disabled = !canSave;
  els.saveEntryButton.title = canSave ? "Guardar en Google Sheets" : connected ? "Redeploya Apps Script con la version corregida" : "Falta pegar la URL /exec de Apps Script en app.js";
}

function getApiUrl() {
  return clean(APPS_SCRIPT_URL);
}

function handleEntrySubmit(event) {
  event.preventDefault();
  const apiUrl = getApiUrl();

  if (!apiUrl) {
    els.entryStatus.textContent = "Falta pegar la URL /exec en app.js";
    return;
  }

  const formData = new FormData(els.entryForm);
  formData.set("action", "append");
  formData.set("amount", String(parseMoney(formData.get("amount"))));

  submitToHiddenFrame(apiUrl, formData);
  els.entryStatus.textContent = "Enviando a Google Sheets...";
  els.saveEntryButton.disabled = true;

  window.setTimeout(() => {
    els.entryForm.reset();
    els.entryDate.value = todayInputValue();
    els.entryMethod.value = "Transferencia";
    els.entryStatus.textContent = "Registro enviado";
    updateEntryUi();
    loadData();
  }, 1600);
}

function submitToHiddenFrame(url, formData) {
  const form = document.createElement("form");
  form.method = "POST";
  form.action = url;
  form.target = "writeFrame";
  form.style.display = "none";

  formData.forEach((value, key) => {
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = key;
    input.value = value;
    form.appendChild(input);
  });

  document.body.appendChild(form);
  form.submit();
  form.remove();
}

async function loadData() {
  const loadId = ++activeLoadId;
  setStatus("loading", "Cargando mes actual...");
  verifyBackendVersion(loadId);

  try {
    const currentSheet = getCurrentSheet();
    const parsed = await loadPublishedSheet(currentSheet);

    state.rows = parsed.rows.sort((a, b) => b.dateValue - a.dateValue);
    state.meta = parsed.meta;
    state.source = "public-sheets-fast";

    buildFilterOptions();
    render();
    setStatus("ready", `Mes actual cargado: ${formatDateTime(new Date())}`);
    renderSheetMeta();
    updateEntryUi();
    loadRemainingSheets(loadId, currentSheet.name);
  } catch (error) {
    console.error(error);
    setStatus("error", "No se pudo cargar la planilla.");
    els.sheetMeta.textContent = "Revisa la publicacion CSV o la URL de Apps Script.";
  } finally {
    if (window.lucide) {
      window.lucide.createIcons();
    }
  }
}

async function verifyBackendVersion(loadId) {
  const apiUrl = getApiUrl();

  if (!apiUrl) {
    state.backendVersion = 0;
    updateEntryUi();
    return;
  }

  try {
    const data = await jsonp(apiUrl, { action: "version", cacheBust: Date.now() });

    if (loadId !== activeLoadId) {
      return;
    }

    if (data.ok && data.version) {
      state.backendVersion = Number(data.version || 0);
    }
    updateEntryUi();
  } catch (error) {
    console.warn(error);

    if (loadId === activeLoadId) {
      updateEntryUi();
    }
  }
}

async function loadRemainingSheets(loadId, loadedSheetName) {
  try {
    const remainingSheets = FALLBACK_SHEETS.filter((sheet) => sheet.name !== loadedSheetName);
    const parsed = await loadFromPublishedSheets(remainingSheets);

    if (loadId !== activeLoadId) {
      return;
    }

    const merged = mergeParsedSheets([parseSheetBundle(state.rows, state.meta), parsed]);
    state.rows = merged.rows.sort((a, b) => b.dateValue - a.dateValue);
    state.meta = merged.meta;

    buildFilterOptions();
    render();
    setStatus("ready", `Datos completos: ${formatDateTime(new Date())}`);
    renderSheetMeta();
    updateEntryUi();
  } catch (error) {
    console.warn(error);
    if (loadId === activeLoadId) {
      setStatus("ready", `Mes actual cargado: ${formatDateTime(new Date())}`);
      els.sheetMeta.textContent = "No se pudieron completar las hojas historicas.";
    }
  }
}

function parseSheetBundle(rows, meta) {
  return {
    rows: [...rows],
    meta: {
      sheetTotal: meta.sheetTotal || 0,
      dolarBlue: meta.dolarBlue || 0,
    },
  };
}

async function loadFromPublishedSheets(sheets = FALLBACK_SHEETS) {
  const parsedSheets = await Promise.all(
    sheets.map((sheet) => loadPublishedSheet(sheet)),
  );

  return mergeParsedSheets(parsedSheets);
}

async function loadPublishedSheet(sheet) {
  const response = await fetch(`${sheet.url}&cacheBust=${Date.now()}`);

  if (!response.ok) {
    throw new Error(`Google Sheets respondio con estado ${response.status}`);
  }

  return parsePublishedSheet(await response.text(), sheet.name);
}

function getCurrentSheet() {
  const name = sheetNameForDate(new Date());
  return FALLBACK_SHEETS.find((sheet) => sheet.name === name) || FALLBACK_SHEETS[FALLBACK_SHEETS.length - 1];
}

function sheetNameForDate(date) {
  const months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const month = months[date.getMonth()];
  const year = date.getFullYear();

  return year === 2025 ? month : `${month} ${String(year).slice(-2)}`;
}

async function loadFromAppsScript(apiUrl) {
  const data = await jsonp(apiUrl, { action: "list", cacheBust: Date.now() });

  if (!data.ok) {
    throw new Error(data.error || "Apps Script no devolvio una respuesta valida.");
  }

  state.backendVersion = Number(data.version || 0);
  const parsedSheets = (data.sheets || [])
    .filter((sheet) => Array.isArray(sheet.values))
    .map((sheet) => parseSheetMatrix(sheet.values, sheet.name));

  return mergeParsedSheets(parsedSheets);
}

function parsePublishedSheet(csv, sourceSheet) {
  return parseSheetMatrix(parseCsv(csv), sourceSheet);
}

function parseSheetMatrix(matrix, sourceSheet = "") {
  const headerIndex = matrix.findIndex((row) => {
    const normalized = row.map(normalize);
    return normalized.some((cell) => cell.includes("categoria")) && normalized.some((cell) => cell.includes("descripcion")) && normalized.some((cell) => cell.includes("metodo"));
  });

  if (headerIndex === -1) {
    console.warn(`No se encontro la fila de encabezados en ${sourceSheet || "una hoja"}.`);
    return { rows: [], meta: {} };
  }

  const meta = parseMetaRows(matrix.slice(0, headerIndex + 1));
  const columnMap = getColumnMap(matrix[headerIndex]);
  const rows = matrix
    .slice(headerIndex + 1)
    .map((row, index) => normalizeDataRow(row, index, sourceSheet, columnMap))
    .filter(Boolean);

  return { rows, meta };
}

function mergeParsedSheets(parsedSheets) {
  return parsedSheets.reduce(
    (merged, parsed) => {
      merged.rows.push(...parsed.rows);
      merged.meta.sheetTotal += parsed.meta.sheetTotal || 0;
      merged.meta.dolarBlue = merged.meta.dolarBlue || parsed.meta.dolarBlue || 0;
      return merged;
    },
    { rows: [], meta: { sheetTotal: 0, dolarBlue: 0 } },
  );
}

function parseMetaRows(rows) {
  const meta = {};

  rows.flat().forEach((cell, index, allCells) => {
    const label = normalize(cell);

    if (label.includes("dolar blue")) {
      meta.dolarBlue = meta.dolarBlue || parseMoney(cell) || parseMoney(allCells[index + 1]);
    }

    if (label.startsWith("total")) {
      meta.sheetTotal = parseMoney(allCells[index + 1]);
    }
  });

  return meta;
}

function getColumnMap(headerRow) {
  const normalized = headerRow.map(normalize);
  const category = findColumn(normalized, "categoria", 1);
  const description = findColumn(normalized, "descripcion", category + 1);

  return {
    date: findColumn(normalized, "fecha", 0),
    category,
    description,
    amount: findColumn(normalized, "monto", description + 1),
    method: findColumn(normalized, "metodo", description + 2),
    notes: findColumn(normalized, "observaciones", description + 4),
  };
}

function findColumn(row, term, fallback) {
  const index = row.findIndex((cell) => cell.includes(term));
  return index >= 0 ? index : fallback;
}

function normalizeDataRow(row, index, sourceSheet = "", columnMap = {}) {
  const rawDate = clean(row[columnMap.date ?? 0]);
  const category = titleCase(clean(row[columnMap.category ?? 1]));
  const description = clean(row[columnMap.description ?? 2]);
  const amount = parseMoney(row[columnMap.amount ?? 3]);
  const method = titleCase(clean(row[columnMap.method ?? 4]));
  const notes = clean(row[columnMap.notes ?? 6]);

  if (!rawDate && !category && !description && !amount && !method && !notes) {
    return null;
  }

  const parsedDate = parseSheetDate(rawDate, sourceSheet);

  return {
    id: `${rawDate}-${category}-${description}-${amount}-${index}`,
    sourceSheet: sourceSheet || "Hoja",
    rawDate,
    dateValue: parsedDate || new Date(0),
    monthKey: parsedDate ? `${parsedDate.getFullYear()}-${String(parsedDate.getMonth() + 1).padStart(2, "0")}` : "sin-fecha",
    monthLabel: parsedDate ? parsedDate.toLocaleDateString("es-AR", { month: "long", year: "numeric" }) : "Sin fecha",
    dateKey: parsedDate ? parsedDate.toISOString().slice(0, 10) : rawDate,
    dayLabel: parsedDate ? parsedDate.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit" }) : rawDate,
    category: category || "Sin categoria",
    description: description || "Sin descripcion",
    amount,
    method: method || "Sin metodo",
    notes,
    searchable: [sourceSheet, rawDate, category, description, method, notes].join(" ").toLowerCase(),
  };
}

function buildFilterOptions() {
  const current = { ...state.filters };
  const months = uniqueBy(
    state.rows.map((row) => ({ value: row.monthKey, label: row.monthLabel, sort: row.monthKey })),
    "value",
  ).sort((a, b) => {
    if (a.value === "sin-fecha") return 1;
    if (b.value === "sin-fecha") return -1;
    return b.sort.localeCompare(a.sort);
  });
  const categories = uniqueSorted(state.rows.map((row) => row.category));
  const methods = uniqueSorted(state.rows.map((row) => row.method));
  const currentMonth = currentMonthKey();
  const selectedMonth = current.month === "all" && months.some((month) => month.value === currentMonth) ? currentMonth : current.month;

  fillSelect(els.monthFilter, [{ value: "all", label: "Todos los meses" }, ...months], selectedMonth);
  fillSelect(els.categoryFilter, [{ value: "all", label: "Todas" }, ...categories.map(toOption)], current.category);
  fillSelect(els.methodFilter, [{ value: "all", label: "Todos" }, ...methods.map(toOption)], current.method);
  els.categoryOptions.innerHTML = categories.map((category) => `<option value="${escapeHtml(category)}"></option>`).join("");

  state.filters.month = els.monthFilter.value;
  state.filters.category = els.categoryFilter.value;
  state.filters.method = els.methodFilter.value;
}

function render() {
  const filtered = getFilteredRows();
  renderKpis(filtered);
  renderCategoryChart(filtered);
  renderDailyChart(filtered);
  renderTable(filtered);
}

function getFilteredRows() {
  return state.rows.filter((row) => {
    const matchesMonth = state.filters.month === "all" || row.monthKey === state.filters.month;
    const matchesCategory = state.filters.category === "all" || row.category === state.filters.category;
    const matchesMethod = state.filters.method === "all" || row.method === state.filters.method;
    const matchesSearch = !state.filters.search || row.searchable.includes(state.filters.search);

    return matchesMonth && matchesCategory && matchesMethod && matchesSearch;
  });
}

function renderKpis(rows) {
  const paidRows = rows.filter((row) => row.amount > 0);
  const total = sum(paidRows, "amount");
  const cash = sum(paidRows.filter((row) => normalize(row.method).includes("efectivo")), "amount");
  const transfer = sum(paidRows.filter((row) => normalize(row.method).includes("transferencia")), "amount");

  els.totalRevenue.textContent = formatMoney(total);
  els.jobCount.textContent = paidRows.length.toLocaleString("es-AR");
  els.avgTicket.textContent = formatMoney(paidRows.length ? total / paidRows.length : 0);
  els.cashTotal.textContent = formatMoney(cash);
  els.transferTotal.textContent = formatMoney(transfer);
}

function renderCategoryChart(rows) {
  const totals = groupTotals(rows.filter((row) => row.amount > 0), "category").slice(0, 8);
  const max = Math.max(...totals.map((item) => item.total), 1);

  els.categoryTotal.textContent = `${totals.length} categorias`;
  els.categoryChart.innerHTML = totals.length
    ? totals
        .map((item) => {
          const width = Math.max((item.total / max) * 100, 4);
          return `
            <div class="bar-row">
              <strong>${escapeHtml(item.label)}</strong>
              <div class="bar-track" aria-hidden="true"><div class="bar-fill" style="width:${width}%"></div></div>
              <span class="bar-value">${formatMoney(item.total)}</span>
            </div>
          `;
        })
        .join("")
    : `<div class="empty-state">No hay movimientos para estos filtros.</div>`;
}

function renderDailyChart(rows) {
  const totals = groupDailyTotals(rows.filter((row) => row.amount > 0));
  const max = Math.max(...totals.map((item) => item.total), 1);

  els.dailyTotal.textContent = `${totals.length} dias`;
  els.dailyChart.innerHTML = totals.length
    ? totals
        .map((item) => {
          const height = Math.max((item.total / max) * 100, 6);
          return `<div class="day-bar" style="height:${height}%" title="${escapeHtml(item.label)} - ${formatMoney(item.total)}"><span>${escapeHtml(item.label)}</span></div>`;
        })
        .join("")
    : `<div class="empty-state">Sin datos diarios.</div>`;
}

function renderTable(rows) {
  els.visibleRows.textContent = `${rows.length.toLocaleString("es-AR")} registros`;

  if (!rows.length) {
    els.rowsTable.innerHTML = `<tr><td colspan="7" class="empty-state">No hay registros para mostrar.</td></tr>`;
    return;
  }

  els.rowsTable.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.rawDate)}</td>
          <td>${escapeHtml(row.sourceSheet)}</td>
          <td><span class="pill">${escapeHtml(row.category)}</span></td>
          <td>${escapeHtml(row.description)}</td>
          <td>${escapeHtml(row.method)}</td>
          <td class="amount-cell">${row.amount ? formatMoney(row.amount) : "-"}</td>
          <td>${escapeHtml(row.notes || "-")}</td>
        </tr>
      `,
    )
    .join("");
}

function renderSheetMeta() {
  const bits = [];

  if (state.meta.sheetTotal) {
    bits.push(`Total en hoja: ${formatMoney(state.meta.sheetTotal)}`);
  }

  if (state.meta.dolarBlue) {
    bits.push(`Dolar blue: ${formatMoney(state.meta.dolarBlue, 0)}`);
  }

  bits.push(state.source === "apps-script" ? "Todas las hojas" : "Hojas publicas");

  els.sheetMeta.textContent = bits.join(" - ");
}

function jsonp(url, params = {}) {
  return new Promise((resolve, reject) => {
    const callbackName = `emmeCallback${Date.now()}${Math.round(Math.random() * 100000)}`;
    const script = document.createElement("script");
    const target = new URL(url);

    Object.entries({ ...params, callback: callbackName }).forEach(([key, value]) => {
      target.searchParams.set(key, value);
    });

    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Apps Script tardo demasiado en responder."));
    }, 15000);

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("No se pudo conectar con Apps Script."));
    };

    script.src = target.toString();
    document.body.appendChild(script);
  });
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && quoted && next === '"') {
      cell += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
}

function parseSheetDate(value, sourceSheet = "") {
  const cleanValue = clean(value).replace(/\s+/g, "");
  const match = cleanValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);

  if (!match) {
    return null;
  }

  const first = Number(match[1]);
  const second = Number(match[2]);
  const rawYear = Number(match[3]);
  const year = rawYear < 100 ? 2000 + rawYear : rawYear;
  const sheetMonth = monthNumberFromSheetName(sourceSheet);
  let day = first;
  let month = second - 1;

  if (sheetMonth && first === sheetMonth && second !== sheetMonth) {
    day = second;
    month = first - 1;
  } else if (sheetMonth && second === sheetMonth) {
    day = first;
    month = second - 1;
  }

  const date = new Date(year, month, day);

  return Number.isNaN(date.getTime()) ? null : date;
}

function monthNumberFromSheetName(sourceSheet) {
  const months = {
    enero: 1,
    febrero: 2,
    marzo: 3,
    abril: 4,
    mayo: 5,
    junio: 6,
    julio: 7,
    agosto: 8,
    septiembre: 9,
    octubre: 10,
    noviembre: 11,
    diciembre: 12,
  };
  const key = normalize(sourceSheet).split(" ")[0];

  return months[key] || 0;
}

function parseMoney(value) {
  if (value === undefined || value === null) {
    return 0;
  }

  const normalized = String(value).replace(/[^\d,.-]/g, "").replace(/,/g, "");
  const parsed = Number(normalized);

  return Number.isFinite(parsed) ? parsed : 0;
}

function groupTotals(rows, field) {
  const map = new Map();

  rows.forEach((row) => {
    const key = row[field] || "Sin dato";
    map.set(key, (map.get(key) || 0) + row.amount);
  });

  return [...map.entries()]
    .map(([label, total]) => ({ label, total }))
    .sort((a, b) => b.total - a.total);
}

function groupDailyTotals(rows) {
  const map = new Map();

  rows.forEach((row) => {
    if (!map.has(row.dateKey)) {
      map.set(row.dateKey, { label: row.dayLabel, total: 0, sort: row.dateKey });
    }
    map.get(row.dateKey).total += row.amount;
  });

  return [...map.values()].sort((a, b) => a.sort.localeCompare(b.sort));
}

function fillSelect(select, options, selectedValue) {
  select.innerHTML = options.map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`).join("");
  select.value = options.some((option) => option.value === selectedValue) ? selectedValue : "all";
}

function setStatus(type, message) {
  els.statusDot.className = `status-dot ${type === "ready" ? "ready" : type === "error" ? "error" : ""}`;
  els.statusText.textContent = message;
}

function sum(rows, field) {
  return rows.reduce((total, row) => total + (Number(row[field]) || 0), 0);
}

function uniqueSorted(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
}

function uniqueBy(items, key) {
  const seen = new Set();
  return items.filter((item) => {
    if (seen.has(item[key])) {
      return false;
    }
    seen.add(item[key]);
    return true;
  });
}

function toOption(value) {
  return { value, label: value };
}

function clean(value) {
  return String(value ?? "").trim();
}

function normalize(value) {
  return clean(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function titleCase(value) {
  return clean(value)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/(^|\s)([a-zA-Z\u00C0-\u017F])/g, (_, space, letter) => `${space}${letter.toUpperCase()}`);
}

function formatMoney(value, digits = 0) {
  return new Intl.NumberFormat("es-AR", {
    style: "currency",
    currency: "ARS",
    maximumFractionDigits: digits,
  }).format(value || 0);
}

function formatDateTime(date) {
  return date.toLocaleString("es-AR", {
    dateStyle: "short",
    timeStyle: "short",
  });
}

function todayInputValue() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function currentMonthKey() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return clean(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
