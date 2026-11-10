const state = {
  sheets: [],
  activeSheet: null,
  upload: null,
  mapping: {},
  keys: new Set(),
};

const $ = (id) => document.getElementById(id);

async function api(url, options) {
  const res = await fetch(url, options);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.detail || res.statusText || "Ошибка запроса");
  }
  return data;
}

function setResult(text, isError = false) {
  const box = $("resultBox");
  box.hidden = false;
  box.classList.toggle("error", isError);
  box.textContent = typeof text === "string" ? text : JSON.stringify(text, null, 2);
}

async function loadWorkbook() {
  const data = await api("/api/workbook");
  state.sheets = data.sheets || [];
  $("fileLabel").textContent = data.file || "Файл не выбран";
  $("downloadBtn").textContent = data.file ? `Скачать ${data.file}` : "Скачать";
  $("sheetsHint").textContent = data.file
    ? `Выберите целевой лист в ${data.file}`
    : "Сначала откройте Excel-файл через «Открыть Excel»";
  renderSheets();
  if (!data.exists) {
    $("activeSheetTitle").textContent = "Загрузка данных";
    $("activeSheetMeta").textContent = "Откройте Excel-файл в шапке страницы";
    $("previewMeta").textContent = "—";
    $("previewTable").querySelector("thead").innerHTML = "";
    $("previewTable").querySelector("tbody").innerHTML = "";
    state.activeSheet = null;
    renderKeyChips();
    renderMapping();
    maybeEnableApply();
    await loadBackups();
    return;
  }
  if (!state.activeSheet && data.sheets.length) {
    selectSheet(data.sheets[0].name);
  } else if (state.activeSheet && data.sheets.some((s) => s.name === state.activeSheet)) {
    selectSheet(state.activeSheet);
  } else if (data.sheets.length) {
    selectSheet(data.sheets[0].name);
  }
  await loadBackups();
}

async function openWorkbookFile(file) {
  const form = new FormData();
  form.append("file", file);
  $("fileLabel").textContent = "Открытие…";
  const data = await api("/api/workbook/open", { method: "POST", body: form });
  state.activeSheet = null;
  state.mapping = {};
  state.keys = new Set();
  setResult(`Открыт файл: ${data.file}\nЛистов: ${data.sheets.length}`);
  await loadWorkbook();
}

function renderSheets() {
  const root = $("sheetList");
  root.innerHTML = "";
  for (const sheet of state.sheets) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "sheet-btn" + (sheet.name === state.activeSheet ? " active" : "");
    btn.innerHTML = `<strong>${escapeHtml(sheet.name)}</strong><span>${sheet.row_count} строк · ${sheet.columns.length} кол.</span>`;
    btn.addEventListener("click", () => selectSheet(sheet.name));
    root.appendChild(btn);
  }
}

async function selectSheet(name) {
  state.activeSheet = name;
  const sheet = state.sheets.find((s) => s.name === name);
  renderSheets();
  $("activeSheetTitle").textContent = name;
  $("activeSheetMeta").textContent = sheet
    ? `Заголовок: строка ${sheet.header_row} · ${sheet.row_count} строк данных · ${sheet.columns.length} колонок`
    : "";
  renderKeyChips();
  renderMapping();
  await loadPreview(name);
  maybeEnableApply();
}

async function loadPreview(name) {
  const data = await api(`/api/sheets/${encodeURIComponent(name)}/preview?limit=40`);
  $("previewMeta").textContent = `Показано ${data.shown} из ~${data.total_estimate}`;
  const thead = $("previewTable").querySelector("thead");
  const tbody = $("previewTable").querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";
  const trh = document.createElement("tr");
  for (const col of data.columns) {
    const th = document.createElement("th");
    th.textContent = col;
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  for (const row of data.rows) {
    const tr = document.createElement("tr");
    for (const col of data.columns) {
      const td = document.createElement("td");
      const val = row[col];
      td.textContent = val == null ? "" : String(val);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}

function currentTargetColumns() {
  const sheet = state.sheets.find((s) => s.name === state.activeSheet);
  return sheet ? sheet.columns.map((c) => c.name) : [];
}

function currentSourceColumns() {
  if (!state.upload) return [];
  const src = state.upload.sheets.find((s) => s.name === $("sourceSheet").value) || state.upload.sheets[0];
  return src ? src.columns : [];
}

function renderKeyChips() {
  const root = $("keyColumns");
  root.innerHTML = "";
  const cols = currentTargetColumns();
  const valid = new Set(cols);
  state.keys = new Set([...state.keys].filter((k) => valid.has(k)));
  for (const col of cols) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip" + (state.keys.has(col) ? " on" : "");
    chip.textContent = col;
    chip.addEventListener("click", () => {
      if (state.keys.has(col)) state.keys.delete(col);
      else state.keys.add(col);
      chip.classList.toggle("on");
    });
    root.appendChild(chip);
  }
}

function renderMapping() {
  const root = $("mappingTable");
  root.innerHTML = "";
  const targetCols = currentTargetColumns();
  const sourceCols = currentSourceColumns();
  if (!targetCols.length) {
    root.innerHTML = `<p class="muted">Выберите целевой лист</p>`;
    return;
  }
  if (!sourceCols.length) {
    root.innerHTML = `<p class="muted">Загрузите файл-источник</p>`;
    return;
  }

  for (const target of targetCols) {
    const row = document.createElement("div");
    row.className = "map-row";
    const left = document.createElement("div");
    left.className = "target";
    left.textContent = target;
    const arrow = document.createElement("div");
    arrow.className = "arrow";
    arrow.textContent = "←";
    const select = document.createElement("select");
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "— пропустить —";
    select.appendChild(empty);
    for (const src of sourceCols) {
      const opt = document.createElement("option");
      opt.value = src;
      opt.textContent = src;
      select.appendChild(opt);
    }
    select.value = state.mapping[target] || "";
    select.addEventListener("change", () => {
      state.mapping[target] = select.value || null;
      maybeEnableApply();
    });
    row.append(left, arrow, select);
    root.appendChild(row);
  }
}

async function autoMap() {
  if (!state.upload || !state.activeSheet) return;
  const data = await api("/api/suggest-mapping", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      source_columns: currentSourceColumns(),
      target_columns: currentTargetColumns(),
    }),
  });
  state.mapping = data.mapping;
  renderMapping();
  maybeEnableApply();
}

function maybeEnableApply() {
  const hasMapped = Object.values(state.mapping).some(Boolean);
  $("applyBtn").disabled = !(state.upload && state.activeSheet && hasMapped);
}

async function handleUpload(file) {
  const form = new FormData();
  form.append("file", file);
  $("uploadInfo").textContent = "Загрузка…";
  const data = await api("/api/upload", { method: "POST", body: form });
  state.upload = data;
  $("uploadInfo").textContent = `${data.filename} · ${data.sheets.length} лист(ов)`;
  const wrap = $("sourceSheetWrap");
  const select = $("sourceSheet");
  select.innerHTML = "";
  for (const sheet of data.sheets) {
    const opt = document.createElement("option");
    opt.value = sheet.name;
    opt.textContent = `${sheet.name} (${sheet.row_count})`;
    select.appendChild(opt);
  }
  wrap.hidden = data.sheets.length < 1;
  await autoMap();
}

async function applyLoad() {
  if (!state.upload || !state.activeSheet) return;
  const mode = document.querySelector('input[name="mode"]:checked').value;
  if (mode === "upsert" && state.keys.size === 0) {
    setResult("Для режима «Обновить / добавить» выберите хотя бы одну ключевую колонку.", true);
    return;
  }
  $("applyBtn").disabled = true;
  try {
    const result = await api("/api/apply", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        upload_id: state.upload.upload_id,
        source_sheet: $("sourceSheet").value || null,
        target_sheet: state.activeSheet,
        mode,
        mapping: state.mapping,
        key_columns: [...state.keys],
      }),
    });
    setResult(
      `Готово.\nРежим: ${result.mode}\nОбновлено: ${result.updated}\nДобавлено: ${result.inserted}\nОчищено ячеек: ${result.cleared}\nБэкап: ${result.backup}`
    );
    await loadWorkbook();
  } catch (err) {
    setResult(err.message, true);
  } finally {
    maybeEnableApply();
  }
}

async function loadBackups() {
  const data = await api("/api/backups");
  const root = $("backupList");
  root.innerHTML = "";
  if (!data.backups.length) {
    root.innerHTML = `<p class="muted">Пока нет бэкапов</p>`;
    return;
  }
  for (const item of data.backups) {
    const row = document.createElement("div");
    row.className = "backup-item";
    const when = new Date(item.mtime * 1000).toLocaleString("ru-RU", { timeZone: "Europe/Moscow" });
    row.innerHTML = `<div><strong>${escapeHtml(item.name)}</strong><br /><span>${when} MSK</span></div>`;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn danger";
    btn.textContent = "Восстановить";
    btn.addEventListener("click", async () => {
      if (!confirm(`Восстановить текущий Excel из ${item.name}?`)) return;
      try {
        await api("/api/restore", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ backup: item.name }),
        });
        setResult(`Восстановлено из ${item.name}`);
        await loadWorkbook();
      } catch (err) {
        setResult(err.message, true);
      }
    });
    row.appendChild(btn);
    root.appendChild(row);
  }
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function wireUi() {
  const drop = $("dropZone");
  const input = $("fileInput");
  drop.addEventListener("click", () => input.click());
  drop.addEventListener("dragover", (e) => {
    e.preventDefault();
    drop.classList.add("drag");
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag"));
  drop.addEventListener("drop", async (e) => {
    e.preventDefault();
    drop.classList.remove("drag");
    const file = e.dataTransfer.files?.[0];
    if (file) {
      try {
        await handleUpload(file);
      } catch (err) {
        setResult(err.message, true);
      }
    }
  });
  input.addEventListener("change", async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      await handleUpload(file);
    } catch (err) {
      setResult(err.message, true);
    }
  });

  $("sourceSheet").addEventListener("change", async () => {
    await autoMap();
  });
  $("autoMapBtn").addEventListener("click", () => autoMap().catch((e) => setResult(e.message, true)));
  $("applyBtn").addEventListener("click", () => applyLoad());
  $("refreshBtn").addEventListener("click", () => loadWorkbook().catch((e) => setResult(e.message, true)));
  $("createSheetBtn").addEventListener("click", () => createSheet().catch((e) => setResult(e.message, true)));
  $("workbookInput").addEventListener("change", async () => {
    const file = $("workbookInput").files?.[0];
    if (!file) return;
    try {
      await openWorkbookFile(file);
    } catch (err) {
      setResult(err.message, true);
    } finally {
      $("workbookInput").value = "";
    }
  });
  $("newSheetFromUpload").addEventListener("change", () => {
    const on = $("newSheetFromUpload").checked;
    $("createSheetHint").textContent = on
      ? "Колонки и строки возьмутся из выбранного листа источника. Поле колонок можно оставить пустым."
      : "Можно создать пустой лист с заголовками или импортировать источник как новый лист.";
  });

  document.querySelectorAll('input[name="mode"]').forEach((el) => {
    el.addEventListener("change", () => {
      $("keysWrap").style.opacity = el.value === "upsert" || document.querySelector('input[name="mode"]:checked').value === "upsert" ? "1" : "0.45";
    });
  });
}

async function createSheet() {
  const name = $("newSheetName").value.trim();
  const columnsRaw = $("newSheetColumns").value.trim();
  const fromUpload = $("newSheetFromUpload").checked;
  if (!name) {
    setResult("Укажите название нового листа.", true);
    return;
  }
  if (fromUpload && !state.upload) {
    setResult("Сначала загрузите файл-источник, либо снимите галочку импорта.", true);
    return;
  }
  const columns = columnsRaw
    ? columnsRaw.split(",").map((c) => c.trim()).filter(Boolean)
    : [];
  if (!fromUpload && !columns.length) {
    setResult("Укажите колонки через запятую или включите импорт из загруженного файла.", true);
    return;
  }

  $("createSheetBtn").disabled = true;
  try {
    const result = await api("/api/sheets/create", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        name,
        columns,
        upload_id: fromUpload ? state.upload.upload_id : null,
        source_sheet: fromUpload ? ($("sourceSheet").value || null) : null,
        fill_data: fromUpload,
      }),
    });
    setResult(
      `Лист создан: ${result.sheet}\nКолонки: ${result.columns.join(", ")}\nСтрок записано: ${result.inserted}\nБэкап: ${result.backup}`
    );
    $("newSheetName").value = "";
    $("newSheetColumns").value = "";
    $("newSheetFromUpload").checked = false;
    state.activeSheet = result.sheet;
    await loadWorkbook();
  } catch (err) {
    setResult(err.message, true);
  } finally {
    $("createSheetBtn").disabled = false;
  }
}

wireUi();
loadWorkbook().catch((err) => setResult(err.message, true));
