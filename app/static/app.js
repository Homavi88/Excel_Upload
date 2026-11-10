const state = {
  sheets: [],
  activeSheet: null,
  upload: null,
  mapping: {},
  keys: new Set(),
  library: { books: [], last_id: null },
  workbookReady: false,
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

function updateFileLabel(data) {
  const label = data.file || "Файл не выбран";
  if (data.file && data.inplace) {
    $("fileLabel").innerHTML = `${escapeHtml(label)} <span class="badge">в оригинал</span>`;
  } else if (data.file) {
    $("fileLabel").innerHTML = `${escapeHtml(label)} <span class="badge warn">копия</span>`;
  } else {
    $("fileLabel").textContent = label;
  }
  const pathEl = $("filePath");
  if (data.path) {
    pathEl.hidden = false;
    pathEl.textContent = data.path;
    pathEl.title = data.path;
  } else {
    pathEl.hidden = true;
    pathEl.textContent = "";
  }
  $("downloadBtn").textContent = data.file ? `Скачать ${data.file}` : "Скачать";
  $("sheetsHint").textContent = data.file
    ? `Выберите целевой лист в ${data.file}`
    : "Сначала откройте Excel-файл через «Открыть Excel»";
}

async function loadWorkbook() {
  const data = await api("/api/workbook");
  state.sheets = data.sheets || [];
  state.inplace = !!data.inplace;
  updateFileLabel(data);
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

async function browseWorkbook() {
  $("fileLabel").textContent = "Открытие проводника…";
  const data = await api("/api/workbook/browse", { method: "POST" });
  state.activeSheet = null;
  state.mapping = {};
  state.keys = new Set();
  state.workbookReady = true;
  hideBookModal();
  setResult(
    `Открыт оригинал: ${data.file}\nПуть: ${data.path}\nЛистов: ${data.sheets.length}\nИзменения будут сохраняться в этот файл.`
  );
  await refreshLibrary();
  await loadWorkbook();
}

async function refreshLibrary() {
  state.library = await api("/api/library");
  renderBookModalList();
  renderSettingsList();
  return state.library;
}

function showBookModal() {
  $("bookModal").hidden = false;
  renderBookModalList();
}

function hideBookModal() {
  $("bookModal").hidden = true;
}

function showSettingsModal() {
  $("settingsModal").hidden = false;
  renderSettingsList();
}

function hideSettingsModal() {
  $("settingsModal").hidden = true;
}

function renderBookModalList() {
  const root = $("bookModalList");
  const books = state.library.books || [];
  root.innerHTML = "";
  if (!books.length) {
    root.innerHTML = `<div class="book-empty">Список пуст. Добавьте первую Excel-книгу.</div>`;
    return;
  }
  for (const book of books) {
    const row = document.createElement("div");
    row.className = "book-item" + (book.exists ? "" : " missing");
    const pick = document.createElement("button");
    pick.type = "button";
    pick.className = "pick";
    pick.innerHTML = `<strong>${escapeHtml(book.name)}</strong><span>${escapeHtml(book.path)}${book.exists ? "" : " · файл не найден"}</span>`;
    pick.disabled = !book.exists;
    pick.addEventListener("click", () => selectLibraryBook(book.id).catch((e) => setResult(e.message, true)));
    row.appendChild(pick);
    root.appendChild(row);
  }
}

function renderSettingsList() {
  const root = $("settingsBookList");
  const books = state.library.books || [];
  root.innerHTML = "";
  if (!books.length) {
    root.innerHTML = `<div class="book-empty">Пока нет сохранённых книг.</div>`;
    return;
  }
  for (const book of books) {
    const row = document.createElement("div");
    row.className = "book-item" + (book.exists ? "" : " missing");
    const info = document.createElement("div");
    info.innerHTML = `<strong>${escapeHtml(book.name)}</strong><span>${escapeHtml(book.path)}${book.exists ? "" : " · файл не найден"}</span>`;
    const actions = document.createElement("div");
    actions.className = "book-actions";

    const openBtn = document.createElement("button");
    openBtn.type = "button";
    openBtn.className = "btn primary";
    openBtn.textContent = "Открыть";
    openBtn.disabled = !book.exists;
    openBtn.addEventListener("click", () => selectLibraryBook(book.id).catch((e) => setResult(e.message, true)));

    const renameBtn = document.createElement("button");
    renameBtn.type = "button";
    renameBtn.className = "btn ghost";
    renameBtn.textContent = "Имя";
    renameBtn.addEventListener("click", async () => {
      const next = prompt("Новое название книги", book.name);
      if (!next) return;
      try {
        const data = await api(`/api/library/${book.id}/rename`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name: next }),
        });
        state.library = data.library;
        renderBookModalList();
        renderSettingsList();
      } catch (err) {
        setResult(err.message, true);
      }
    });

    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "btn danger";
    delBtn.textContent = "Удалить";
    delBtn.addEventListener("click", async () => {
      if (!confirm(`Убрать «${book.name}» из списка?`)) return;
      try {
        state.library = await api(`/api/library/${book.id}`, { method: "DELETE" });
        renderBookModalList();
        renderSettingsList();
      } catch (err) {
        setResult(err.message, true);
      }
    });

    actions.append(openBtn, renameBtn, delBtn);
    row.append(info, actions);
    root.appendChild(row);
  }
}

async function selectLibraryBook(bookId) {
  $("fileLabel").textContent = "Открытие…";
  const data = await api(`/api/library/select/${bookId}`, { method: "POST" });
  state.library = data.library || state.library;
  state.activeSheet = null;
  state.mapping = {};
  state.keys = new Set();
  state.workbookReady = true;
  hideBookModal();
  hideSettingsModal();
  setResult(
    `Выбрана книга: ${data.library_name || data.file}\nПуть: ${data.path}\nЛистов: ${data.sheets.length}`
  );
  await loadWorkbook();
}

async function addBookViaBrowse() {
  const data = await api("/api/library/add-browse", { method: "POST" });
  state.library = data.library;
  if (data.opened) {
    state.activeSheet = null;
    state.mapping = {};
    state.keys = new Set();
    state.workbookReady = true;
    hideBookModal();
    setResult(
      `Добавлена и открыта книга: ${data.opened.library_name || data.opened.file}\nПуть: ${data.opened.path}`
    );
    await loadWorkbook();
  } else {
    renderBookModalList();
    renderSettingsList();
  }
}

async function shutdownApp() {
  if (!confirm("Выключить Excel Loader и закрыть окно терминала?")) return;
  $("shutdownBtn").disabled = true;
  setResult("Выключение приложения…");
  try {
    await api("/api/shutdown", { method: "POST" });
  } catch (_) {
    // Server may die before response is fully received — expected.
  }
  document.body.innerHTML = `
    <div style="min-height:100vh;display:grid;place-items:center;font-family:IBM Plex Sans,sans-serif;background:#eef3f7;color:#173042;">
      <div style="text-align:center;padding:24px;">
        <h1 style="font-family:Space Grotesk,sans-serif;margin:0 0 8px;">Приложение выключено</h1>
        <p style="margin:0;color:#5f7383;">Окно терминала должно закрыться автоматически.</p>
      </div>
    </div>`;
}

async function openWorkbookFile(file) {
  const form = new FormData();
  form.append("file", file);
  $("fileLabel").textContent = "Открытие…";
  const data = await api("/api/workbook/open", { method: "POST", body: form });
  state.activeSheet = null;
  state.mapping = {};
  state.keys = new Set();
  setResult(
    `Открыта копия: ${data.file}\nЛистов: ${data.sheets.length}\nЧтобы писать в оригинал, нажмите «Открыть Excel» (проводнику Windows).`
  );
  await loadWorkbook();
}

function renderSheets() {
  const root = $("sheetList");
  root.innerHTML = "";
  const canDelete = state.sheets.length > 1;
  for (const sheet of state.sheets) {
    const row = document.createElement("div");
    row.className = "sheet-row" + (sheet.name === state.activeSheet ? " active" : "");

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "sheet-btn" + (sheet.name === state.activeSheet ? " active" : "");
    btn.innerHTML = `<strong>${escapeHtml(sheet.name)}</strong><span>${sheet.row_count} строк · ${sheet.columns.length} кол.</span>`;
    btn.addEventListener("click", () => selectSheet(sheet.name));

    const del = document.createElement("button");
    del.type = "button";
    del.className = "btn danger sheet-del";
    del.textContent = "Удалить";
    del.title = canDelete ? "Удалить лист из книги" : "Нельзя удалить последний лист";
    del.disabled = !canDelete;
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteSheet(sheet.name).catch((err) => setResult(err.message, true));
    });

    row.append(btn, del);
    root.appendChild(row);
  }
}

async function deleteSheet(name) {
  if (!confirm(`Удалить лист «${name}» из книги?\nПеред удалением будет создан backup.`)) {
    return;
  }
  const result = await api("/api/sheets/delete", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name }),
  });
  setResult(
    `Лист удалён: ${result.sheet}\nСохранено в: ${result.saved_to || "—"}\nБэкап: ${result.backup}`
  );
  if (state.activeSheet === name) {
    state.activeSheet = null;
  }
  await loadWorkbook();
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
  $("browseBtn").addEventListener("click", () => browseWorkbook().catch((e) => setResult(e.message, true)));
  $("chooseBookBtn").addEventListener("click", () => showBookModal());
  $("settingsBtn").addEventListener("click", () => showSettingsModal());
  $("settingsCloseBtn").addEventListener("click", () => hideSettingsModal());
  $("bookModalAddBtn").addEventListener("click", () => addBookViaBrowse().catch((e) => setResult(e.message, true)));
  $("bookModalSettingsBtn").addEventListener("click", () => showSettingsModal());
  $("settingsAddBtn").addEventListener("click", () => addBookViaBrowse().catch((e) => setResult(e.message, true)));
  $("shutdownBtn").addEventListener("click", () => shutdownApp().catch((e) => setResult(e.message, true)));
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
      `Лист создан: ${result.sheet}\nКолонки: ${result.columns.join(", ")}\nСтрок записано: ${result.inserted}\nСохранено в: ${result.saved_to || "—"}\nБэкап: ${result.backup}`
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

async function boot() {
  wireUi();
  await refreshLibrary();
  showBookModal();
  // Preload default workbook info in background, but keep modal until user chooses.
  try {
    await loadWorkbook();
  } catch (err) {
    // No active workbook yet is fine on first launch.
    console.warn(err);
  }
}

boot().catch((err) => setResult(err.message, true));
