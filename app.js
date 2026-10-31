// app.js — интерфейс приложения. Склеивает model.js (чистые вычисления), api.js (сеть/демо)
// и queue.js (кеш + очередь на отправку).

import * as model from "./model.js";
import * as api from "./api.js";
import * as queue from "./queue.js";
import * as config from "./config.js";

const HOUR_TILES = [8, 4, 9, 10, 11, 12];

const state = {
  employees: [],
  codes: [],
  normHours: 8,
  monthCache: new Map(), // "YYYY-MM" -> {entries, loadedAt, offline}
  outbox: [],
  currentDate: new Date(),
  screen: "day",
  monthHalf: 1,
  paySystemFilter: "", // "" | "оклад" | "сдельная"
  selectedEmployeeTab: null,
  sheetTarget: null, // {tabNum, dateKey}
  loadError: null, // текст реальной причины, если не удалось получить справочники
};

const el = (id) => document.getElementById(id);

// Демо-режим и обычный режим живут в одной IndexedDB (один origin) — без этого
// префикса демо-данные перезаписывали бы реальный кеш и наоборот (см. DECISIONS.md).
function cacheKey(key) {
  return (config.isDemoMode() ? "demo:" : "real:") + key;
}

function isCurrentModeRecord(record) {
  return !!record.demo === config.isDemoMode();
}

// ---------- запуск ----------

async function main() {
  applyTheme();
  registerServiceWorker();

  if (!config.isConfigured()) {
    showSetupScreen();
    return;
  }

  el("app").classList.remove("hidden");
  wireStaticEvents();

  await loadReferenceData();
  state.outbox = (await queue.getOutbox()).filter(isCurrentModeRecord);
  await ensureMonthLoaded(monthKeyOf(state.currentDate));

  renderAll();
  updateNetworkBanner();

  window.addEventListener("online", onNetworkChange);
  window.addEventListener("offline", onNetworkChange);
  if (navigator.onLine) flushQueue();
}

function showSetupScreen() {
  el("setup-screen").classList.remove("hidden");
  el("setup-url").value = config.getScriptUrl();
  el("setup-pin").value = config.getPin();
  el("setup-save").addEventListener("click", () => {
    const url = el("setup-url").value.trim();
    const pin = el("setup-pin").value.trim();
    if (!url || !pin) {
      showToast("Укажите адрес и PIN");
      return;
    }
    config.setScriptUrl(url);
    config.setPin(pin);
    location.reload();
  });
}

function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
}

// ---------- тема ----------

function applyTheme() {
  const theme = config.getTheme();
  if (theme === "auto") {
    document.documentElement.removeAttribute("data-theme");
  } else {
    document.documentElement.setAttribute("data-theme", theme);
  }
}

// ---------- загрузка данных ----------

async function loadReferenceData() {
  let serverNormHours = state.normHours;
  state.loadError = null;
  try {
    const data = await api.fetchInit();
    state.employees = data.employees;
    state.codes = data.codes;
    serverNormHours = data.normHours;
    await queue.setCache(cacheKey("ref"), data);
  } catch (err) {
    const cached = await queue.getCache(cacheKey("ref"));
    if (cached) {
      state.employees = cached.value.employees;
      state.codes = cached.value.codes;
      serverNormHours = cached.value.normHours;
      state.loadError = `Не удалось обновить с сервера (${err.message || err}) — показаны данные из кеша от ${new Date(cached.savedAt).toLocaleString("ru-RU")}`;
    } else {
      state.loadError = `Не удалось загрузить список сотрудников: ${err.message || err}`;
      showToast(state.loadError);
    }
  }
  // Норма из таблицы — значение по умолчанию; пользователь может переопределить её
  // локально в «Настройках» (влияет только на этот телефон, см. DECISIONS.md).
  state.normHours = config.getNormHoursOverride() || serverNormHours;
}

function monthKeyOf(date) {
  return `${date.getFullYear()}-${model.pad2(date.getMonth() + 1)}`;
}

async function ensureMonthLoaded(monthKey) {
  try {
    const data = await api.fetchMonth(monthKey);
    const entry = { entries: data.entries, loadedAt: new Date().toISOString(), offline: false };
    state.monthCache.set(monthKey, entry);
    await queue.setCache(cacheKey("month:" + monthKey), { entries: data.entries, loadedAt: entry.loadedAt });
  } catch (err) {
    const cached = await queue.getCache(cacheKey("month:" + monthKey));
    if (cached) {
      state.monthCache.set(monthKey, { entries: cached.value.entries, loadedAt: cached.value.loadedAt, offline: true });
    } else {
      state.monthCache.set(monthKey, { entries: [], loadedAt: null, offline: true });
    }
  }
}

// Журнал месяца + несинхронизированные правки (последние по порядку — значит, выигрывают).
function effectiveJournalForMonth(monthKey) {
  const cached = state.monthCache.get(monthKey) || { entries: [] };
  const pending = state.outbox
    .filter((r) => r.date.slice(0, 7) === monthKey)
    .map((r) => ({ date: r.date, tabNum: r.id, value: r.value }));
  return [...cached.entries, ...pending];
}

function activeEmployees() {
  const active = state.employees.filter((e) => e.active);
  return model.filterByPaySystem(active, state.paySystemFilter);
}

// ---------- рендер ----------

function renderAll() {
  renderTopbar();
  updateNetworkBanner();
  if (state.screen === "day") renderDayScreen();
  if (state.screen === "month") renderMonthScreen();
  if (state.screen === "employee") renderEmployeeScreen();
  if (state.screen === "settings") renderSettingsScreen();
  renderSaveBar();
  el("day-chips").classList.toggle("hidden", state.screen !== "day");
  el("pay-filter").classList.toggle("hidden", state.screen !== "day" && state.screen !== "month");
}

const WEEKDAY_NAMES = ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];
const MONTH_NAMES_NOM = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];

function renderTopbar() {
  const titleText = el("topbar-title-text");
  const subtitle = el("topbar-subtitle");
  el("nav-prev").classList.toggle("hidden", state.screen !== "day");
  el("nav-next").classList.toggle("hidden", state.screen !== "day");

  if (state.screen === "day") {
    const d = state.currentDate;
    const dateKey = model.toDateKey(d.getFullYear(), d.getMonth() + 1, d.getDate());
    titleText.textContent = `${WEEKDAY_NAMES[model.weekday(dateKey)]}, ${d.getDate()} ${MONTH_NAMES_NOM[d.getMonth()]}`;
    subtitle.textContent = model.isSunday(dateKey) ? "Выходной" : "";
  } else if (state.screen === "month") {
    titleText.textContent = "Месяц";
    subtitle.textContent = "";
  } else if (state.screen === "employee") {
    titleText.textContent = "Сотрудник";
    subtitle.textContent = "";
  } else {
    titleText.textContent = "Настройки";
    subtitle.textContent = "";
  }
}

function renderDayScreen() {
  const d = state.currentDate;
  const monthKey = monthKeyOf(d);
  const dateKey = model.toDateKey(d.getFullYear(), d.getMonth() + 1, d.getDate());
  const journal = effectiveJournalForMonth(monthKey);
  const emps = activeEmployees();
  const sunday = model.isSunday(dateKey);

  const totals = model.dayTotals(journal, emps, dateKey);
  renderChips(totals);

  el("fill-all-eight").classList.toggle("hidden", sunday);
  el("fill-all-eight").textContent = `Всем ${state.normHours} ч`;

  const list = el("day-list");
  list.innerHTML = "";

  if (emps.length === 0) {
    list.innerHTML = `
      <div class="chip" style="white-space:normal;line-height:1.5">
        Список сотрудников пуст.${state.loadError ? " " + escapeHtml(state.loadError) : ""}
        Проверьте: лист «Сотрудники» в таблице (есть строки, в столбце «Активен» — «да»),
        адрес веб-приложения и PIN в «Настройках», и что веб-приложение Apps Script опубликовано
        с доступом «Все».
      </div>
    `;
    return;
  }

  for (const emp of emps) {
    const value = model.getCellValue(journal, dateKey, emp.tabNum);
    const row = document.createElement("button");
    row.className = "person-row";
    row.innerHTML = `
      <span class="person-row__num">${emp.tabNum}</span>
      <span class="person-row__info">
        <span class="person-row__name">${escapeHtml(emp.name)}</span>
        <span class="person-row__position">${escapeHtml(emp.position)}${emp.paySystem ? " · " + escapeHtml(emp.paySystem) : ""}</span>
      </span>
    `;
    const badge = document.createElement("span");
    if (model.isEmptyValue(value)) {
      if (sunday) {
        badge.className = "value-badge value-badge--suggest";
        badge.style.borderColor = "#B3E5A1";
        badge.textContent = "В";
      } else {
        badge.className = "value-badge value-badge--suggest";
        badge.textContent = String(state.normHours);
      }
    } else {
      badge.className = "value-badge";
      badge.style.background = model.colorForValue(value, state.codes, state.normHours) || "#AFD2D3";
      badge.textContent = String(value);
    }
    row.appendChild(badge);

    if (hasPendingEdit(dateKey, emp.tabNum)) {
      const dot = document.createElement("span");
      dot.className = "pending-dot";
      row.appendChild(dot);
    }

    row.addEventListener("click", () => openSheet(emp.tabNum, dateKey));
    list.appendChild(row);
  }
}

function renderChips(totals) {
  const chips = el("day-chips");
  const byCodeText = Object.entries(totals.byCode)
    .map(([code, count]) => `${escapeHtml(code)} ${count}`)
    .join(" · ");
  chips.innerHTML = `
    <span class="chip">На смене <strong>${totals.onShift}</strong></span>
    ${byCodeText ? `<span class="chip">${byCodeText}</span>` : ""}
    <span class="chip">Не отмечено <strong>${totals.unmarked}</strong></span>
    <span class="chip">Σ <strong>${totals.hoursSum}</strong> ч</span>
  `;
}

function hasPendingEdit(dateKey, tabNum) {
  return state.outbox.some((r) => r.date === dateKey && r.id === tabNum);
}

function renderMonthScreen() {
  const d = state.currentDate;
  const year = d.getFullYear();
  const month = d.getMonth() + 1;
  const monthKey = monthKeyOf(d);
  const journal = effectiveJournalForMonth(monthKey);
  const grid = model.monthGrid(journal, activeEmployees(), year, month);
  const totalDays = model.daysInMonth(year, month);
  const half = state.monthHalf;
  const fromDay = half === 1 ? 1 : 16;
  const toDay = half === 1 ? Math.min(15, totalDays) : totalDays;

  const table = el("month-table");
  let thead = `<tr><th class="name-col">Сотрудник</th>`;
  for (let day = fromDay; day <= toDay; day++) {
    const dateKey = model.toDateKey(year, month, day);
    thead += `<th class="${model.isSunday(dateKey) ? "sunday" : ""}">${day}</th>`;
  }
  thead += `<th>Дни</th><th>Часы</th><th>Перераб.</th></tr>`;

  let rows = "";
  for (const emp of grid) {
    const summary = model.monthSummaryForEmployee(journal, emp.tabNum, year, month, state.normHours);
    rows += `<tr><td class="name-col" data-tab="${emp.tabNum}">${escapeHtml(emp.name)}</td>`;
    for (let day = fromDay; day <= toDay; day++) {
      const dateKey = model.toDateKey(year, month, day);
      const value = emp.values[dateKey];
      const color = model.colorForValue(value, state.codes, state.normHours);
      const sunday = model.isSunday(dateKey) ? "sunday" : "";
      const text = model.isEmptyValue(value) ? "" : String(value);
      rows += `<td class="value-cell ${sunday}" data-tab="${emp.tabNum}" data-date="${dateKey}" style="${color ? `background:${color}` : ""}">${escapeHtml(text)}</td>`;
    }
    rows += `<td>${summary.days}</td><td>${summary.hours}</td><td>${summary.overtime}</td></tr>`;
  }

  table.innerHTML = `<thead>${thead}</thead><tbody>${rows}</tbody>`;

  table.querySelectorAll("td.value-cell").forEach((cell) => {
    cell.addEventListener("click", () => openSheet(Number(cell.dataset.tab), cell.dataset.date));
  });
  table.querySelectorAll("td.name-col").forEach((cell) => {
    cell.addEventListener("click", () => openEmployeeScreen(Number(cell.dataset.tab)));
  });

  const monthTotals = model.monthSummaryTotals(journal, activeEmployees(), year, month, state.normHours);
  el("month-totals").innerHTML = `
    <span>ИТОГО дней: <strong>${monthTotals.days}</strong></span>
    <span>часов: <strong>${monthTotals.hours}</strong></span>
    <span>переработка: <strong>${monthTotals.overtime}</strong></span>
  `;
}

function openEmployeeScreen(tabNum) {
  state.selectedEmployeeTab = tabNum;
  switchScreen("employee");
}

function renderEmployeeScreen() {
  const tabNum = state.selectedEmployeeTab;
  const emp = state.employees.find((e) => e.tabNum === tabNum);
  if (!emp) return;
  const d = state.currentDate;
  const monthKey = monthKeyOf(d);
  const journal = effectiveJournalForMonth(monthKey);
  const summary = model.monthSummaryForEmployee(journal, tabNum, d.getFullYear(), d.getMonth() + 1, state.normHours);

  el("employee-name").textContent = emp.name;
  el("employee-position").textContent = emp.paySystem ? `${emp.position} · ${emp.paySystem}` : emp.position;
  el("employee-days").textContent = summary.days;
  el("employee-hours").textContent = summary.hours;
  el("employee-overtime").textContent = summary.overtime;

  const codesHtml = Object.entries(summary.byCode)
    .map(([code, count]) => {
      const def = state.codes.find((c) => c.code === code);
      const color = def ? def.color : "#ccc";
      return `<span class="chip" style="background:${color};color:#15171a">${escapeHtml(code)}: ${count}</span>`;
    })
    .join("");
  el("employee-codes").innerHTML = codesHtml || `<span class="chip">Нет отметок по кодам</span>`;
}

function renderSettingsScreen() {
  el("settings-url").value = config.getScriptUrl();
  el("settings-pin").value = config.getPin();
  el("settings-norm").value = state.normHours || "";

  document.querySelectorAll("#theme-switch button").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.theme === config.getTheme());
  });

  el("settings-queue-note").textContent = `Не отправлено: ${state.outbox.length}`;
}

function renderSaveBar() {
  const wrap = el("save-bar-wrap");
  if (state.outbox.length === 0) {
    wrap.innerHTML = "";
    return;
  }
  wrap.innerHTML = `<button class="save-bar" id="save-bar-btn">Сохранить (${state.outbox.length}) <span>→</span></button>`;
  el("save-bar-btn").addEventListener("click", () => flushQueue());
}

// ---------- нижняя панель ввода ----------

function openSheet(tabNum, dateKey) {
  const emp = state.employees.find((e) => e.tabNum === tabNum);
  if (!emp) return;
  state.sheetTarget = { tabNum, dateKey };

  const monthKey = dateKey.slice(0, 7);
  const journal = effectiveJournalForMonth(monthKey);
  const currentValue = model.getCellValue(journal, dateKey, tabNum);

  const tiles = [state.normHours, ...HOUR_TILES.filter((h) => h !== state.normHours)];

  const overlay = document.createElement("div");
  overlay.className = "sheet-overlay";
  overlay.innerHTML = `
    <div class="sheet">
      <div class="sheet__title">${escapeHtml(emp.name)} · ${dateKey}</div>
      <div class="sheet__grid">
        ${tiles.map((h) => `<button class="sheet__btn" data-hours="${h}">${h}</button>`).join("")}
      </div>
      <div class="sheet__other">
        <input type="number" step="0.5" min="0" max="24" placeholder="Другое, ч" id="sheet-other-input" />
        <button class="sheet__btn" id="sheet-other-apply">OK</button>
      </div>
      <div class="sheet__codes">
        ${state.codes.map((c) => `<button class="code-pill" style="background:${c.color};color:#15171a" data-code="${escapeHtml(c.code)}">${escapeHtml(c.code)}</button>`).join("")}
      </div>
      <div class="sheet__footer">
        <button class="sheet__clear" id="sheet-clear">Очистить</button>
        <button class="sheet__close" id="sheet-close">Закрыть</button>
      </div>
    </div>
  `;
  el("sheet-root").appendChild(overlay);

  overlay.querySelectorAll(".sheet__btn[data-hours]").forEach((btn) => {
    btn.addEventListener("click", () => applySheetValue(Number(btn.dataset.hours)));
  });
  overlay.querySelector("#sheet-other-apply").addEventListener("click", () => {
    const v = Number(overlay.querySelector("#sheet-other-input").value);
    if (!v && v !== 0) return;
    applySheetValue(v);
  });
  overlay.querySelectorAll(".code-pill").forEach((btn) => {
    btn.addEventListener("click", () => applySheetValue(btn.dataset.code));
  });
  overlay.querySelector("#sheet-clear").addEventListener("click", () => applySheetValue(""));
  overlay.querySelector("#sheet-close").addEventListener("click", closeSheet);
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeSheet();
  });
}

function closeSheet() {
  el("sheet-root").innerHTML = "";
  state.sheetTarget = null;
}

async function applySheetValue(value) {
  const target = state.sheetTarget;
  if (!target) return;
  const { tabNum, dateKey } = target;
  const monthKey = dateKey.slice(0, 7);
  const cached = state.monthCache.get(monthKey) || { entries: [] };
  const baseline = cached.entries
    .filter((e) => e.date === dateKey && e.tabNum === tabNum)
    .pop();

  const existingPending = state.outbox.find((r) => r.date === dateKey && r.id === tabNum);
  if (existingPending) {
    await queue.removeFromOutbox(existingPending.uuid);
    state.outbox = state.outbox.filter((r) => r.uuid !== existingPending.uuid);
  }

  const record = api.makeRecord(dateKey, tabNum, value, baseline ? baseline.ts : undefined);
  await queue.enqueue(record);
  state.outbox.push(record);

  closeSheet();

  const nextTab = findNextUnmarkedTab(dateKey, tabNum);
  renderAll();
  if (nextTab != null && state.screen === "day") {
    openSheet(nextTab, dateKey);
  }
}

function findNextUnmarkedTab(dateKey, afterTabNum) {
  const monthKey = dateKey.slice(0, 7);
  const journal = effectiveJournalForMonth(monthKey);
  const emps = activeEmployees();
  const idx = emps.findIndex((e) => e.tabNum === afterTabNum);
  for (let i = idx + 1; i < emps.length; i++) {
    const value = model.getCellValue(journal, dateKey, emps[i].tabNum);
    if (model.isEmptyValue(value)) return emps[i].tabNum;
  }
  return null;
}

// ---------- действия ----------

async function fillAllEight() {
  const d = state.currentDate;
  const dateKey = model.toDateKey(d.getFullYear(), d.getMonth() + 1, d.getDate());
  if (model.isSunday(dateKey)) return;
  const monthKey = monthKeyOf(d);
  const journal = effectiveJournalForMonth(monthKey);

  for (const emp of activeEmployees()) {
    const value = model.getCellValue(journal, dateKey, emp.tabNum);
    if (!model.isEmptyValue(value)) continue;
    const record = api.makeRecord(dateKey, emp.tabNum, state.normHours, undefined);
    await queue.enqueue(record);
    state.outbox.push(record);
  }
  renderAll();
}

async function checkConflicts() {
  const months = [...new Set(state.outbox.map((r) => r.date.slice(0, 7)))];
  let serverEntries = [];
  for (const m of months) {
    try {
      const data = await api.fetchMonth(m);
      serverEntries = serverEntries.concat(data.entries);
      const loadedAt = new Date().toISOString();
      state.monthCache.set(m, { entries: data.entries, loadedAt, offline: false });
      await queue.setCache(cacheKey("month:" + m), { entries: data.entries, loadedAt });
    } catch (err) {
      // эта помесячная проверка недоступна офлайн — конфликт по ней просто не будет обнаружен сейчас
    }
  }
  const outboxForConflict = state.outbox.map((r) => ({ uuid: r.uuid, date: r.date, tabNum: r.id, value: r.value, knownTs: r.knownTs }));
  return model.findConflicts(outboxForConflict, serverEntries);
}

function showConflictSheet(conflicts) {
  const overlay = document.createElement("div");
  overlay.className = "sheet-overlay";
  const rows = conflicts
    .map((c) => {
      const emp = state.employees.find((e) => e.tabNum === c.tabNum);
      return `
        <div class="sheet__title" style="margin-bottom:4px">${escapeHtml(emp ? emp.name : "Таб. №" + c.tabNum)} · ${c.date}</div>
        <p class="settings-note">В таблице уже записано значение новее того, что вы видели при загрузке.</p>
        <div class="sheet__grid" style="grid-template-columns:1fr 1fr">
          <button class="sheet__btn keep-local" data-uuid="${c.uuid}">Моё: ${escapeHtml(String(c.localValue))}</button>
          <button class="sheet__btn accept-remote" data-uuid="${c.uuid}" data-value="${escapeHtml(String(c.remoteValue))}">Из таблицы: ${escapeHtml(String(c.remoteValue))}</button>
        </div>
      `;
    })
    .join("<hr style='border-color:var(--border);margin:12px 0'>");

  overlay.innerHTML = `<div class="sheet"><div class="sheet__title">Конфликт правок (${conflicts.length})</div>${rows}</div>`;
  el("sheet-root").appendChild(overlay);

  overlay.querySelectorAll(".keep-local").forEach((btn) => {
    btn.addEventListener("click", () => resolveConflict(btn.dataset.uuid, "local"));
  });
  overlay.querySelectorAll(".accept-remote").forEach((btn) => {
    btn.addEventListener("click", () => resolveConflict(btn.dataset.uuid, "remote"));
  });
}

async function resolveConflict(uuid, choice) {
  const rec = state.outbox.find((r) => r.uuid === uuid);
  if (!rec) return;
  if (choice === "remote") {
    await queue.removeFromOutbox(uuid);
  } else {
    rec.knownTs = new Date().toISOString();
    await queue.enqueue(rec);
  }
  state.outbox = (await queue.getOutbox()).filter(isCurrentModeRecord);
  closeSheet();
  renderAll();
  flushQueue();
}

async function flushQueue() {
  if (state.outbox.length === 0) return;
  if (navigator.onLine) {
    const conflicts = await checkConflicts();
    if (conflicts.length) {
      showConflictSheet(conflicts);
      return;
    }
  }
  const result = await queue.flush(api.sendBatch, isCurrentModeRecord);
  state.outbox = (await queue.getOutbox()).filter(isCurrentModeRecord);
  renderAll();
  if (result.networkError) {
    showToast("Нет сети — изменения сохранены локально");
  } else if (result.failed > 0) {
    showToast(`Сохранено ${result.sent}, ошибок: ${result.failed}`);
  } else if (result.sent > 0) {
    showToast("Сохранено в таблицу");
    const monthKey = monthKeyOf(state.currentDate);
    await ensureMonthLoaded(monthKey);
    renderAll();
  }
}

function onNetworkChange() {
  updateNetworkBanner();
  if (navigator.onLine) flushQueue();
}

function updateNetworkBanner() {
  const banner = el("network-banner");
  const monthKey = monthKeyOf(state.currentDate);
  const cached = state.monthCache.get(monthKey);

  if (state.loadError) {
    banner.classList.add("offline");
    banner.textContent = state.loadError;
    return;
  }

  if (!navigator.onLine) {
    banner.classList.add("offline");
    const time = cached && cached.loadedAt ? new Date(cached.loadedAt).toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit" }) : "—";
    banner.textContent = `Нет сети · данные от ${time}`;
  } else {
    banner.classList.remove("offline");
    banner.textContent = state.outbox.length ? `Не отправлено: ${state.outbox.length}` : "";
  }
}

// ---------- переключение экранов ----------

function switchScreen(name) {
  state.screen = name;
  document.querySelectorAll(".screen").forEach((s) => s.classList.remove("active"));
  document.querySelectorAll("nav.bottom-nav button").forEach((b) => b.classList.remove("active"));
  const target = name === "employee" ? "screen-employee" : `screen-${name}`;
  el(target).classList.add("active");
  const navBtn = document.querySelector(`nav.bottom-nav button[data-screen="${name === "employee" ? "month" : name}"]`);
  if (navBtn) navBtn.classList.add("active");
  renderAll();
}

async function changeDay(delta) {
  const d = new Date(state.currentDate);
  d.setDate(d.getDate() + delta);
  const prevMonthKey = monthKeyOf(state.currentDate);
  state.currentDate = d;
  const newMonthKey = monthKeyOf(d);
  if (newMonthKey !== prevMonthKey && !state.monthCache.has(newMonthKey)) {
    await ensureMonthLoaded(newMonthKey);
  }
  renderAll();
}

// ---------- статические обработчики ----------

function wireStaticEvents() {
  el("nav-prev").addEventListener("click", () => changeDay(-1));
  el("nav-next").addEventListener("click", () => changeDay(1));
  el("fill-all-eight").addEventListener("click", fillAllEight);

  document.querySelectorAll("nav.bottom-nav button").forEach((btn) => {
    btn.addEventListener("click", () => switchScreen(btn.dataset.screen));
  });

  el("employee-back").addEventListener("click", () => switchScreen("month"));

  el("month-half-1").addEventListener("click", () => {
    state.monthHalf = 1;
    document.querySelectorAll(".month-toggle button").forEach((b) => b.classList.remove("active"));
    el("month-half-1").classList.add("active");
    renderMonthScreen();
  });
  el("month-half-2").addEventListener("click", () => {
    state.monthHalf = 2;
    document.querySelectorAll(".month-toggle button").forEach((b) => b.classList.remove("active"));
    el("month-half-2").classList.add("active");
    renderMonthScreen();
  });

  document.querySelectorAll("#pay-filter button").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.paySystemFilter = btn.dataset.pay;
      document.querySelectorAll("#pay-filter button").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      renderAll();
    });
  });

  document.querySelectorAll("#theme-switch button").forEach((btn) => {
    btn.addEventListener("click", () => {
      config.setTheme(btn.dataset.theme);
      applyTheme();
      renderSettingsScreen();
    });
  });

  el("settings-save").addEventListener("click", () => {
    config.setScriptUrl(el("settings-url").value.trim());
    config.setPin(el("settings-pin").value.trim());
    const normInput = el("settings-norm").value.trim();
    if (normInput === "") {
      config.setNormHoursOverride(null); // пусто — вернуться к значению из таблицы
    } else {
      const n = Number(normInput);
      if (!Number.isFinite(n) || n <= 0 || n > 24) {
        showToast("Норма часов: число от 0 до 24");
        return;
      }
      config.setNormHoursOverride(n);
    }
    showToast("Настройки сохранены");
    location.reload();
  });
}

// ---------- мелкие утилиты ----------

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

let toastTimer = null;
function showToast(text) {
  const root = el("toast-root");
  root.innerHTML = `<div class="toast">${escapeHtml(text)}</div>`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    root.innerHTML = "";
  }, 2500);
}

main();
