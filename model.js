// model.js — чистые функции над данными табеля. Без DOM, без сети, без побочных эффектов.
// Журнал: запись {date: "YYYY-MM-DD", tabNum: number, value: number|string|"", who, ts, device, uuid}
// Сотрудник: {tabNum, name, position, active: boolean, paySystem: "оклад"|"сдельная"|""}
// Код: {code, label, color, paid: boolean}

const NORM_HOURS = 8;

// ---------- даты ----------

export function pad2(n) {
  return String(n).padStart(2, "0");
}

export function toDateKey(year, month, day) {
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

export function parseDateKey(dateKey) {
  const [y, m, d] = dateKey.split("-").map(Number);
  return { year: y, month: m, day: d };
}

export function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

// День недели: 0 = воскресенье (используем getUTCDay через фиксированную полночь UTC,
// чтобы не зависеть от часового пояса среды выполнения).
export function weekday(dateKey) {
  const { year, month, day } = parseDateKey(dateKey);
  return new Date(Date.UTC(year, month - 1, day)).getUTCDay();
}

export function isSunday(dateKey) {
  return weekday(dateKey) === 0;
}

export function sundaysOfMonth(year, month) {
  const total = daysInMonth(year, month);
  const result = [];
  for (let day = 1; day <= total; day++) {
    if (isSunday(toDateKey(year, month, day))) result.push(day);
  }
  return result;
}

export function monthDateKeys(year, month) {
  const total = daysInMonth(year, month);
  const result = [];
  for (let day = 1; day <= total; day++) {
    result.push(toDateKey(year, month, day));
  }
  return result;
}

// ---------- журнал: последняя запись на пару (дата, таб.№) ----------

function journalKey(dateKey, tabNum) {
  return `${dateKey}__${tabNum}`;
}

// Возвращает Map<"date__tabNum", запись> — последняя по порядку в массиве запись
// для каждой пары. Порядок массива должен соответствовать порядку строк в листе «Журнал»
// (т.е. порядку appendRow), не полю ts.
export function buildLatestMap(journal) {
  const map = new Map();
  for (const entry of journal) {
    map.set(journalKey(entry.date, entry.tabNum), entry);
  }
  return map;
}

// Текущее значение ячейки (дата, таб.№) с учётом истории исправлений/очисток.
// Возвращает undefined, если по этой паре в журнале вообще нет записей.
export function getCellValue(journal, dateKey, tabNum) {
  const map = buildLatestMap(journal);
  const entry = map.get(journalKey(dateKey, tabNum));
  return entry ? entry.value : undefined;
}

// ---------- классификация значения ----------

export function isNumericValue(value) {
  return typeof value === "number" && !Number.isNaN(value);
}

export function isEmptyValue(value) {
  return value === undefined || value === null || value === "";
}

export function isCodeValue(value) {
  return typeof value === "string" && value !== "";
}

// ---------- сотрудники ----------

// paySystem — "оклад" | "сдельная" | "" (пусто = не назначена, не фильтруется).
// Пустая/falsy фильтр-строка означает «все системы оплаты» — возвращает employees как есть.
export function filterByPaySystem(employees, paySystem) {
  if (!paySystem) return employees;
  return employees.filter((e) => e.paySystem === paySystem);
}

// ---------- переработка и цвет ----------

export function overtimeForHours(hours, normHours = NORM_HOURS) {
  if (!isNumericValue(hours)) return 0;
  return hours > normHours ? hours - normHours : 0;
}

// Границы при norm=8: (0;7.5) неполный день, [7.5;8] норма, [8.5;24] переработка —
// т.е. пол-отступа от нормы в обе стороны (шаг значений 0.5, см. DECISIONS.md).
// normHours настраивается в «Настройках» (локально, на устройстве) — это не меняет
// норму и цвета в самой Google-таблице, см. DECISIONS.md.
export function colorForValue(value, codes = [], normHours = NORM_HOURS) {
  if (isEmptyValue(value)) return null;
  if (isNumericValue(value)) {
    if (value >= normHours + 0.5) return "#4D94D8";
    if (value >= normHours - 0.5) return "#AFD2D3";
    if (value > 0) return "#BF4F14";
    return null;
  }
  if (isCodeValue(value)) {
    const code = codes.find((c) => c.code === value);
    return code ? code.color : null;
  }
  return null;
}

// ---------- итоги за день ----------

// employees — активные сотрудники за день; journal — весь журнал (функция сама найдёт
// актуальные значения на нужную дату).
export function dayTotals(journal, employees, dateKey) {
  const map = buildLatestMap(journal);
  let onShift = 0;
  let hoursSum = 0;
  let unmarked = 0;
  const byCode = {};

  for (const emp of employees) {
    const entry = map.get(journalKey(dateKey, emp.tabNum));
    const value = entry ? entry.value : undefined;
    if (isEmptyValue(value)) {
      unmarked++;
    } else if (isNumericValue(value)) {
      onShift++;
      hoursSum += value;
    } else if (isCodeValue(value)) {
      byCode[value] = (byCode[value] || 0) + 1;
    }
  }

  return { onShift, hoursSum, unmarked, byCode };
}

// ---------- итоги за месяц по сотруднику ----------

export function monthSummaryForEmployee(journal, tabNum, year, month, normHours = NORM_HOURS) {
  const map = buildLatestMap(journal);
  const dateKeys = monthDateKeys(year, month);
  let days = 0;
  let hours = 0;
  let overtime = 0;
  const byCode = {};

  for (const dateKey of dateKeys) {
    const entry = map.get(journalKey(dateKey, tabNum));
    const value = entry ? entry.value : undefined;
    if (isNumericValue(value)) {
      days++;
      hours += value;
      overtime += overtimeForHours(value, normHours);
    } else if (isCodeValue(value)) {
      byCode[value] = (byCode[value] || 0) + 1;
    }
  }

  return { days, hours, overtime, byCode };
}

// Итоги месяца по всем активным сотрудникам (для «ИТОГО» на экране «Месяц»).
export function monthSummaryTotals(journal, employees, year, month, normHours = NORM_HOURS) {
  let days = 0;
  let hours = 0;
  let overtime = 0;
  for (const emp of employees) {
    const s = monthSummaryForEmployee(journal, emp.tabNum, year, month, normHours);
    days += s.days;
    hours += s.hours;
    overtime += s.overtime;
  }
  return { days, hours, overtime };
}

// ---------- календарь месяца (для экрана «Месяц» / «Сотрудник») ----------

// Возвращает [{tabNum, name, position, values: {dateKey: value}}] для активных сотрудников.
export function monthGrid(journal, employees, year, month) {
  const map = buildLatestMap(journal);
  const dateKeys = monthDateKeys(year, month);
  return employees
    .filter((e) => e.active)
    .map((emp) => {
      const values = {};
      for (const dateKey of dateKeys) {
        const entry = map.get(journalKey(dateKey, emp.tabNum));
        values[dateKey] = entry ? entry.value : undefined;
      }
      return { tabNum: emp.tabNum, name: emp.name, position: emp.position, values };
    });
}

// ---------- конфликты офлайн-очереди ----------

// outbox: [{uuid, date, tabNum, value, knownTs}] — knownTs это ts записи, которую клиент
// видел в момент начала правки (или undefined, если ячейка была пуста при загрузке).
// serverEntries: [{date, tabNum, value, ts}] — актуальные записи месяца с сервера.
// Возвращает список конфликтов: когда на сервере появилась запись новее той, что видел
// клиент, и её значение отличается от того, что клиент пытается отправить.
export function findConflicts(outbox, serverEntries) {
  const serverMap = new Map();
  for (const entry of serverEntries) {
    serverMap.set(journalKey(entry.date, entry.tabNum), entry);
  }

  const conflicts = [];
  for (const item of outbox) {
    const server = serverMap.get(journalKey(item.date, item.tabNum));
    if (!server) continue;
    const knownTs = item.knownTs ? new Date(item.knownTs).getTime() : 0;
    const serverTs = server.ts ? new Date(server.ts).getTime() : 0;
    if (serverTs > knownTs && server.value !== item.value) {
      conflicts.push({
        uuid: item.uuid,
        date: item.date,
        tabNum: item.tabNum,
        localValue: item.value,
        remoteValue: server.value,
        remoteTs: server.ts,
      });
    }
  }
  return conflicts;
}

export const constants = { NORM_HOURS };
