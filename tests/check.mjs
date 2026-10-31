// node tests/check.mjs
// Простой тестовый раннер без зависимостей, проверяет model.js на фикстуре.

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import * as model from "../model.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixture = JSON.parse(readFileSync(join(__dirname, "fixture.json"), "utf8"));
const { employees, codes, journal } = fixture;
const activeEmployees = employees.filter((e) => e.active);

let passed = 0;
let failed = 0;
const failures = [];

function eq(actual, expected) {
  return JSON.stringify(actual) === JSON.stringify(expected);
}

function check(name, actual, expected) {
  if (eq(actual, expected)) {
    passed++;
  } else {
    failed++;
    failures.push(`${name}\n  ожидалось: ${JSON.stringify(expected)}\n  получено:  ${JSON.stringify(actual)}`);
  }
}

// --- даты и воскресенья ---

check(
  "воскресенья ноября 2026",
  model.sundaysOfMonth(2026, 11),
  [1, 8, 15, 22, 29]
);

check("дней в ноябре 2026 = 30 (31-го нет)", model.daysInMonth(2026, 11), 30);

const novDates = model.monthDateKeys(2026, 11);
check("monthDateKeys длина 30", novDates.length, 30);
check("monthDateKeys последняя дата 2026-11-30", novDates[novDates.length - 1], "2026-11-30");
check("monthDateKeys не содержит 2026-11-31", novDates.includes("2026-11-31"), false);

check("isSunday(2026-11-01) = true", model.isSunday("2026-11-01"), true);
check("isSunday(2026-11-02) = false", model.isSunday("2026-11-02"), false);

// --- последняя запись побеждает, пустое значение очищает ---

check("последняя запись побеждает (таб.1, 02.11 -> 9)", model.getCellValue(journal, "2026-11-02", 1), 9);
check("пустое значение очищает (таб.2, 03.11 -> \"\")", model.getCellValue(journal, "2026-11-03", 2), "");
check("нет записи -> undefined (таб.2, 02.11)", model.getCellValue(journal, "2026-11-02", 2), undefined);
check("код как значение (таб.3, 02.11 -> Б)", model.getCellValue(journal, "2026-11-02", 3), "Б");

// --- итоги дня ---

check("итоги дня 02.11", model.dayTotals(journal, activeEmployees, "2026-11-02"), {
  onShift: 2,
  hoursSum: 16.5,
  unmarked: 1,
  byCode: { Б: 1 },
});

check("итоги дня 03.11 (с очисткой)", model.dayTotals(journal, activeEmployees, "2026-11-03"), {
  onShift: 1,
  hoursSum: 8.5,
  unmarked: 3,
  byCode: {},
});

// --- итоги месяца по сотруднику: дни = число числовых отметок, часы = сумма, переработка = сумма (ч-8) > 8 ---

check("итоги месяца таб.1 (дни/часы/переработка)", model.monthSummaryForEmployee(journal, 1, 2026, 11), {
  days: 3,
  hours: 27,
  overtime: 3,
  byCode: {},
});

check("итоги месяца таб.2 (всё очищено)", model.monthSummaryForEmployee(journal, 2, 2026, 11), {
  days: 0,
  hours: 0,
  overtime: 0,
  byCode: {},
});

check("итоги месяца таб.3 (только код Б)", model.monthSummaryForEmployee(journal, 3, 2026, 11), {
  days: 0,
  hours: 0,
  overtime: 0,
  byCode: { Б: 1 },
});

check("итоги месяца таб.4 (граничные часы 7.5 и 8.5)", model.monthSummaryForEmployee(journal, 4, 2026, 11), {
  days: 2,
  hours: 16,
  overtime: 0.5,
  byCode: {},
});

check("итоги месяца по всем активным", model.monthSummaryTotals(journal, activeEmployees, 2026, 11), {
  days: 5,
  hours: 43,
  overtime: 3.5,
});

// --- цвета по границам 7.5 / 8 / 8.5 ---

check("цвет 7.4 (неполный)", model.colorForValue(7.4, codes), "#BF4F14");
check("цвет 0.5 (неполный)", model.colorForValue(0.5, codes), "#BF4F14");
check("цвет 7.5 (норма)", model.colorForValue(7.5, codes), "#AFD2D3");
check("цвет 8 (норма)", model.colorForValue(8, codes), "#AFD2D3");
check("цвет 8.5 (переработка)", model.colorForValue(8.5, codes), "#4D94D8");
check("цвет 24 (переработка)", model.colorForValue(24, codes), "#4D94D8");
check("цвет кода В", model.colorForValue("В", codes), "#B3E5A1");
check("цвет кода ог (строчный)", model.colorForValue("ог", codes), "#FABF8F");
check("цвет пустого значения", model.colorForValue("", codes), null);
check("цвет отсутствующего значения", model.colorForValue(undefined, codes), null);

// --- сетка месяца: только активные, значения по датам из журнала ---

const grid = model.monthGrid(journal, employees, 2026, 11);
check("monthGrid: только активные (4 из 5)", grid.length, 4);
check("monthGrid: уволенный не попал", grid.some((e) => e.tabNum === 5), false);
const grid1 = grid.find((e) => e.tabNum === 1);
check("monthGrid: таб.1, 02.11 = 9 (с учётом исправления)", grid1.values["2026-11-02"], 9);
const grid2 = grid.find((e) => e.tabNum === 2);
check("monthGrid: таб.2, 03.11 = \"\" (очищено)", grid2.values["2026-11-03"], "");

// --- фильтр по системе оплаты труда ---

check(
  "filterByPaySystem('') -> все без изменений",
  model.filterByPaySystem(activeEmployees, "").map((e) => e.tabNum),
  [1, 2, 3, 4]
);
check(
  "filterByPaySystem('оклад') -> только таб.1",
  model.filterByPaySystem(activeEmployees, "оклад").map((e) => e.tabNum),
  [1]
);
check(
  "filterByPaySystem('сдельная') -> таб.2 и 3",
  model.filterByPaySystem(activeEmployees, "сдельная").map((e) => e.tabNum),
  [2, 3]
);
check(
  "filterByPaySystem: пустая paySystem у сотрудника не попадает ни в один конкретный фильтр",
  model.filterByPaySystem(activeEmployees, "оклад").some((e) => e.tabNum === 4),
  false
);

// --- настраиваемая норма часов (граница 7.5/8/8.5 сдвигается вместе с нормой) ---

check("цвет 8 при норме 10 (неполный)", model.colorForValue(8, codes, 10), "#BF4F14");
check("цвет 9.5 при норме 10 (норма)", model.colorForValue(9.5, codes, 10), "#AFD2D3");
check("цвет 10.5 при норме 10 (переработка)", model.colorForValue(10.5, codes, 10), "#4D94D8");
check("переработка при норме 10 (10ч -> 0)", model.overtimeForHours(10, 10), 0);
check("переработка при норме 10 (12ч -> 2)", model.overtimeForHours(12, 10), 2);
check("итоги месяца таб.4 при норме 10 (переработки нет)", model.monthSummaryForEmployee(journal, 4, 2026, 11, 10).overtime, 0);

// --- конфликты офлайн-очереди ---

const serverEntries = [
  { date: "2026-11-02", tabNum: 1, value: "Б", ts: "2026-11-02T20:00:00Z" },
  { date: "2026-11-03", tabNum: 2, value: 4, ts: "2026-11-03T08:00:00Z" },
];

check(
  "конфликт: сервер новее и значение другое",
  model.findConflicts(
    [{ uuid: "a", date: "2026-11-02", tabNum: 1, value: 9, knownTs: "2026-11-02T08:00:00Z" }],
    serverEntries
  ),
  [
    {
      uuid: "a",
      date: "2026-11-02",
      tabNum: 1,
      localValue: 9,
      remoteValue: "Б",
      remoteTs: "2026-11-02T20:00:00Z",
    },
  ]
);

check(
  "нет конфликта: знали запись новее серверной",
  model.findConflicts(
    [{ uuid: "b", date: "2026-11-02", tabNum: 1, value: 9, knownTs: "2026-11-03T00:00:00Z" }],
    serverEntries
  ),
  []
);

check(
  "нет конфликта: значение совпадает с серверным",
  model.findConflicts(
    [{ uuid: "c", date: "2026-11-03", tabNum: 2, value: 4, knownTs: "2026-11-01T00:00:00Z" }],
    serverEntries
  ),
  []
);

// --- итог ---

console.log(`Пройдено: ${passed}, провалено: ${failed}`);
if (failed > 0) {
  console.error("\nПровалившиеся проверки:\n");
  for (const f of failures) console.error(f + "\n");
  process.exit(1);
}
