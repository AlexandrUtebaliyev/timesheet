// demo-data.js — данные для режима ?demo=1. Работает без сети, на текущем месяце.
import { toDateKey, daysInMonth, isSunday } from "./model.js";

export const DEMO_CODES = [
  { code: "В", label: "Выходной", color: "#B3E5A1", paid: null },
  { code: "О", label: "Отпуск", color: "#FAE2D5", paid: true },
  { code: "Б", label: "Больничный", color: "#FFFF66", paid: true },
  { code: "БС", label: "Без сохранения содержания", color: "#FFCCFF", paid: false },
  { code: "К", label: "Командировка", color: "#33CCCC", paid: true },
  { code: "П", label: "Прогул", color: "#D9D9D9", paid: false },
  { code: "ог", label: "Отгул", color: "#FABF8F", paid: false },
  { code: "до", label: "Дополнительный отпуск", color: "#99FF99", paid: true },
  { code: "у", label: "Учёба", color: "#FFFFCC", paid: false },
];

export const DEMO_EMPLOYEES = [
  { tabNum: 1, name: "Иванов Иван Иванович", position: "токарь", active: true },
  { tabNum: 2, name: "Петров Пётр Петрович", position: "фрезеровщик", active: true },
  { tabNum: 3, name: "Сидоров Сидор Сидорович", position: "слесарь", active: true },
  { tabNum: 4, name: "Кузнецов Кузьма Кузьмич", position: "сварщик", active: true },
  { tabNum: 5, name: "Соколов Семён Матвеевич", position: "токарь", active: true },
  { tabNum: 6, name: "Попов Павел Романович", position: "слесарь", active: true },
  { tabNum: 7, name: "Лебедев Лев Егорович", position: "наладчик", active: true },
  { tabNum: 8, name: "Новиков Никита Фёдорович", position: "сварщик", active: true },
  { tabNum: 9, name: "Морозов Марк Данилович", position: "фрезеровщик", active: true },
  { tabNum: 10, name: "Волков Владимир Ильич", position: "слесарь", active: true },
  { tabNum: 11, name: "Зайцев Захар Андреевич", position: "токарь", active: true },
  { tabNum: 12, name: "Уволенный Устин Юрьевич", position: "слесарь", active: false },
];

export const DEMO_NORM_HOURS = 8;

export function buildDemoJournal(today = new Date()) {
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const total = daysInMonth(year, month);
  const todayDay = today.getDate();
  const journal = [];
  let seq = 0;

  function push(date, tabNum, value) {
    seq++;
    journal.push({
      date,
      tabNum,
      value,
      who: "мастер",
      ts: new Date(today.getTime() - (total - 1) * 86400000 + seq * 60000).toISOString(),
      device: "demo",
      uuid: "demo-" + seq,
    });
  }

  for (let day = 1; day < todayDay; day++) {
    const date = toDateKey(year, month, day);
    if (isSunday(date)) continue;
    for (const emp of DEMO_EMPLOYEES) {
      if (!emp.active) continue;
      if (emp.tabNum === 3 && day === 2) {
        push(date, emp.tabNum, "Б");
        continue;
      }
      if (emp.tabNum === 5 && day === 3) {
        push(date, emp.tabNum, "О");
        continue;
      }
      if (emp.tabNum === 7 && (day === 1 || day === 2)) {
        push(date, emp.tabNum, "К");
        continue;
      }
      if (emp.tabNum === 2 && day === 4) {
        push(date, emp.tabNum, 10); // переработка
        continue;
      }
      push(date, emp.tabNum, 8);
    }
  }

  // демонстрация «последняя запись побеждает»: в первый рабочий день месяца
  // сотрудника 1 сначала отметили на 8, затем исправили на 9.
  const firstWorkDay = isSunday(toDateKey(year, month, 1)) ? 2 : 1;
  if (firstWorkDay < todayDay) {
    push(toDateKey(year, month, firstWorkDay), 1, 9);
  }

  return journal;
}

export function buildDemoSnapshot(today = new Date()) {
  return {
    employees: DEMO_EMPLOYEES,
    codes: DEMO_CODES,
    normHours: DEMO_NORM_HOURS,
    journal: buildDemoJournal(today),
    loadedAt: new Date().toISOString(),
  };
}
