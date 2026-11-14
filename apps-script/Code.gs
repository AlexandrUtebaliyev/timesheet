/**
 * Табель — backend на Google Apps Script, привязанный к таблице.
 * Публикуется как веб-приложение (Выполнять от имени: я; Доступ: все).
 * Сам скрипт таблицу не публикует и ничего в неё не пишет при деплое — см. README.md.
 */

var SHEET_EMPLOYEES = "Сотрудники";
var SHEET_CODES = "Коды";
var SHEET_JOURNAL = "Журнал";
var SHEET_TABEL = "Табель";
var SHEET_HARMFUL = "Вредники";
var SHEET_ORDERS = "Наряды";
var SHEET_OLD = "Октябрь(Объем)";
var SHEET_ARCHIVE_SOURCE = "Архив_исходный";

var JOURNAL_HEADERS = ["Дата", "Таб. №", "Значение", "Кто", "Когда записано", "ID устройства", "UUID"];

var CODE_TABLE = [
  ["В", "Выходной", "#B3E5A1", ""],
  ["О", "Отпуск", "#FAE2D5", "да"],
  ["Б", "Больничный", "#FFFF66", "да"],
  ["БС", "Без сохранения содержания", "#FFCCFF", "нет"],
  ["К", "Командировка", "#33CCCC", "да"],
  ["П", "Прогул", "#D9D9D9", "нет"],
  ["ог", "Отгул", "#FABF8F", "нет"],
  ["до", "Дополнительный отпуск", "#99FF99", "да"],
  ["у", "Учёба", "#FFFFCC", "нет"],
];

var NORM_HOURS = 8;
var MILK_LITERS_PER_DAY = 0.3; // литров молока за каждый отработанный день (как в исходных «Вредники»)
var ORDERS_MIN_ROWS = 50; // строк в «Нарядах» по умолчанию; если в данных больше — таблица растёт
var ORDERS_CAP = 400; // максимум строк с формулами в «Нарядах»
var ORDERS_SOURCE_ID = "1qlp65ur6bxked5niY2O0Zn1Awh7vHmdBSmIv20G93TE"; // файл с данными нарядов (листы January..December)
var SCHEDULE_5 = "5-дневная";
var SCHEDULE_6 = "6-дневная";
var SHIFT_HOURS = 8; // длина рабочего дня по графику
var SHEET_PARAMS = "Параметры";
var PARAMS_HOLIDAYS = "E2:E61"; // праздничные дни (даты) — исключаются из нормы дней месяца
// Параметры в листе «Параметры»: строка = индекс + 2, значение в столбце B. Порядок не менять
// без пересборки: формулы ссылаются на конкретные ячейки.
var PARAM_DEFS = [
  { key: "org", label: "Организация", value: 'ТОО "ПК Алтын Таға"', note: "Шапка «Табеля», «Нарядов», «Вредников»" },
  { key: "unit", label: "Структурное подразделение", value: "Цех металлообработки", note: "Шапка «Табеля», «Нарядов», «Вредников»" },
  { key: "title", label: "Заголовок табеля", value: "ТАБЕЛЬ УЧЕТА РАБОЧЕГО ВРЕМЕНИ И РАСЧЕТА ЗАРАБОТНОЙ ПЛАТЫ", note: "Печатная форма «Табеля»" },
  { key: "headPos", label: "Должность руководителя («Утверждаю»)", value: "Директор", note: "Далее добавляется название организации" },
  { key: "headName", label: "ФИО руководителя", value: "Нұрмағамбет Ә. Е.", note: "Подпись «Утверждаю»" },
  { key: "chief1Pos", label: "Подписант 1 — должность", value: "Нач. цеха", note: "Под таблицей «Табеля» и «Нарядов»" },
  { key: "chief1Name", label: "Подписант 1 — ФИО", value: "Утебалиев А.В.", note: "" },
  { key: "chief2Pos", label: "Подписант 2 — должность", value: "Гл. инженер", note: "" },
  { key: "chief2Name", label: "Подписант 2 — ФИО", value: "Минбаев К.К.", note: "" },
  { key: "shift", label: "Длина смены, ч", value: SHIFT_HOURS, note: "Для расчёта оклада: оклад ÷ (рабочие дни × смена) × часы" },
  { key: "overtimeFrom", label: "Переработка начинается после, ч/день", value: NORM_HOURS, note: "Столбец «переработка» в «Табеле»" },
  { key: "overtimeK", label: "Коэффициент оплаты переработки", value: 1, note: "1 = как обычные часы, 1,5 = полуторная. Только для окладников" },
  { key: "milk", label: "Молоко, литров за отработанный день", value: MILK_LITERS_PER_DAY, note: "Лист «Вредники»" },
];
var TOTAL_FILL = "#FFFF99"; // заливка строк «ИТОГО» на всех листах
var DEFAULT_FONT = "Arial"; // шрифт по умолчанию во всех листах

// ===================== меню =====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Табель")
    .addItem("Сохранить месяц в архив", "archiveCurrentMonth")
    .addItem("Пересоздать форму", "rebuildTabelSheet")
    .addItem("Обновить «Вредники»", "rebuildHarmfulSheet")
    .addItem("Пересоздать наряды", "rebuildOrdersSheet")
    .addItem("Обновить рамку нарядов", "refreshOrdersLayout")
    .addToUi();
  try {
    refreshOrdersLayout(); // подогнать рамку «Нарядов» под текущее число строк данных
  } catch (err) {
    Logger.log("refreshOrdersLayout: " + err);
  }
}

// ===================== setupSheets: разовая миграция =====================

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  renameOldSheetToArchive(ss, log);
  ensureParamsSheet(ss, log);
  var employeesCreated = ensureEmployeesSheet(ss, log);
  ensureCodesSheet(ss, log);
  ensureJournalSheet(ss, log);
  ensureTabelSheet(ss, log, employeesCreated);
  ensureHarmfulSheet(ss, log);

  Logger.log(log.join("\n"));
  return log.join("\n");
}

function renameOldSheetToArchive(ss, log) {
  var old = ss.getSheetByName(SHEET_OLD);
  var archive = ss.getSheetByName(SHEET_ARCHIVE_SOURCE);
  if (archive) {
    log.push('Лист "' + SHEET_ARCHIVE_SOURCE + '" уже существует — пропуск переименования.');
    return;
  }
  if (old) {
    old.setName(SHEET_ARCHIVE_SOURCE);
    log.push('Лист "' + SHEET_OLD + '" переименован в "' + SHEET_ARCHIVE_SOURCE + '".');
  } else {
    log.push('Лист "' + SHEET_OLD + '" не найден — пропуск переименования (возможно, уже переименован ранее).');
  }
}

function ensureEmployeesSheet(ss, log) {
  var existing = ss.getSheetByName(SHEET_EMPLOYEES);
  if (existing) {
    log.push('Лист "' + SHEET_EMPLOYEES + '" уже существует — пропуск.');
    return false;
  }
  var sheet = ss.insertSheet(SHEET_EMPLOYEES);
  sheet.getRange(1, 1, 1, 8).setValues([["Таб. №", "ФИО", "Должность", "Активен", "Система оплаты", "Оклад", "Вредность", "График"]]);
  sheet.setFrozenRows(1);

  var source = ss.getSheetByName(SHEET_ARCHIVE_SOURCE);
  var rows = [];
  if (source) {
    var data = source.getRange(18, 1, 12, 3).getValues(); // A:C, строки 18-29
    for (var i = 0; i < data.length; i++) {
      var name = data[i][1];
      var position = data[i][2];
      if (!name) continue;
      rows.push([i + 1, name, position, "да", "", "", "нет", SCHEDULE_5]); // систему оплаты назначает пользователь вручную
    }
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  }
  // Выпадающий список в "Система оплаты" — чтобы не разъезжались варианты написания
  // ("Оклад" vs "оклад"), от которых зависит фильтр в приложении.
  var paySystemRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["оклад", "сдельная"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 5, 999, 1).setDataValidation(paySystemRule);
  // «Вредность»: "да" — сотрудник попадает в ведомость молока на листе «Вредники».
  var harmfulRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["да", "нет"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 7, 999, 1).setDataValidation(harmfulRule);
  // «График»: 8-часовой день, 5- или 6-дневная неделя — по нему считается норма месяца
  // и зарплата окладников в «Табеле».
  var scheduleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([SCHEDULE_5, SCHEDULE_6], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 8, 999, 1).setDataValidation(scheduleRule);
  sheet.autoResizeColumns(1, 8);
  log.push('Лист "' + SHEET_EMPLOYEES + '" создан, заполнено сотрудников: ' + rows.length + ".");
  return true;
}

function ensureCodesSheet(ss, log) {
  var existing = ss.getSheetByName(SHEET_CODES);
  if (existing) {
    log.push('Лист "' + SHEET_CODES + '" уже существует — пропуск.');
    return;
  }
  var sheet = ss.insertSheet(SHEET_CODES);
  sheet.getRange(1, 1, 1, 4).setValues([["Код", "Расшифровка", "Цвет", "Оплачивается"]]);
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, CODE_TABLE.length, 4).setValues(CODE_TABLE);
  sheet.autoResizeColumns(1, 4);
  log.push('Лист "' + SHEET_CODES + '" создан, кодов: ' + CODE_TABLE.length + ".");
}

function ensureJournalSheet(ss, log) {
  var existing = ss.getSheetByName(SHEET_JOURNAL);
  if (existing) {
    log.push('Лист "' + SHEET_JOURNAL + '" уже существует — пропуск.');
    return;
  }
  var sheet = ss.insertSheet(SHEET_JOURNAL);
  sheet.getRange(1, 1, 1, JOURNAL_HEADERS.length).setValues([JOURNAL_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd");
  sheet.autoResizeColumns(1, JOURNAL_HEADERS.length);

  var protection = sheet.protect().setDescription("Журнал: строки не удалять и не править вручную");
  protection.setWarningOnly(true);

  log.push('Лист "' + SHEET_JOURNAL + '" создан и защищён (предупреждение при правке).');
}

function ensureTabelSheet(ss, log, employeesJustCreated) {
  // «Табель» и «Наряды» ссылаются друг на друга (месяц / итог нарядов), поэтому оба листа
  // сначала создаются пустыми, и только потом собираются.
  var tabelExisting = ss.getSheetByName(SHEET_TABEL);
  var ordersExisting = ss.getSheetByName(SHEET_ORDERS);
  var tabel = tabelExisting || ss.insertSheet(SHEET_TABEL);
  var orders = ordersExisting || ss.insertSheet(SHEET_ORDERS);

  if (tabelExisting) {
    log.push('Лист "' + SHEET_TABEL + '" уже существует — пропуск (используйте "Пересоздать форму" для обновления формул).');
  } else {
    buildTabelSheet(tabel);
    log.push('Лист "' + SHEET_TABEL + '" создан.');
  }
  if (ordersExisting) {
    log.push('Лист "' + SHEET_ORDERS + '" уже существует — пропуск (используйте "Пересоздать форму").');
  } else {
    buildOrdersSheet(orders);
    log.push('Лист "' + SHEET_ORDERS + '" создан.');
  }
  moveOrdersNextToTabel(ss);
}

// «Наряды» стоят сразу справа от «Табеля».
function moveOrdersNextToTabel(ss) {
  var tabel = ss.getSheetByName(SHEET_TABEL);
  var orders = ss.getSheetByName(SHEET_ORDERS);
  if (!tabel || !orders) return;
  ss.setActiveSheet(orders);
  ss.moveActiveSheet(tabel.getIndex() + 1);
  ss.setActiveSheet(tabel);
}

// Шрифт по умолчанию на всём листе (форматирование из архивных листов тянет свой шрифт).
function applyDefaultFont(sheet) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily(DEFAULT_FONT);
}

// ===================== лист «Параметры» =====================

// Ссылка на ячейку значения параметра (для формул).
function paramCell(key) {
  for (var i = 0; i < PARAM_DEFS.length; i++) {
    if (PARAM_DEFS[i].key === key) return SHEET_PARAMS + "!$B$" + (i + 2);
  }
  throw new Error("Неизвестный параметр: " + key);
}

// Числовой параметр для формулы; если ячейка пуста или не число — значение по умолчанию.
function paramNum(key) {
  var cell = paramCell(key);
  var def = PARAM_DEFS.filter(function (d) { return d.key === key; })[0].value;
  return "IF(ISNUMBER(" + cell + ")," + cell + "," + def + ")";
}

// Формула «текст параметра» с необязательным префиксом/суффиксом (строковые литералы).
function paramText(key, prefix, suffix) {
  var q = function (s) { return '"' + String(s).replace(/"/g, '""') + '"'; };
  return (prefix ? q(prefix) + "&" : "") + paramCell(key) + (suffix ? "&" + q(suffix) : "");
}

// Лист создаётся один раз; существующие значения не перезаписываются.
function ensureParamsSheet(ss, log) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var existing = ss.getSheetByName(SHEET_PARAMS);
  if (existing) {
    if (log) log.push('Лист "' + SHEET_PARAMS + '" уже существует — пропуск.');
    return existing;
  }
  var sheet = ss.insertSheet(SHEET_PARAMS);
  sheet.getRange(1, 1, 1, 3).setValues([["Параметр", "Значение", "Примечание"]]).setFontWeight("bold");
  sheet.getRange(1, 5).setValue("Праздничные дни (даты)").setFontWeight("bold");
  var rows = PARAM_DEFS.map(function (d) { return [d.label, d.value, d.note]; });
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(2, 2, rows.length, 1).setBackground("#FFFFCC").setHorizontalAlignment("left");
  sheet.getRange(PARAMS_HOLIDAYS).setNumberFormat("dd.mm.yyyy").setBackground("#FFFFCC");
  sheet.getRange("G1").setValue("Жёлтые ячейки можно менять — формулы подхватят значения сразу. Даты праздников вписывайте в столбец E: они вычитаются из нормы дней месяца (оклад).")
    .setFontColor("#888888");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 420);
  sheet.setColumnWidth(3, 380);
  sheet.setColumnWidth(4, 20);
  sheet.setColumnWidth(5, 150);
  applyDefaultFont(sheet);
  if (log) log.push('Лист "' + SHEET_PARAMS + '" создан.');
  return sheet;
}

// Ширина столбца под самое длинное значение (в одну строку, без переноса): ~8 px на символ + поля.
// Существующая ширина (например, из архивной формы) не уменьшается.
function fitColumnWidth(sheet, col, texts) {
  var max = 0;
  texts.forEach(function (t) { max = Math.max(max, String(t == null ? "" : t).trim().length); });
  var width = Math.max(sheet.getColumnWidth(col), Math.round(max * 8) + 20);
  sheet.setColumnWidth(col, width);
}

// Строка «ИТОГО»: жирным и светлая яркая заливка.
function styleTotalRow(range) {
  range.setFontWeight("bold").setBackground(TOTAL_FILL);
}

function resetSheet(sheet) {
  sheet.clear();
  sheet.clearConditionalFormatRules();
}

function rebuildTabelSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureParamsSheet(ss);
  var tabel = ss.getSheetByName(SHEET_TABEL) || ss.insertSheet(SHEET_TABEL);
  var orders = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS);
  resetSheet(tabel);
  resetSheet(orders);
  buildOrdersSheet(orders);
  buildTabelSheet(tabel);
  moveOrdersNextToTabel(ss);
  ss.toast("Листы «Табель» и «Наряды» пересобраны.");
}

function rebuildOrdersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureParamsSheet(ss);
  var orders = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS);
  resetSheet(orders);
  buildOrdersSheet(orders);
  moveOrdersNextToTabel(ss);
  ss.toast("Лист «Наряды» пересобран.");
}

/**
 * setFormula принимает формулы в английской нотации (запятые), но таблица с региональными
 * настройками, где разделитель аргументов другой (например «Казахстан»: ";"), отвергает их
 * с «Formula parse error» во всех ячейках. Пробуем простую формулу в служебной ячейке и,
 * если она не считается, останавливаемся с понятным сообщением вместо «сломанной» формы.
 */
function assertFormulasSupported(sheet) {
  var probe = sheet.getRange("A1");
  probe.setFormula("=IF(TRUE,1,2)");
  SpreadsheetApp.flush();
  var ok = probe.getValue() === 1;
  probe.clear();
  if (!ok) {
    var locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
    throw new Error(
      "Формулы не разбираются при региональных настройках таблицы (" + locale + "). " +
      "Откройте Файл → Настройки → Региональные настройки, выберите «Россия» (или «США»), " +
      "сохраните и запустите снова."
    );
  }
}

/**
 * Собирает лист «Табель»: шапку, формулы сетки дней и итогов, условное форматирование.
 * Формулы пишутся в английской нотации — это требование Apps Script (SpreadsheetApp
 * принимает setFormula только с английскими именами функций); пользователю в интерфейсе
 * таблицы они всё равно показываются по-русски, если у таблицы русская локаль.
 */
function buildTabelSheet(sheet) {
  assertFormulasSupported(sheet);
  sheet.getRange("A3").setFormula("=" + paramText("org"));
  sheet.getRange("A4").setValue("(наименование организации)");
  sheet.getRange("A5").setFormula("=" + paramText("unit"));
  sheet.getRange("A6").setValue("(структурное подразделение)");
  sheet.getRange("P10").setFormula("=" + paramText("title"));

  // AN11 — первое число месяца; источник правды для всей сетки. Прячем его, т.к. это
  // технический параметр, а не часть печатной формы.
  // ВАЖНО: не класть эту ячейку в столбец AH — AH одновременно последний столбец сетки
  // дней (день 31, D..AH = 31 столбец), и hideColumns на нём прятал бы весь день 31 у
  // всех сотрудников. AN (40) — первый свободный столбец правее итогов (AI..AM).
  // Строим дату на полдень UTC (а не "new Date(); setDate(1)", которая берёт текущее
  // время суток) — иначе при несовпадении часового пояса проекта Apps Script и часового
  // пояса таблицы дата может "съехать" на соседний день и формулы D18:AH29 перестанут
  // находить записи Журнала за этот день (см. DECISIONS.md).
  var now = new Date();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var ym = Utilities.formatDate(now, tz, "yyyy-MM");
  var first = new Date(ym + "-01T12:00:00Z");
  if (sheet.getMaxColumns() < 43) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 43 - sheet.getMaxColumns());
  }
  sheet.getRange("AN11").setValue(first).setNumberFormat("dd.mm.yyyy");
  sheet.hideColumns(40, 1); // AN = столбец 40

  // Норма рабочих дней месяца по графику (без учёта праздников): AP — 5-дневка (сб, вс
  // выходные), AQ — 6-дневка (выходной только вс). Нужны для зарплаты окладников в AM.
  sheet.getRange("AP11").setFormula('=NETWORKDAYS.INTL($AN$11,EOMONTH($AN$11,0),"0000011",' + SHEET_PARAMS + '!$' + PARAMS_HOLIDAYS.replace(":", ":$") + ')');
  sheet.getRange("AQ11").setFormula('=NETWORKDAYS.INTL($AN$11,EOMONTH($AN$11,0),"0000001",' + SHEET_PARAMS + '!$' + PARAMS_HOLIDAYS.replace(":", ":$") + ')');

  sheet.getRange("O11").setFormula(
    '="За "&SWITCH(MONTH($AN$11),1,"январь",2,"февраль",3,"март",4,"апрель",5,"май",6,"июнь",7,"июль",8,"август",9,"сентябрь",10,"октябрь",11,"ноябрь",12,"декабрь")&" "&YEAR($AN$11)&" года"'
  );

  sheet.getRange("AI3").setValue("Утверждаю");
  sheet.getRange("AI4").setFormula("=" + paramCell("headPos") + '&" "&' + paramCell("org"));
  sheet.getRange("AI5").setFormula("=" + paramText("headName", "______________"));

  // Шапка итогов — три вложенные строки, как в "Архив_исходный":
  // 13: общий заголовок; 14: "дней"/"часов"; 15: "часов"/"из них"; 16: разбивка.
  sheet.getRange("AI13").setValue("Итого отработано за месяц");
  sheet.getRange("AI14").setValue("дней");
  sheet.getRange("AJ14").setValue("часов");
  sheet.getRange("AJ15").setValue("часов");
  sheet.getRange("AK15").setValue("из них");
  sheet.getRange("AK16").setValue("ремонтных");
  sheet.getRange("AL16").setValue("переработка");
  sheet.getRange("AM16").setValue("зарплата");

  // Строка 14 — числа дней месяца; строка 13 — пометка выходных.
  for (var col = 4; col <= 34; col++) { // D..AH
    var colLetter = columnToLetter(col);
    sheet
      .getRange(colLetter + "14")
      .setFormula("=IF(MONTH($AN$11+COLUMN()-4)=MONTH($AN$11),$AN$11+COLUMN()-4,\"\")")
      .setNumberFormat("d");
    sheet.getRange(colLetter + "13").setFormula('=IF(' + colLetter + '14="","",IF(WEEKDAY(' + colLetter + '14)=1,"В",""))');
  }

  // Сотрудники разбиты на группы по системе оплаты — сперва "оклад", потом "сдельная",
  // и (если такие есть) сотрудники без назначенной системы оплаты последними — с строкой
  // промежуточного итога сразу после каждой непустой группы. computeTabelLayout() — общая
  // точка для buildTabelSheet() и archiveCurrentMonth(), чтобы раскладка (а значит и
  // диапазон защиты) не могла разойтись между пересборкой формы и архивацией месяца.
  var firstEmployeeRow = 18;
  var layout = computeTabelLayout(firstEmployeeRow);
  var groups = layout.groups;
  var totalRow = layout.totalRow;

  // AO<строка начала группы> — технический (скрытый) спиллящийся список табельных
  // номеров ЭТОЙ группы, отсортированный по фамилии (столбец B «ФИО» начинается с
  // фамилии). Раньше был один общий спилл на весь список — но с подытоговыми строками,
  // вклинивающимися между группами, один большой спилл просто не может перепрыгнуть
  // через подытог (спилл требует сплошного диапазона), поэтому теперь по формуле на
  // группу, каждая спиллится только в свой сплошной блок строк.
  // Условие активности — TRIM(LOWER(...))="да", а не точное "=да": без этого лишний
  // пробел или другой регистр в столбце D ("Активен") тихо исключал сотрудника из
  // формулы, хотя readEmployees() в API уже делает trim/lowercase и видел бы его —
  // несовпадение ровно так теряло новых сотрудников при «Пересоздать форму» (живой
  // пример: "Активен" = " да" с пробелом у 30-го сотрудника).
  // Столбец A в печатной форме — сквозной порядковый номер (1, 2, 3… по всем группам
  // подряд, строки подытогов не считаются), не таб. №: таб. номера после сортировки по
  // системе оплаты/фамилии больше не идут по порядку, а «№» в бланке — это позиция в
  // списке, не идентификатор сотрудника. B/C и ячейки дней ниже смотрят на $AO<row>
  // (реальный таб. №) по значению, а не на позицию строки — поэтому сортировка,
  // разбиение на группы и добавление нового сотрудника в любое место списка не путают
  // часы уже существующих сотрудников: каждая строка находит свой журнал по таб. №, где
  // бы она ни оказалась после пересборки, независимо от того, что показывает колонка A.
  var activeCond = 'TRIM(LOWER(' + SHEET_EMPLOYEES + '!$D$2:$D$1000))="да"';
  var employeeOrdinal = 0;

  for (var gi = 0; gi < groups.length; gi++) {
    var group = groups[gi];

    if (group.kind) {
      var conds = [activeCond].concat(paySystemKindConditions(group.kind)).join(",");
      var tabNumFilter = "FILTER(" + SHEET_EMPLOYEES + "!$A$2:$A$1000," + conds + ")";
      var nameFilter = "FILTER(TRIM(" + SHEET_EMPLOYEES + "!$B$2:$B$1000)," + conds + ")";
      sheet.getRange("AO" + group.startRow).setFormula(
        "=IFERROR(SORT(" + tabNumFilter + "," + nameFilter + ",TRUE),\"\")"
      );
    }

    for (var row = group.startRow; row <= group.endRow; row++) {
      employeeOrdinal++;
      sheet.getRange("A" + row).setValue(employeeOrdinal); // порядковый номер
      sheet.getRange("B" + row).setFormula('=IFERROR(VLOOKUP(AO' + row + ',' + SHEET_EMPLOYEES + '!$A:$C,2,FALSE),"")');
      sheet.getRange("C" + row).setFormula('=IFERROR(VLOOKUP(AO' + row + ',' + SHEET_EMPLOYEES + '!$A:$C,3,FALSE),"")');

      for (var col2 = 4; col2 <= 34; col2++) {
        var cl = columnToLetter(col2);
        var cell = cl + row;
        var dayCell = cl + "$14";
        sheet.getRange(cell).setFormula(
          '=IF(' + dayCell + '="","",IFERROR(INDEX(FILTER(' + SHEET_JOURNAL + '!$C:$C,' + SHEET_JOURNAL + '!$B:$B=$AO' + row + ',' + SHEET_JOURNAL + '!$A:$A=' + dayCell + '),COUNTIFS(' + SHEET_JOURNAL + '!$B:$B,$AO' + row + ',' + SHEET_JOURNAL + '!$A:$A,' + dayCell + ')),IF(WEEKDAY(' + dayCell + ')=1,"В","")))'
        );
      }

      sheet.getRange("AI" + row).setFormula("=COUNT(D" + row + ":AH" + row + ")");
      sheet.getRange("AJ" + row).setFormula("=SUM(D" + row + ":AH" + row + ")");
      sheet.getRange("AL" + row).setFormula(
        "=SUM(ARRAYFORMULA(IF(ISNUMBER(D" + row + ":AH" + row + "),IF(D" + row + ":AH" + row + ">" + paramNum("overtimeFrom") + ",D" + row + ":AH" + row + "-" + paramNum("overtimeFrom") + ",0),0)))"
      );
      if (group.kind === "оклад") {
        sheet.getRange("AM" + row).setFormula(okladSalaryFormula(row));
      } else if (group.kind === "сдельная") {
        // часовая ставка сдельщиков — в C<rateRow> (см. ниже), зарплата = ставка × часы
        sheet.getRange("AM" + row).setFormula('=IF($AO' + row + '="","",IFERROR(ROUND($C$' + (totalRow + 2) + '*AJ' + row + ',0),0))');
      }
    }

    if (group.subtotalRow !== null) {
      // Метка — в B, не в A: A на строке сотрудника держит порядковый номер и копирует
      // узкий 24px формат архивного столбца "№" (под 1-2 цифры), в который "Итого
      // (сдельная)" физически не влезает. В архиве подпись "ИТОГО:" и так стоит в B30
      // (широкий столбец ФИО, 236px) — используем то же место для единообразия.
      sheet.getRange("B" + group.subtotalRow).setValue(group.label);
      sheet.getRange("AI" + group.subtotalRow).setFormula("=SUM(AI" + group.startRow + ":AI" + group.endRow + ")");
      sheet.getRange("AJ" + group.subtotalRow).setFormula("=SUM(AJ" + group.startRow + ":AJ" + group.endRow + ")");
      sheet.getRange("AL" + group.subtotalRow).setFormula("=SUM(AL" + group.startRow + ":AL" + group.endRow + ")");
      sheet.getRange("AM" + group.subtotalRow).setFormula("=SUM(AM" + group.startRow + ":AM" + group.endRow + ")");
    }
  }

  // Грандтотал суммирует строки подытогов групп, а не весь диапазон firstEmployeeRow:
  // totalRow-1 — тот диапазон включает и сами подытоговые строки, и их сложение с
  // построчными суммами задвоило бы итог для каждой непустой группы.
  sheet.getRange("B" + totalRow).setValue("ИТОГО"); // B, не A — см. комментарий у подытогов групп выше
  sheet.getRange("AI" + totalRow).setFormula("=" + sumOfGroupSubtotals(groups, "AI"));
  sheet.getRange("AJ" + totalRow).setFormula("=" + sumOfGroupSubtotals(groups, "AJ"));
  sheet.getRange("AL" + totalRow).setFormula("=" + sumOfGroupSubtotals(groups, "AL"));
  sheet.getRange("AM" + totalRow).setFormula("=" + sumOfGroupSubtotals(groups, "AM"));

  // Сдельщики: общая сумма нарядов делится на общее число их часов — часовая ставка.
  var sdelnaya = groups.filter(function (g) { return g.kind === "сдельная"; })[0];
  sheet.getRange("B" + (totalRow + 1)).setValue("Сумма нарядов, тг");
  sheet.getRange("C" + (totalRow + 1)).setFormula("='" + SHEET_ORDERS + "'!$K$7").setNumberFormat("#,##0");
  sheet.getRange("B" + (totalRow + 2)).setValue("Часовая ставка сдельщиков, тг/ч");
  if (sdelnaya) {
    sheet.getRange("C" + (totalRow + 2)).setFormula("=IFERROR(C" + (totalRow + 1) + "/AJ" + sdelnaya.subtotalRow + ",0)").setNumberFormat("#,##0.00");
  } else {
    sheet.getRange("C" + (totalRow + 2)).setValue(0);
  }

  sheet.getRange("B" + (totalRow + 3)).setFormula("=" + paramCell("chief1Pos"));
  sheet.getRange("C" + (totalRow + 3)).setFormula("=" + paramCell("chief1Name"));
  sheet.getRange("B" + (totalRow + 6)).setFormula("=" + paramCell("chief2Pos"));
  sheet.getRange("C" + (totalRow + 6)).setFormula("=" + paramCell("chief2Name"));

  applyTabelConditionalFormatting(sheet, groups);
  copyArchiveFormatting(sheet, firstEmployeeRow, groups, totalRow);
  // В строке дней — только число дня (1, 2, 3…), а не дата: формат архива при копировании
  // мог затереть "d", поэтому задаём его снова после copyArchiveFormatting.
  sheet.getRange("D14:AH14").setNumberFormat("d");
  sheet.getRange("AM" + firstEmployeeRow + ":AM" + totalRow).setNumberFormat("#,##0");
  sheet.getRange("C" + (totalRow + 1)).setNumberFormat("#,##0");
  sheet.getRange("C" + (totalRow + 2)).setNumberFormat("#,##0.00");
  groups.forEach(function (g) {
    if (g.subtotalRow !== null) styleTotalRow(sheet.getRange(g.subtotalRow, 1, 1, 39));
  });
  styleTotalRow(sheet.getRange(totalRow, 1, 1, 39));
  applyDefaultFont(sheet);

  // ФИО (B) и должность (C) — в одну строку: без переноса, ширина по самому длинному значению.
  var tabelEmployees = readEmployees(true);
  sheet.getRange(firstEmployeeRow, 2, totalRow - firstEmployeeRow, 2).setWrap(false);
  fitColumnWidth(sheet, 2, tabelEmployees.map(function (e) { return e.name; }));
  fitColumnWidth(sheet, 3, tabelEmployees.map(function (e) { return e.position; }));

  // Скрываем строку 12 (технический разделитель в "Архив_исходный") и столбцы с
  // технической датой AN и техническими таб. номерами AO — после копирования формата
  // из архива, т.к. оно не трогает видимость строк/столбцов, но порядок на всякий
  // случай фиксируем явно.
  sheet.hideRows(12, 1);
  sheet.hideColumns(40, 4); // AN (40), AO (41), AP (42), AQ (43)

  // Повторный запуск "Пересоздать форму" не должен копить защиты поверх друг друга.
  var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var p = 0; p < existingProtections.length; p++) {
    existingProtections[p].remove();
  }
  var dataRange = sheet.getRange(firstEmployeeRow, 4, totalRow - firstEmployeeRow, 31); // D..AH на сотрудников и подытоги
  var protection = dataRange.protect().setDescription("Табель — печатная форма, собирается формулами из Журнала");
  protection.setWarningOnly(true);
}

/**
 * Зарплата окладника: оклад ÷ (норма дней месяца по графику × смена) × (часы AJ + переработка AL × (коэффициент − 1)).
 * График (5/6-дневная) и оклад берутся из «Сотрудников» по таб. № из AO. Оклад бывает
 * записан текстом с пробелами ("400 000") — приводим к числу. Пустой график = 5-дневная.
 */
function okladSalaryFormula(row) {
  var lookup = function (col) {
    return "VLOOKUP($AO" + row + "," + SHEET_EMPLOYEES + "!$A:$H," + col + ",FALSE)";
  };
  return (
    '=IF($AO' + row + '="","",IFERROR(ROUND(LET(' +
    "o," + lookup(6) + "," +
    'sal,IF(ISNUMBER(o),o,VALUE(REGEXREPLACE(o&"","[^0-9]",""))),' +
    "nd,IF(REGEXMATCH(" + lookup(8) + '&"","6"),$AQ$11,$AP$11),' +
    "sal/(nd*" + paramNum("shift") + ")*(AJ" + row + "+AL" + row + "*(" + paramNum("overtimeK") + "-1))" +
    "),0),0))"
  );
}

/**
 * Строит выражение "SUM(AI35,AI48,...)" из ячеек подытогов всех групп для данного
 * столбца (дней/часов/переработка) — грандтотал складывает готовые подытоги, а не
 * диапазон строк целиком (который иначе задвоил бы каждую непустую группу).
 */
function sumOfGroupSubtotals(groups, col) {
  var cells = [];
  for (var i = 0; i < groups.length; i++) {
    if (groups[i].subtotalRow !== null) cells.push(col + groups[i].subtotalRow);
  }
  if (cells.length === 0) return "0"; // сотрудников нет вовсе — суммировать нечего
  return "SUM(" + cells.join(",") + ")";
}

/**
 * Условие(я) FILTER для столбца «Система оплаты» под конкретную группу —
 * возвращает массив предикатов (они AND'ятся через запятую внутри FILTER).
 */
function paySystemKindConditions(kind) {
  var normalized = "TRIM(LOWER(" + SHEET_EMPLOYEES + "!$E$2:$E$1000))";
  if (kind === "оклад") return [normalized + '="оклад"'];
  if (kind === "сдельная") return [normalized + '="сдельная"'];
  return [normalized + '<>"оклад"', normalized + '<>"сдельная"']; // без категории: ни то ни другое (включая пусто)
}

/**
 * Раскладка строк «Табель» по группам системы оплаты: сперва "оклад", потом
 * "сдельная", потом сотрудники без назначенной системы оплаты — каждая группа
 * только если в ней есть хоть один активный сотрудник, со строкой промежуточного
 * итога сразу после. Общий для buildTabelSheet() и archiveCurrentMonth(), чтобы
 * раскладка не могла разойтись между пересборкой формы и архивацией месяца.
 */
function computeTabelLayout(firstEmployeeRow) {
  var employees = readEmployees(true);
  var isOklad = function (e) {
    return e.paySystem === "оклад";
  };
  var isSdelnaya = function (e) {
    return e.paySystem === "сдельная";
  };
  var groupDefs = [
    { kind: "оклад", label: "Итого (оклад)", count: employees.filter(isOklad).length },
    { kind: "сдельная", label: "Итого (сдельная)", count: employees.filter(isSdelnaya).length },
    {
      kind: "other",
      label: "Итого (без категории)",
      count: employees.filter(function (e) {
        return !isOklad(e) && !isSdelnaya(e);
      }).length,
    },
  ];

  var groups = [];
  var row = firstEmployeeRow;
  for (var g = 0; g < groupDefs.length; g++) {
    if (groupDefs[g].count === 0) continue;
    var startRow = row;
    var endRow = row + groupDefs[g].count - 1;
    row = endRow + 1;
    var subtotalRow = row;
    row++;
    groups.push({ kind: groupDefs[g].kind, label: groupDefs[g].label, startRow: startRow, endRow: endRow, subtotalRow: subtotalRow });
  }

  if (groups.length === 0) {
    // нет ни одного активного сотрудника — одна пустая строка без подытога,
    // чтобы сетка не исчезала совсем.
    groups.push({ kind: null, label: null, startRow: row, endRow: row, subtotalRow: null });
    row++;
  }

  return { groups: groups, employeeCount: employees.length, totalRow: row };
}

/**
 * Копирует реальное форматирование (ширины столбцов, высоты строк, шрифты,
 * границы, выравнивание) из листа "Архив_исходный" — вместо того чтобы
 * подбирать значения на глаз.Шапка (строки 1-17) копируется как есть;
 * строка сотрудника из архива используется как шаблон и тиражируется на
 * текущее количество сотрудников; строка ИТОГО и подписи ниже неё копируются
 * со сдвигом на разницу в номере строки ИТОГО между архивом и текущим листом.
 */
function copyArchiveFormatting(sheet, firstEmployeeRow, groups, totalRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archive = ss.getSheetByName(SHEET_ARCHIVE_SOURCE);
  if (!archive) return;

  var lastCol = 39; // A..AM — печатная форма; технические AN/AO (40/41) не трогаем
  var archiveFirstEmployeeRow = 18;
  var archiveTotalRow = 30;
  var archiveLastRow = 37; // захватывает строку ИТОГО и подписи B33/B36 с запасом

  for (var c = 1; c <= lastCol; c++) {
    sheet.setColumnWidth(c, archive.getColumnWidth(c));
  }
  for (var r = 1; r < firstEmployeeRow; r++) {
    sheet.setRowHeight(r, archive.getRowHeight(r));
  }

  archive
    .getRange(1, 1, firstEmployeeRow - 1, lastCol)
    .copyTo(sheet.getRange(1, 1, firstEmployeeRow - 1, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // Строка сотрудника из архива — шаблон на каждую группу отдельно (блоки сотрудников
  // не идут сплошным диапазоном, их разделяют строки подытогов). Подытоговые строки
  // оформлены как архивная строка ИТОГО (жирный шрифт, границы) — в архиве нет
  // собственного образца "подытога", но внешне ИТОГО подходит лучше всего.
  var employeeRowHeight = archive.getRowHeight(archiveFirstEmployeeRow);
  var subtotalRowHeight = archive.getRowHeight(archiveTotalRow);
  var employeeTemplate = archive.getRange(archiveFirstEmployeeRow, 1, 1, lastCol);
  var subtotalTemplate = archive.getRange(archiveTotalRow, 1, 1, lastCol);

  for (var g = 0; g < groups.length; g++) {
    var group = groups[g];
    var count = group.endRow - group.startRow + 1;
    employeeTemplate.copyTo(sheet.getRange(group.startRow, 1, count, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    for (var er = group.startRow; er <= group.endRow; er++) {
      sheet.setRowHeight(er, employeeRowHeight);
    }
    if (group.subtotalRow !== null) {
      subtotalTemplate.copyTo(sheet.getRange(group.subtotalRow, 1, 1, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.setRowHeight(group.subtotalRow, subtotalRowHeight);
    }
  }

  var tailRows = archiveLastRow - archiveTotalRow + 1;
  archive
    .getRange(archiveTotalRow, 1, tailRows, lastCol)
    .copyTo(sheet.getRange(totalRow, 1, tailRows, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  for (var tr = 0; tr < tailRows; tr++) {
    sheet.setRowHeight(totalRow + tr, archive.getRowHeight(archiveTotalRow + tr));
  }
}

// groups — блоки сотрудников (без строк подытогов): цвет часов/кодов должен
// раскрашивать только реальные отметки, а не формулы сумм в подытоговых строках.
function applyTabelConditionalFormatting(sheet, groups) {
  var ranges = groups.map(function (g) {
    return sheet.getRange("D" + g.startRow + ":AH" + g.endRow);
  });
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(D18),D18>0,D18<=7.5)')
      .setBackground("#BF4F14")
      .setRanges(ranges)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(D18),D18>7.5,D18<=8)')
      .setBackground("#AFD2D3")
      .setRanges(ranges)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(D18),D18>=8.5,D18<=24)')
      .setBackground("#4D94D8")
      .setRanges(ranges)
      .build(),
  ];
  for (var i = 0; i < CODE_TABLE.length; i++) {
    var code = CODE_TABLE[i][0];
    var color = CODE_TABLE[i][2];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=D18="' + code + '"')
        .setBackground(color)
        .setRanges(ranges)
        .build()
    );
  }
  sheet.setConditionalFormatRules(rules);
}

function columnToLetter(column) {
  var letter = "";
  while (column > 0) {
    var rem = (column - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

// ===================== архив месяца =====================

function archiveCurrentMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TABEL);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист "' + SHEET_TABEL + '" не найден.');
    return;
  }
  var firstDay = sheet.getRange("AN11").getValue();
  var monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
  var name = "Табель_" + monthNames[firstDay.getMonth()] + "_" + firstDay.getFullYear();

  if (ss.getSheetByName(name)) {
    SpreadsheetApp.getUi().alert('Лист "' + name + '" уже существует.');
    return;
  }

  var copy = sheet.copyTo(ss);
  copy.setName(name);
  var values = copy.getDataRange().getValues();
  copy.getDataRange().setValues(values); // заменяем формулы статическими значениями
  var layout = computeTabelLayout(18); // та же раскладка, что и при пересборке формы
  var protection = copy.getRange(18, 4, layout.totalRow - 18, 31).protect();
  protection.setWarningOnly(true);
  copy.hideSheet(); // архив месяца не должен мешаться в списке вкладок

  // «Наряды» архивируются вместе с «Табелем»: тоже значениями и в скрытый лист.
  var archivedNames = ['"' + name + '"'];
  var orders = ss.getSheetByName(SHEET_ORDERS);
  var ordersName = "Наряды_" + monthNames[firstDay.getMonth()] + "_" + firstDay.getFullYear();
  if (orders && !ss.getSheetByName(ordersName)) {
    var ordersCopy = orders.copyTo(ss);
    ordersCopy.setName(ordersName);
    ordersCopy.getDataRange().setValues(ordersCopy.getDataRange().getValues());
    ordersCopy.clearConditionalFormatRules();
    ordersCopy.hideSheet();
    archivedNames.push('"' + ordersName + '"');
  }

  SpreadsheetApp.getUi().alert('Месяц сохранён в скрытые листы ' + archivedNames.join(", ") + '. Открыть их можно через "Вид → Скрытые листы".');
}

// ===================== «Наряды» и «Вредники» =====================

/**
 * Лист «Наряды» — «Объем выполненных работ». Строки (название из столбца B, вес из D)
 * подтягиваются IMPORTRANGE из файла данных ORDERS_SOURCE_ID, лист — по английскому
 * названию месяца «Табеля» (G5, можно переписать вручную). Цена за кг (G) вводится вручную,
 * «всего» (H) = вес × цена × кол-во. Таблица — минимум ORDERS_MIN_ROWS строк; если в данных
 * строк больше, она растёт, а ИТОГО и подписанты сдвигаются вниз формулами (положение
 * считается в K5). Служебные ячейки — в J:O. Общая сумма — K7 (её берёт «Табель»).
 */
function buildOrdersSheet(sheet) {
  assertFormulasSupported(sheet);
  var T = "'" + SHEET_TABEL + "'!";
  var money = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@';
  var first = 9;
  var last = first + ORDERS_CAP - 1;

  sheet.getRange("F2").setFormula("=" + paramCell("headPos") + '&" "&' + paramCell("org"));
  sheet.getRange("F4").setFormula("=" + paramText("headName", "_______________"));
  sheet.getRange("F2:H4").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange("B4").setFormula("=" + paramText("unit"));
  sheet.getRange("B5").setValue("Дата:");
  sheet.getRange("C5").setValue(new Date()).setNumberFormat("dd.mm.yyyy").setHorizontalAlignment("left");
  sheet.getRange("F5").setValue("Лист данных:").setHorizontalAlignment("right");
  sheet.getRange("G5").setFormula(
    '=CHOOSE(MONTH(' + T + '$AN$11),"January","February","March","April","May","June","July","August","September","October","November","December")'
  ).setHorizontalAlignment("center");
  sheet.getRange("C6").setValue("Объем выполненных работ").setFontWeight("bold");
  sheet.getRange("B7").setFormula(
    '=IF(ISERROR($K$9),"Нет данных: проверьте название листа в G5 и разрешите доступ к файлу (выделите K9 → «Разрешить доступ»)",IF($K$3>' + ORDERS_CAP + ',"Строк в данных больше ' + ORDERS_CAP + ' — лишние не показаны",""))'
  ).setFontColor("#C00000");

  sheet.getRange("B8:H8")
    .setValues([["№", "Наименование конструкции и изделий", "ед.изм.", "кол-во", "вес, кг", "цена за кг, тенге", "всего, тенге"]])
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Служебные ячейки.
  sheet.getRange("J1:J7").setValues([["ID файла данных"], ["Лист данных"], ["Строк в данных"], ["Строк в таблице"], ["Строка ИТОГО"], ["Общий вес, кг"], ["Общая сумма, тг"]]);
  sheet.getRange("K1").setValue(ORDERS_SOURCE_ID);
  sheet.getRange("K2").setFormula("=$G$5");
  sheet.getRange("K3").setFormula("=IF(ISERROR($K$9),0,COUNTA($K$9:$K))");
  sheet.getRange("K4").setFormula("=MAX(" + ORDERS_MIN_ROWS + ",$K$3)");
  sheet.getRange("K5").setFormula("=" + first + "+$K$4");
  sheet.getRange("K6").setFormula("=IF(ISERROR($K$9),0,SUM(ARRAYFORMULA(IFERROR(VALUE($M$9:$M),0))))");
  sheet.getRange("K7").setFormula("=SUM($O$" + first + ":$O$" + last + ")").setNumberFormat("#,##0");
  sheet.getRange("K8:M8").setValues([["Название", "Заказчик", "Вес"]]);
  sheet.getRange("O8").setValue("Сумма");
  sheet.getRange("K9").setFormula(
    '=LET(d,IMPORTRANGE($K$1,$K$2&"!B2:D"),FILTER(d,INDEX(d,,1)<>""))'
  );

  var main = [];
  var hCol = [];
  var oCol = [];
  for (var r = first; r <= last; r++) {
    var idx = "ROW()-8";
    var isData = idx + "<=$K$3";
    main.push([
      '=IF(' + idx + '<=$K$4,' + idx + ',"")',
      '=IF(' + isData + ',INDEX($K$9:$K,' + idx + '),IF(ROW()=$K$5,"ИТОГО:",IF(ROW()=$K$5+3,' + paramCell("chief1Pos") + ',IF(ROW()=$K$5+6,' + paramCell("chief2Pos") + ',""))))',
      '=IF(' + isData + ',"шт",IF(ROW()=$K$5+3,' + paramCell("chief1Name") + ',IF(ROW()=$K$5+6,' + paramCell("chief2Name") + ',"")))',
      '=IF(' + isData + ',1,"")',
      '=IF(' + isData + ',IFERROR(VALUE(INDEX($M$9:$M,' + idx + ')),0),IF(ROW()=$K$5,$K$6,""))',
    ]);
    hCol.push(['=IF(' + isData + ',$O' + r + ',IF(ROW()=$K$5,$K$7,""))']);
    oCol.push(['=IF(' + isData + ',N($F' + r + ')*N($G' + r + ')*N($E' + r + '),0)']);
  }
  sheet.getRange(first, 2, ORDERS_CAP, 5).setFormulas(main);
  sheet.getRange(first, 8, ORDERS_CAP, 1).setFormulas(hCol);
  sheet.getRange(first, 15, ORDERS_CAP, 1).setFormulas(oCol);

  var body = ORDERS_CAP;
  sheet.getRange(first, 2, body, 1).setHorizontalAlignment("center");
  sheet.getRange(first, 4, body, 2).setHorizontalAlignment("center");
  sheet.getRange(first, 6, body, 1).setNumberFormat("#,##0").setHorizontalAlignment("right");
  sheet.getRange(first, 7, body, 1).setNumberFormat("#,##0.00").setHorizontalAlignment("center");
  sheet.getRange(first, 8, body, 1).setNumberFormat(money).setHorizontalAlignment("center");
  sheet.getRange(first, 15, body, 1).setNumberFormat("#,##0");
  sheet.getRange("J1:O" + last).setFontColor("#888888");

  // Строка ИТОГО «ездит» вместе с данными — выделяем её условным форматированием.
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=ROW()=$K$5")
      .setBold(true)
      .setBackground(TOTAL_FILL)
      .setRanges([sheet.getRange("B" + first + ":H" + last)])
      .build(),
  ]);

  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 380);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 30);
  sheet.setColumnWidths(10, 6, 110);
  sheet.setFrozenRows(8);
  applyDefaultFont(sheet);
  SpreadsheetApp.flush();
  refreshOrdersLayout(sheet);
}

/**
 * Рамка «Нарядов» ровно по текущему числу строк (минимум 50 + строка ИТОГО): добавленные
 * строки выглядят так же, как верхние, а пустые ниже таблицы без рамки. Условным
 * форматированием границы не рисуются, поэтому обновляется скриптом — при пересборке,
 * при открытии таблицы и из меню «Обновить наряды».
 */
function refreshOrdersLayout(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  if (!sheet) return;
  var n = Number(sheet.getRange("K4").getValue());
  if (!(n >= ORDERS_MIN_ROWS)) n = ORDERS_MIN_ROWS;
  n = Math.min(n, ORDERS_CAP);
  var first = 9;
  sheet.getRange(first, 2, ORDERS_CAP, 7).setBorder(false, false, false, false, false, false);
  sheet.getRange(8, 2, n + 2, 7).setBorder(true, true, true, true, true, true); // шапка + n строк + ИТОГО
}

function ensureHarmfulSheet(ss, log) {
  ensureParamsSheet(ss);
  if (ss.getSheetByName(SHEET_HARMFUL)) {
    log.push('Лист "' + SHEET_HARMFUL + '" уже существует — пропуск (используйте "Обновить «Вредники»").');
    return;
  }
  var sheet = ss.insertSheet(SHEET_HARMFUL);
  buildHarmfulSheet(sheet);
  log.push('Лист "' + SHEET_HARMFUL + '" создан.');
}

function rebuildHarmfulSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureParamsSheet(ss);
  var sheet = ss.getSheetByName(SHEET_HARMFUL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HARMFUL);
  } else {
    sheet.clear();
  }
  buildHarmfulSheet(sheet);
  ss.toast("Лист «Вредники» пересобран.");
}

/**
 * Ведомость выдачи молока. В список попадают активные сотрудники с «Вредность» = "да"
 * (лист «Сотрудники», столбец G), по фамилии. «Отработано дней» — сколько числовых
 * значений часов в «Табеле» у сотрудника за период с C9 по E9 (число месяца), «литров» =
 * дни × «Молоко» из листа «Параметры» — та же математика, что в исходном файле. Число строк
 * берётся по сотрудникам на момент сборки: добавили вредника — «Обновить «Вредники»».
 */
function buildHarmfulSheet(sheet) {
  assertFormulasSupported(sheet);
  var harmful = readEmployees(true).filter(function (e) { return e.harmful; });
  var count = Math.max(harmful.length, 1);
  var T = "'" + SHEET_TABEL + "'!";
  var E = SHEET_EMPLOYEES + "!";
  var cond = 'TRIM(LOWER(' + E + '$D$2:$D$1000))="да",TRIM(LOWER(' + E + '$G$2:$G$1000))="да"';
  var sortedNames = 'SORT(FILTER(TRIM(' + E + '$B$2:$B$1000),' + cond + '))';

  // Раскладка и оформление — как в «Вредники» исходного файла: шапка в столбце A с
  // объединениями, жёлтые ячейки месяца/периода, светло-жёлтая шапка таблицы, рамки.
  var widths = [24, 206, 174, 97, 62, 62, 79, 89];
  for (var c = 0; c < widths.length; c++) sheet.setColumnWidth(c + 1, widths[c]);

  sheet.getRange("A2:F2").merge().setHorizontalAlignment("center")
    .setValue("ВЕДОМОСТЬ\nвыдачи молока (или равноценных пищевых продуктов)");
  sheet.getRange("A4:C4").merge().setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setFormula("=" + paramText("org", "Наименование организации: "));
  sheet.getRange("A6:C6").merge().setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setFormula("=" + paramText("unit", "Подразделение: "));
  sheet.getRange("A8").setValue("Месяц");
  sheet.getRange("C8").setFormula(
    '=SWITCH(MONTH(' + T + '$AN$11),1,"январь",2,"февраль",3,"март",4,"апрель",5,"май",6,"июнь",7,"июль",8,"август",9,"сентябрь",10,"октябрь",11,"ноябрь",12,"декабрь")'
  );
  sheet.getRange("A9").setValue("Период с (число месяца):");
  sheet.getRange("C9").setValue(1);
  sheet.getRange("D9").setValue("по (число месяца):").setHorizontalAlignment("right").setFontWeight("bold");
  sheet.getRange("E9").setFormula("=DAY(EOMONTH(" + T + "$AN$11,0))");
  sheet.getRange("C8:C9").setBackground("#FFFF99").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("E9").setBackground("#FFFF99").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A9").setFontWeight("bold");
  sheet.getRange("A10:F11").merge().setWrap(true)
    .setValue("Основание: Перечень профессий и должностей, имеющих право на получение молока за работу во вредных условиях труда.");

  var head = sheet.getRange("A13:H13");
  head.setValues([["№", "Фамилия имя отчество", "Должность", "Отработано за период (дней)", "Дата выдачи", "Кол-во (литров)", "Подпись работника", "Подпись выдавшего"]]);
  head.setWrap(true);
  var nums = sheet.getRange("A14:H14");
  nums.setValues([[1, 2, 3, 4, 4, 5, 5, 6]]);
  sheet.getRange("A13:H14").setBackground("#FFFFCC").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("A13:H14").setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(13, 48);
  sheet.setRowHeight(14, 21);

  var first = 15;
  var last = first + count - 1;
  var dayRow = T + "$D$14:$AH$14";
  for (var r = first; r <= last; r++) {
    sheet.getRange("A" + r).setFormula('=IF(B' + r + '="","",' + (r - first + 1) + ")");
    sheet.getRange("B" + r).setFormula('=IFERROR(INDEX(' + sortedNames + ',' + (r - first + 1) + '),"")');
    sheet.getRange("C" + r).setFormula('=IFERROR(INDEX(FILTER(' + E + '$C$2:$C$1000,ARRAYFORMULA(TRIM(' + E + '$B$2:$B$1000))=B' + r + '),1),"")');
    // Табель: в строке 14 — даты, поэтому границы периода переводим в даты месяца.
    sheet.getRange("D" + r).setFormula(
      '=IF(B' + r + '="","",IFERROR(SUMPRODUCT(ISNUMBER(INDEX(' + T + '$D$18:$AH$200,MATCH(B' + r + ',' + T + '$B$18:$B$200,0),0))' +
      '*(' + dayRow + '>=DATE(YEAR(' + T + '$AN$11),MONTH(' + T + '$AN$11),$C$9))' +
      '*(' + dayRow + '<=DATE(YEAR(' + T + '$AN$11),MONTH(' + T + '$AN$11),$E$9))),0))'
    );
    sheet.getRange("F" + r).setFormula('=IF(B' + r + '="","",D' + r + "*" + paramNum("milk") + ")");
    sheet.setRowHeight(r, 20);
  }
  var totalRow = last + 1;
  var body = sheet.getRange("A" + first + ":H" + totalRow);
  body.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle");
  sheet.getRange("A" + first + ":A" + last).setHorizontalAlignment("center");
  sheet.getRange("B" + first + ":C" + totalRow).setHorizontalAlignment("left").setWrap(false);
  fitColumnWidth(sheet, 2, harmful.map(function (e) { return e.name; }));
  fitColumnWidth(sheet, 3, harmful.map(function (e) { return e.position; }));
  sheet.getRange("D" + first + ":H" + totalRow).setHorizontalAlignment("center").setFontWeight("bold");
  sheet.getRange("F" + first + ":F" + totalRow).setNumberFormat("0.0");

  sheet.getRange("B" + totalRow).setValue("ИТОГО:");
  sheet.getRange("D" + totalRow).setFormula("=SUM(D" + first + ":D" + last + ")");
  sheet.getRange("F" + totalRow).setFormula("=SUM(F" + first + ":F" + last + ")");
  sheet.setRowHeight(totalRow, 20);
  styleTotalRow(sheet.getRange("A" + totalRow + ":H" + totalRow));

  sheet.getRange("B" + (totalRow + 2)).setFormula('="Итого выдано: "&ROUND(F' + totalRow + ',1)&" литров"');
  sheet.getRange("B" + (totalRow + 4)).setValue("Выдал:");
  sheet.getRange("B" + (totalRow + 6)).setValue("Проверил:");
  sheet.getRange("B" + (totalRow + 8)).setValue("Дата");

  applyDefaultFont(sheet);
}

// ===================== чтение справочников =====================

function readEmployees(onlyActive) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  var rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 8).getValues();
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var active = String(r[3]).trim().toLowerCase() === "да";
    if (onlyActive && !active) continue;
    // Систему оплаты назначает пользователь сам в таблице; пусто — пока не назначена
    // (в приложении такой сотрудник просто не попадает ни под один конкретный фильтр).
    var paySystem = r[4] ? String(r[4]).trim().toLowerCase() : "";
    result.push({ tabNum: r[0], name: r[1], position: r[2], active: active, paySystem: paySystem, salary: r[5] === "" ? null : r[5], harmful: String(r[6]).trim().toLowerCase() === "да", workDays: String(r[7]).indexOf("6") >= 0 ? 6 : 5 });
  }
  return result;
}

function readCodes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CODES);
  var rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 4).getValues();
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    result.push({ code: r[0], label: r[1], color: r[2], paid: String(r[3]).trim().toLowerCase() === "да" });
  }
  return result;
}

// ===================== веб-API =====================

// "master" — полный доступ (чтение + запись через doPost), PIN в Script Properties.
// "timekeeper" (табельщик) — чтение + запись, но processRecord() разрешает писать только
// сотрудникам с paySystem="сдельная"; отдельный TIMEKEEPER_PIN, необязателен.
// "viewer" — только чтение (doGet), отдельный VIEW_PIN; необязателен.
// Все дополнительные PIN необязательны — если не заданы в Script Properties, роли просто нет.
// null — PIN не подошёл ни под одну роль.
function resolveRole(pin) {
  var props = PropertiesService.getScriptProperties();
  var masterPin = props.getProperty("PIN");
  var timekeeperPin = props.getProperty("TIMEKEEPER_PIN");
  var viewPin = props.getProperty("VIEW_PIN");
  if (masterPin && String(pin) === String(masterPin)) return "master";
  if (timekeeperPin && String(pin) === String(timekeeperPin)) return "timekeeper";
  if (viewPin && String(pin) === String(viewPin)) return "viewer";
  return null;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var params = e.parameter || {};
    var role = resolveRole(params.pin);
    if (!role) {
      return jsonResponse({ ok: false, error: "Доступ запрещён" });
    }

    if (params.action === "init") {
      return jsonResponse({
        ok: true,
        role: role,
        employees: readEmployees(true),
        codes: readCodes(),
        normHours: NORM_HOURS,
      });
    }

    if (params.action === "month") {
      return jsonResponse({ ok: true, entries: readMonthEntries(params.month) });
    }

    return jsonResponse({ ok: false, error: "Неизвестное действие" });
  } catch (err) {
    return jsonResponse({ ok: false, error: "Внутренняя ошибка" });
  }
}

function readMonthEntries(month) {
  // month: "2026-11"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_JOURNAL);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var latest = {}; // "date__tab" -> запись

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var dateKey = formatDateKey(r[0]);
    if (dateKey.slice(0, 7) !== month) continue;
    var key = dateKey + "__" + r[1];
    latest[key] = {
      date: dateKey,
      tabNum: r[1],
      value: r[2],
      who: r[3],
      ts: r[4] ? formatIso(r[4]) : "",
      device: r[5],
      uuid: r[6],
    };
  }

  var result = [];
  for (var k in latest) result.push(latest[k]);
  return result;
}

function formatDateKey(date) {
  if (Object.prototype.toString.call(date) !== "[object Date]") date = new Date(date);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatIso(date) {
  if (Object.prototype.toString.call(date) !== "[object Date]") date = new Date(date);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function isValidHours(value) {
  if (typeof value !== "number" || isNaN(value)) return false;
  if (value < 0 || value > 24) return false;
  return Math.round(value * 2) === value * 2; // шаг 0.5
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var records = body.records || body; // поддержка {pin, records:[...]} и просто массива с pin в каждой записи
    var pin = body.pin;

    if (!Array.isArray(records)) {
      return jsonResponse({ ok: false, error: "Неверный формат запроса" });
    }
    var role = resolveRole(pin);
    if (!role) {
      return jsonResponse({ ok: false, error: "Доступ запрещён" });
    }
    if (role === "viewer") {
      return jsonResponse({ ok: false, error: "Доступ только для чтения" });
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var results = records.map(function (rec) {
        return processRecord(rec, role);
      });
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({ ok: true, results: results });
  } catch (err) {
    return jsonResponse({ ok: false, error: "Внутренняя ошибка" });
  }
}

function processRecord(rec, role) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_JOURNAL);

  if (!rec.uuid) {
    return { uuid: rec.uuid, ok: false, error: "Отсутствует uuid" };
  }
  if (findRowByUuid(sheet, rec.uuid)) {
    return { uuid: rec.uuid, ok: true, duplicate: true };
  }

  var employees = readEmployees(false);
  var employee = employees.filter(function (emp) {
    return String(emp.tabNum) === String(rec.id);
  })[0];
  if (!employee) {
    return { uuid: rec.uuid, ok: false, error: "Таб. № не найден" };
  }
  // Табельщик пишет только сдельщикам — граница реальная (сервер), а не только в интерфейсе.
  if (role === "timekeeper" && employee.paySystem !== "сдельная") {
    return { uuid: rec.uuid, ok: false, error: "Табельщик может отмечать только сдельщиков" };
  }

  var value = rec.value;
  if (value !== "" && typeof value !== "number") {
    var codes = readCodes();
    var code = codes.filter(function (c) {
      return c.code === value;
    })[0];
    if (!code) {
      return { uuid: rec.uuid, ok: false, error: "Код не найден" };
    }
  } else if (typeof value === "number" && !isValidHours(value)) {
    return { uuid: rec.uuid, ok: false, error: "Часы должны быть от 0 до 24 с шагом 0.5" };
  }

  var who = Session.getActiveUser().getEmail() || "мастер";
  // Полдень UTC — защита от сдвига на соседний календарный день при несовпадении
  // часового пояса проекта и часового пояса таблицы (см. buildTabelSheet и DECISIONS.md).
  var date = new Date(rec.date + "T12:00:00Z");
  sheet.appendRow([date, rec.id, value, who, new Date(), rec.deviceId || "", rec.uuid]);

  return { uuid: rec.uuid, ok: true };
}

function findRowByUuid(sheet, uuid) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var uuids = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
  for (var i = 0; i < uuids.length; i++) {
    if (uuids[i][0] === uuid) return true;
  }
  return false;
}
