/**
 * Табель — backend на Google Apps Script, привязанный к таблице.
 * Публикуется как веб-приложение (Выполнять от имени: я; Доступ: все).
 * Сам скрипт таблицу не публикует и ничего в неё не пишет при деплое — см. README.md.
 */

var SHEET_EMPLOYEES = "Сотрудники";
var SHEET_CODES = "Коды";
var SHEET_JOURNAL = "Журнал";
var SHEET_TABEL = "Табель";
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

// ===================== меню =====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Табель")
    .addItem("Сохранить месяц в архив", "archiveCurrentMonth")
    .addItem("Пересоздать форму", "rebuildTabelSheet")
    .addToUi();
}

// ===================== setupSheets: разовая миграция =====================

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  renameOldSheetToArchive(ss, log);
  var employeesCreated = ensureEmployeesSheet(ss, log);
  ensureCodesSheet(ss, log);
  ensureJournalSheet(ss, log);
  ensureTabelSheet(ss, log, employeesCreated);

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
  sheet.getRange(1, 1, 1, 4).setValues([["Таб. №", "ФИО", "Должность", "Активен"]]);
  sheet.setFrozenRows(1);

  var source = ss.getSheetByName(SHEET_ARCHIVE_SOURCE);
  var rows = [];
  if (source) {
    var data = source.getRange(18, 1, 12, 3).getValues(); // A:C, строки 18-29
    for (var i = 0; i < data.length; i++) {
      var name = data[i][1];
      var position = data[i][2];
      if (!name) continue;
      rows.push([i + 1, name, position, "да"]);
    }
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  sheet.autoResizeColumns(1, 4);
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
  var existing = ss.getSheetByName(SHEET_TABEL);
  if (existing) {
    log.push('Лист "' + SHEET_TABEL + '" уже существует — пропуск (используйте "Пересоздать форму" для обновления формул).');
    return;
  }
  var sheet = ss.insertSheet(SHEET_TABEL);
  buildTabelSheet(sheet);
  log.push('Лист "' + SHEET_TABEL + '" создан.');
}

function rebuildTabelSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TABEL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TABEL);
  } else {
    sheet.clear();
  }
  buildTabelSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast("Лист «Табель» пересобран.");
}

/**
 * Собирает лист «Табель»: шапку, формулы сетки дней и итогов, условное форматирование.
 * Формулы пишутся в английской нотации — это требование Apps Script (SpreadsheetApp
 * принимает setFormula только с английскими именами функций); пользователю в интерфейсе
 * таблицы они всё равно показываются по-русски, если у таблицы русская локаль.
 */
function buildTabelSheet(sheet) {
  sheet.getRange("A3").setValue('ТОО "ПК Алтын Таға"');
  sheet.getRange("A5").setValue("Цех металообработки");
  sheet.getRange("P10").setValue("ТАБЕЛЬ УЧЕТА РАБОЧЕГО ВРЕМЕНИ И РАСЧЕТА ЗАРАБОТНОЙ ПЛАТЫ");

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
  if (sheet.getMaxColumns() < 40) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 40 - sheet.getMaxColumns());
  }
  sheet.getRange("AN11").setValue(first).setNumberFormat("dd.mm.yyyy");
  sheet.hideColumns(40, 1); // AN = столбец 40

  sheet.getRange("O11").setFormula(
    '="За "&SWITCH(MONTH($AN$11),1,"январь",2,"февраль",3,"март",4,"апрель",5,"май",6,"июнь",7,"июль",8,"август",9,"сентябрь",10,"октябрь",11,"ноябрь",12,"декабрь")&" "&YEAR($AN$11)&" года"'
  );

  sheet.getRange("AI3").setValue("Утверждаю:");
  sheet.getRange("AI4").setValue("Директор");
  sheet.getRange("AI5").setValue("Нұрмағамбет Ә. Е.");

  sheet.getRange("AI13").setValue("Итого отработано за месяц");
  sheet.getRange("AJ13").setValue("часов");
  sheet.getRange("AK13").setValue("ремонтных");
  sheet.getRange("AL13").setValue("переработка");
  sheet.getRange("AM13").setValue("зарплата");

  // Строка 14 — числа дней месяца; строка 13 — пометка выходных.
  for (var col = 4; col <= 34; col++) { // D..AH
    var colLetter = columnToLetter(col);
    sheet
      .getRange(colLetter + "14")
      .setFormula("=IF(MONTH($AN$11+COLUMN()-4)=MONTH($AN$11),$AN$11+COLUMN()-4,\"\")")
      .setNumberFormat("d");
    sheet.getRange(colLetter + "13").setFormula('=IF(' + colLetter + '14="","",IF(WEEKDAY(' + colLetter + '14)=1,"В",""))');
  }

  // Сотрудники: A18 — спиллящийся список активных табельных номеров, B/C — ВПР.
  sheet.getRange("A18").setFormula(
    '=IFERROR(FILTER(' + SHEET_EMPLOYEES + '!$A$2:$A$1000,' + SHEET_EMPLOYEES + '!$D$2:$D$1000="да"),"")'
  );
  for (var row = 18; row <= 29; row++) {
    sheet.getRange("B" + row).setFormula('=IFERROR(VLOOKUP(A' + row + ',' + SHEET_EMPLOYEES + '!$A:$C,2,FALSE),"")');
    sheet.getRange("C" + row).setFormula('=IFERROR(VLOOKUP(A' + row + ',' + SHEET_EMPLOYEES + '!$A:$C,3,FALSE),"")');

    for (var col2 = 4; col2 <= 34; col2++) {
      var cl = columnToLetter(col2);
      var cell = cl + row;
      var dayCell = cl + "$14";
      sheet.getRange(cell).setFormula(
        '=IF(' + dayCell + '="","",IFERROR(INDEX(FILTER(' + SHEET_JOURNAL + '!$C:$C,' + SHEET_JOURNAL + '!$B:$B=$A' + row + ',' + SHEET_JOURNAL + '!$A:$A=' + dayCell + '),COUNTIFS(' + SHEET_JOURNAL + '!$B:$B,$A' + row + ',' + SHEET_JOURNAL + '!$A:$A,' + dayCell + ')),IF(WEEKDAY(' + dayCell + ')=1,"В","")))'
      );
    }

    sheet.getRange("AI" + row).setFormula("=COUNT(D" + row + ":AH" + row + ")");
    sheet.getRange("AJ" + row).setFormula("=SUM(D" + row + ":AH" + row + ")");
    sheet.getRange("AL" + row).setFormula(
      "=SUM(ARRAYFORMULA(IF(ISNUMBER(D" + row + ":AH" + row + "),IF(D" + row + ":AH" + row + ">" + NORM_HOURS + ",D" + row + ":AH" + row + "-" + NORM_HOURS + ",0),0)))"
    );
  }

  sheet.getRange("A30").setValue("ИТОГО");
  sheet.getRange("AI30").setFormula("=SUM(AI18:AI29)");
  sheet.getRange("AJ30").setFormula("=SUM(AJ18:AJ29)");
  sheet.getRange("AL30").setFormula("=SUM(AL18:AL29)");

  sheet.getRange("B33").setValue("Нач. цеха");
  sheet.getRange("C33").setValue("Утебалиев А.В.");
  sheet.getRange("B36").setValue("Гл. инженер");
  sheet.getRange("C36").setValue("Минбаев К.К.");

  applyTabelConditionalFormatting(sheet);

  // Лист целиком со скрытыми линиями сетки (hideGridlines) — без явных границ таблица
  // выглядит совсем без линий. Обводим шапку с номерами дней, саму сетку отметок и
  // строку ИТОГО: A13:AM30 (включает A..C ФИО/должность, D..AH дни, AI..AM итоги).
  var tableRange = sheet.getRange("A13:AM30");
  tableRange.setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);

  // Ширины столбцов и выравнивание — компактная сетка 31 дня + читаемые ФИО/итоги.
  sheet.setColumnWidth(1, 40); // A — №
  sheet.setColumnWidth(2, 230); // B — ФИО
  sheet.setColumnWidth(3, 130); // C — Должность
  sheet.setColumnWidths(4, 31, 30); // D..AH — дни
  sheet.setColumnWidths(35, 5, 65); // AI..AM — итоги
  sheet.getRange("D13:AH29").setHorizontalAlignment("center");
  sheet.getRange("D14:AH14").setFontWeight("bold");
  sheet.getRange("A30:AM30").setFontWeight("bold");
  sheet.getRange("A18:A29").setHorizontalAlignment("center");

  var dataRange = sheet.getRange(18, 4, 12, 31); // D18:AH29
  var protection = dataRange.protect().setDescription("Табель — печатная форма, собирается формулами из Журнала");
  protection.setWarningOnly(true);
}

function applyTabelConditionalFormatting(sheet) {
  var range = sheet.getRange("D18:AH29");
  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(D18),D18>0,D18<=7.5)')
      .setBackground("#BF4F14")
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(D18),D18>7.5,D18<=8)')
      .setBackground("#AFD2D3")
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER(D18),D18>=8.5,D18<=24)')
      .setBackground("#4D94D8")
      .setRanges([range])
      .build(),
  ];
  for (var i = 0; i < CODE_TABLE.length; i++) {
    var code = CODE_TABLE[i][0];
    var color = CODE_TABLE[i][2];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=D18="' + code + '"')
        .setBackground(color)
        .setRanges([range])
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
  var protection = copy.getRange(18, 4, 12, 31).protect();
  protection.setWarningOnly(true);

  SpreadsheetApp.getUi().alert('Месяц сохранён в лист "' + name + '".');
}

// ===================== чтение справочников =====================

function readEmployees(onlyActive) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  var rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 4).getValues();
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var active = String(r[3]).trim().toLowerCase() === "да";
    if (onlyActive && !active) continue;
    result.push({ tabNum: r[0], name: r[1], position: r[2], active: active });
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

function checkPin(pin) {
  var expected = PropertiesService.getScriptProperties().getProperty("PIN");
  return expected && String(pin) === String(expected);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var params = e.parameter || {};
    if (!checkPin(params.pin)) {
      return jsonResponse({ ok: false, error: "Доступ запрещён" });
    }

    if (params.action === "init") {
      return jsonResponse({
        ok: true,
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
    if (!checkPin(pin)) {
      return jsonResponse({ ok: false, error: "Доступ запрещён" });
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var results = records.map(processRecord);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({ ok: true, results: results });
  } catch (err) {
    return jsonResponse({ ok: false, error: "Внутренняя ошибка" });
  }
}

function processRecord(rec) {
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
