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
  sheet.getRange(1, 1, 1, 5).setValues([["Таб. №", "ФИО", "Должность", "Активен", "Система оплаты"]]);
  sheet.setFrozenRows(1);

  var source = ss.getSheetByName(SHEET_ARCHIVE_SOURCE);
  var rows = [];
  if (source) {
    var data = source.getRange(18, 1, 12, 3).getValues(); // A:C, строки 18-29
    for (var i = 0; i < data.length; i++) {
      var name = data[i][1];
      var position = data[i][2];
      if (!name) continue;
      rows.push([i + 1, name, position, "да", ""]); // систему оплаты назначает пользователь вручную
    }
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  // Выпадающий список в "Система оплаты" — чтобы не разъезжались варианты написания
  // ("Оклад" vs "оклад"), от которых зависит фильтр в приложении.
  var paySystemRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["оклад", "сдельная"], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 5, 999, 1).setDataValidation(paySystemRule);
  sheet.autoResizeColumns(1, 5);
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
  sheet.getRange("A4").setValue("(наименование организации)");
  sheet.getRange("A5").setValue("Цех металообработки");
  sheet.getRange("A6").setValue("(структурное подразделение)");
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

  sheet.getRange("AI3").setValue("Утверждаю");
  sheet.getRange("AI4").setValue('Директор ТОО "ПК Алтын Таға"');
  sheet.getRange("AI5").setValue("______________Нұрмағамбет Ә. Е.");

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

  // Сотрудники: количество активных строк из "Сотрудники" определяет размер сетки —
  // если сотрудников больше 12 (или меньше), таблица подстраивается, а не ломается
  // (раньше FILTER на фиксированную сетку 18:29 падал в #REF!, если активных
  // сотрудников становилось больше 12 и спилл "наезжал" на строку ИТОГО).
  var firstEmployeeRow = 18;
  var employeeCount = Math.max(readEmployees(true).length, 1);
  var lastEmployeeRow = firstEmployeeRow + employeeCount - 1;
  var totalRow = lastEmployeeRow + 1;

  // A18 — спиллящийся список активных табельных номеров, отсортированный: сперва
  // "оклад", потом "сдельная" (столбец E «Система оплаты»; алфавитный порядок кириллицы
  // уже даёт нужную последовательность — "о" раньше "с"), внутри каждой группы — по
  // фамилии (столбец B «ФИО» начинается с фамилии). Пустая «Система оплаты» (не
  // назначена) — SORT уводит такие строки в конец независимо от направления сортировки.
  // Условие активности — TRIM(LOWER(...))="да", а не точное "=да": без этого лишний
  // пробел или другой регистр в столбце D ("Активен") тихо исключал сотрудника из
  // ЭТОЙ формулы, хотя readEmployees() в API уже делает trim/lowercase и видел бы его —
  // несовпадение ровно так теряло новых сотрудников при «Пересоздать форму» (живой
  // пример: "Активен" = " да" с пробелом у 30-го сотрудника).
  // B/C и ячейки дней ниже смотрят на свой же $A<row> по значению, а не по позиции —
  // поэтому сортировка (и добавление нового сотрудника в любое место списка) не путает
  // часы уже существующих сотрудников: каждая строка находит свой журнал по таб. №,
  // где бы она ни оказалась после сортировки.
  var activeCond = 'TRIM(LOWER(' + SHEET_EMPLOYEES + '!$D$2:$D$1000))="да"';
  var tabNumFilter = 'FILTER(' + SHEET_EMPLOYEES + '!$A$2:$A$1000,' + activeCond + ')';
  var paySystemFilter = 'FILTER(TRIM(LOWER(' + SHEET_EMPLOYEES + '!$E$2:$E$1000)),' + activeCond + ')';
  var nameFilter = 'FILTER(TRIM(' + SHEET_EMPLOYEES + '!$B$2:$B$1000),' + activeCond + ')';
  sheet.getRange("A" + firstEmployeeRow).setFormula(
    '=IFERROR(SORT(' + tabNumFilter + ',' + paySystemFilter + ',TRUE,' + nameFilter + ',TRUE),"")'
  );
  for (var row = firstEmployeeRow; row <= lastEmployeeRow; row++) {
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

  sheet.getRange("A" + totalRow).setValue("ИТОГО");
  sheet.getRange("AI" + totalRow).setFormula("=SUM(AI" + firstEmployeeRow + ":AI" + lastEmployeeRow + ")");
  sheet.getRange("AJ" + totalRow).setFormula("=SUM(AJ" + firstEmployeeRow + ":AJ" + lastEmployeeRow + ")");
  sheet.getRange("AL" + totalRow).setFormula("=SUM(AL" + firstEmployeeRow + ":AL" + lastEmployeeRow + ")");

  sheet.getRange("B" + (totalRow + 3)).setValue("Нач. цеха");
  sheet.getRange("C" + (totalRow + 3)).setValue("Утебалиев А.В.");
  sheet.getRange("B" + (totalRow + 6)).setValue("Гл. инженер");
  sheet.getRange("C" + (totalRow + 6)).setValue("Минбаев К.К.");

  applyTabelConditionalFormatting(sheet, firstEmployeeRow, lastEmployeeRow);
  copyArchiveFormatting(sheet, firstEmployeeRow, lastEmployeeRow, totalRow);

  // Скрываем строку 12 (технический разделитель в "Архив_исходный") и столбец
  // с технической датой AN — после копирования формата из архива, т.к. оно не
  // трогает видимость строк/столбцов, но порядок на всякий случай фиксируем явно.
  sheet.hideRows(12, 1);
  sheet.hideColumns(40, 1);

  // Повторный запуск "Пересоздать форму" не должен копить защиты поверх друг друга.
  var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var p = 0; p < existingProtections.length; p++) {
    existingProtections[p].remove();
  }
  var dataRange = sheet.getRange(firstEmployeeRow, 4, employeeCount, 31); // D..AH на сотрудников
  var protection = dataRange.protect().setDescription("Табель — печатная форма, собирается формулами из Журнала");
  protection.setWarningOnly(true);
}

/**
 * Копирует реальное форматирование (ширины столбцов, высоты строк, шрифты,
 * границы, выравнивание) из листа "Архив_исходный" — вместо того чтобы
 * подбирать значения на глаз.Шапка (строки 1-17) копируется как есть;
 * строка сотрудника из архива используется как шаблон и тиражируется на
 * текущее количество сотрудников; строка ИТОГО и подписи ниже неё копируются
 * со сдвигом на разницу в номере строки ИТОГО между архивом и текущим листом.
 */
function copyArchiveFormatting(sheet, firstEmployeeRow, lastEmployeeRow, totalRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archive = ss.getSheetByName(SHEET_ARCHIVE_SOURCE);
  if (!archive) return;

  var lastCol = 39; // A..AM — печатная форма; технический AN (40) не трогаем
  var archiveFirstEmployeeRow = 18;
  var archiveTotalRow = 30;
  var archiveLastRow = 37; // захватывает строку ИТОГО и подписи B33/B36 с запасом

  for (var c = 1; c <= lastCol; c++) {
    sheet.setColumnWidth(c, archive.getColumnWidth(c));
  }
  for (var r = 1; r < archiveFirstEmployeeRow; r++) {
    sheet.setRowHeight(r, archive.getRowHeight(r));
  }

  archive
    .getRange(1, 1, archiveFirstEmployeeRow - 1, lastCol)
    .copyTo(sheet.getRange(1, 1, archiveFirstEmployeeRow - 1, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  var employeeCount = lastEmployeeRow - firstEmployeeRow + 1;
  var employeeRowHeight = archive.getRowHeight(archiveFirstEmployeeRow);
  archive
    .getRange(archiveFirstEmployeeRow, 1, 1, lastCol)
    .copyTo(sheet.getRange(firstEmployeeRow, 1, employeeCount, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  for (var er = firstEmployeeRow; er <= lastEmployeeRow; er++) {
    sheet.setRowHeight(er, employeeRowHeight);
  }

  var tailRows = archiveLastRow - archiveTotalRow + 1;
  archive
    .getRange(archiveTotalRow, 1, tailRows, lastCol)
    .copyTo(sheet.getRange(totalRow, 1, tailRows, lastCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  for (var tr = 0; tr < tailRows; tr++) {
    sheet.setRowHeight(totalRow + tr, archive.getRowHeight(archiveTotalRow + tr));
  }
}

function applyTabelConditionalFormatting(sheet, firstEmployeeRow, lastEmployeeRow) {
  var range = sheet.getRange("D" + firstEmployeeRow + ":AH" + lastEmployeeRow);
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
  var employeeCount = Math.max(readEmployees(true).length, 1);
  var protection = copy.getRange(18, 4, employeeCount, 31).protect();
  protection.setWarningOnly(true);
  copy.hideSheet(); // архив месяца не должен мешаться в списке вкладок

  SpreadsheetApp.getUi().alert('Месяц сохранён в скрытый лист "' + name + '". Открыть его можно через "Вид → Скрытые листы".');
}

// ===================== чтение справочников =====================

function readEmployees(onlyActive) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLOYEES);
  var rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 5).getValues();
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var active = String(r[3]).trim().toLowerCase() === "да";
    if (onlyActive && !active) continue;
    // Систему оплаты назначает пользователь сам в таблице; пусто — пока не назначена
    // (в приложении такой сотрудник просто не попадает ни под один конкретный фильтр).
    var paySystem = r[4] ? String(r[4]).trim().toLowerCase() : "";
    result.push({ tabNum: r[0], name: r[1], position: r[2], active: active, paySystem: paySystem });
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
