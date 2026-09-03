/**
 * Вечерний отчёт (evening brief) — собирает цифры с вкладок мониторинга
 * в текстовый шаблон, готовый для копирования.
 *
 * КАК УСТРОЕН ПОИСК ЗНАЧЕНИЙ
 * На вкладках "Сводная", "Выводы", "Нагрузка L2", "Нагрузка L1",
 * "Нагрузка Fraud", "Зависшие тикеты" одинаковая структура: подписи
 * метрик — в столбце B, а числа по дням месяца — в столбцах C..AG
 * (C = 1-е число, D = 2-е, ..., AG = 31-е). Поэтому вместо жёстких
 * адресов вида "Сводная!C5" скрипт ищет нужную строку ПО ТЕКСТУ
 * ПОДПИСИ в столбце B, а нужный столбец — по числу месяца. Так отчёт
 * не сломается, если в таблице добавят/уберут строку выше нужной.
 *
 * Если для какого-то месяца лист переименован или структура правда
 * поменялась — поправьте константы в блоке SHEETS и NEEDLE ниже.
 */

var SHEETS = {
  SVODNAYA: 'Сводная',
  VYVODY: 'Выводы',
  L2: 'Нагрузка L2',
  L1: 'Нагрузка L1',
  FRAUD: 'Нагрузка Fraud',
  ZAVISSHIE: 'Зависшие тикеты'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Вечерний отчёт')
    .addItem('Сформировать отчёт…', 'openReportDialog')
    .addToUi();
}

function openReportDialog() {
  var html = HtmlService.createTemplateFromFile('Dialog')
    .evaluate()
    .setWidth(620)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Вечерний отчёт');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Дата "данных" по умолчанию — вчерашний день (ISO yyyy-MM-dd). */
function getDefaultDataDate() {
  var tz = Session.getScriptTimeZone();
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/** Сегодняшняя дата (ISO), для поля "Дата" в шапке отчёта. */
function getTodayDate() {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

// ---------- низкоуровневый поиск по подписям ----------

function normalize_(v) {
  return v === null || v === undefined ? '' : v.toString().replace(/\s+/g, ' ').trim();
}

/** Первая строка в столбце B, где текст ТОЧНО совпадает с label (после нормализации пробелов). */
function findRowExact_(sheet, label, startRow, endRow) {
  startRow = startRow || 1;
  endRow = endRow || sheet.getLastRow();
  var needle = normalize_(label);
  var values = sheet.getRange(startRow, 2, endRow - startRow + 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (normalize_(values[i][0]) === needle) return startRow + i;
  }
  return null;
}

/** Первая строка в столбце B, чей текст СОДЕРЖИТ needle (после нормализации). */
function findRowContains_(sheet, needle, startRow, endRow) {
  startRow = startRow || 1;
  endRow = endRow || sheet.getLastRow();
  var values = sheet.getRange(startRow, 2, endRow - startRow + 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (normalize_(values[i][0]).indexOf(needle) !== -1) return startRow + i;
  }
  return null;
}

/** Столбец для даты: C = 1-е число месяца, D = 2-е, ... AG = 31-е. */
function dateColumn_(date) {
  return date.getDate() + 2;
}

function toNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  var n = parseFloat(v);
  return isNaN(n) ? null : n;
}

/** Формат числа с пробелом как разделителем тысяч: 11820 -> "11 820". */
function formatNum_(v) {
  var n = toNumber_(v);
  if (n === null) return typeof v === 'string' ? v : '';
  var rounded = Math.round(n);
  var sign = rounded < 0 ? '-' : '';
  var s = Math.abs(rounded).toString();
  var out = '';
  while (s.length > 3) {
    out = ' ' + s.slice(-3) + out;
    s = s.slice(0, -3);
  }
  return sign + s + out;
}

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('Не найден лист "' + name + '"');
  return sheet;
}

/** Значение по точному тексту подписи в столбце B данного листа, для нужной даты. */
function valueByLabel_(sheetName, label, date, opts) {
  opts = opts || {};
  var sheet = getSheet_(sheetName);
  var row = opts.contains
    ? findRowContains_(sheet, label, opts.startRow, opts.endRow)
    : findRowExact_(sheet, label, opts.startRow, opts.endRow);
  if (!row) return '';
  return sheet.getRange(row, dateColumn_(date)).getValue();
}

/** Значение подписи label, но поиск ведём только в окне ниже anchorRow (чтобы не спутать
 *  одинаковые подписи "Mena 1x"/"Mena Leads 1x", повторяющиеся в разных блоках листа). */
function valueNearAnchor_(sheetName, anchorLabel, label, date, maxOffset) {
  var sheet = getSheet_(sheetName);
  var anchorRow = findRowExact_(sheet, anchorLabel);
  if (!anchorRow) return '';
  var row = findRowExact_(sheet, label, anchorRow + 1, anchorRow + (maxOffset || 20));
  if (!row) return '';
  return sheet.getRange(row, dateColumn_(date)).getValue();
}

// ---------- сбор всех полей отчёта ----------
//
// В строках вида "API (Team A/B): X / Y" из таблицы берётся только X
// (значение ДО "/"). Y — всегда заполняется вручную в диалоге, см.
// MANUAL_SECOND_VALUE_KEYS ниже и соответствующие поля в Dialog.html.

function collectReportValues_(date) {
  var v = {};

  v.created = valueByLabel_(SHEETS.SVODNAYA, 'Кол-во созданных тикетов', date);
  v.vyvody = valueByLabel_(SHEETS.VYVODY, 'уммарная нагрузка', date, { contains: true });

  // Нагрузка по процессам — берём только первое (авто) значение,
  // второе после "/" — ручной ввод (apiTeamB, pspSecond, ...).
  v.apiTeamA = valueByLabel_(SHEETS.L2, 'Team A', date);
  v.pspTotal = valueByLabel_(SHEETS.L2, 'уммарная нагрузка "PSP"', date, { contains: true });
  v.btMena1x = valueByLabel_(SHEETS.L2, 'Mena 1x', date);
  v.smpTotal = valueByLabel_(SHEETS.L2, 'уммарная нагрузка "SMP M"', date, { contains: true });
  v.l2l1Mena1x = valueNearAnchor_(SHEETS.L2, 'L2/L1 Mena 1x', 'Суммарное кол-во Депозиты', date, 3);

  // L1 / Fraud — тоже только первое значение.
  v.l1Mena1x = valueByLabel_(SHEETS.L1, 'Нагрузка Mena 1x', date);
  v.fraudMena1x = valueByLabel_(SHEETS.FRAUD, 'Нагрузка Mena 1x', date);

  // Зависшие (24+) — тоже только первое значение.
  v.zavPsp = valueByLabel_(SHEETS.ZAVISSHIE, 'PSP', date);
  v.zavBtMena1x = valueByLabel_(SHEETS.ZAVISSHIE, 'Mena 1x', date);
  v.zavSmp = valueByLabel_(SHEETS.ZAVISSHIE, 'SMP M', date);

  var btInProgressMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'PT 24 часа In Progress (M)', date, 10);
  var btInProgressMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'PT 24 часа In Progress (M)', date, 10);
  v.inProgressBt = sumValues_(btInProgressMena1x, btInProgressMenaLeads1x);

  // Эти две строки в шаблоне полностью автоматические (обе части MENA 1X / Leads 1X).
  v.btSentMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'BT Sent for processing (M) 72h+', date, 10);
  v.btSentMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'BT Sent for processing (M) 72h+', date, 10);
  v.btNewMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'New request (M) 72h+', date, 10);
  v.btNewMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'New request (M) 72h+', date, 10);

  v.smpSent = valueByLabel_(SHEETS.ZAVISSHIE, 'SMP Sent for processing 72h+', date);
  v.smpNew = valueByLabel_(SHEETS.ZAVISSHIE, 'New request 72h+', date);

  return v;
}

function sumValues_(a, b) {
  var na = toNumber_(a);
  var nb = toNumber_(b);
  if (na === null && nb === null) return '';
  return (na || 0) + (nb || 0);
}

// ---------- сборка итогового текста ----------

var TEMPLATE =
  'Дата: {{reportDate}}\n' +
  '\n' +
  'Общие показатели:\n' +
  'Создано тикетов (API/BT/PSP/SMP) за прошлый день: {{created}}\n' +
  '\n' +
  'Выводы (MENA 1X / Leads 1X): {{vyvody}}\n' +
  '\n' +
  'Нагрузка по процессам (MENA 1X / Leads 1X):\n' +
  'API (Team A/B):  {{apiTeamA}} / {{apiTeamB}}\n' +
  'PSP:  {{pspTotal}} / {{pspSecond}}\n' +
  'BT M: {{btMena1x}} / {{btMenaLeads1x}}\n' +
  'SMP M: {{smpTotal}} / {{smpSecond}}\n' +
  'L2/L1 (депозиты): {{l2l1Mena1x}} / {{l2l1MenaLeads1x}}\n' +
  '\n' +
  'L1:\n' +
  'Нагрузка: {{l1Mena1x}} / {{l1MenaLeads1x}}\n' +
  '\n' +
  'Fraud:\n' +
  'Нагрузка: {{fraudMena1x}} / {{fraudMenaLeads1x}}\n' +
  '\n' +
  'Зависшие (24+):\n' +
  'PSP/API  {{zavPsp}} / {{zavApi}}\n' +
  'BT M: {{zavBtMena1x}} / {{zavBtMenaLeads1x}}\n' +
  'SMP M: {{zavSmp}} / {{zavSmpSecond}}\n' +
  'In Progress (BT/SMP): {{inProgressBt}} / {{inProgressSmp}}\n' +
  '\n' +
  'BT Sent for processing (M) 72h+ MENA 1X / Leads 1X : {{btSentMena1x}}  / {{btSentMenaLeads1x}}\n' +
  'BT New request (M) 72h+  MENA 1X / Leads 1X: {{btNewMena1x}} / {{btNewMenaLeads1x}}\n' +
  'SMP Sent for processing (M) 72h+ : {{smpSent}}\n' +
  'SMP New Request 72h+: {{smpNew}}\n' +
  '\n' +
  'Чат 72: {{chatLink}}\n' +
  'Ожидают ответа: {{waiting}}\n' +
  '\n' +
  'Массовый Approved: {{approved}}\n' +
  'Ошибка - Статус: {{errorStatus}}\n' +
  '\n' +
  'Проблемные области:\n' +
  '{{problems}}';

// Значения ПОСЛЕ "/" в этих строках в таблице не найдены (или таблица
// в принципе не даёт для них второй разбивки) — их всегда вводят
// вручную в диалоге. Ключи совпадают с полями payload и плейсхолдерами
// в TEMPLATE.
var MANUAL_SECOND_VALUE_KEYS = [
  'apiTeamB', 'pspSecond', 'btMenaLeads1x', 'smpSecond',
  'l2l1MenaLeads1x', 'l1MenaLeads1x', 'fraudMenaLeads1x',
  'zavApi', 'zavBtMenaLeads1x', 'zavSmpSecond', 'inProgressSmp'
];

/**
 * Точка входа для диалога: payload = {
 *   reportDate: 'yyyy-MM-dd',  dataDate: 'yyyy-MM-dd',
 *   chatLink, waiting, approved, errorStatus, problems,
 *   ...MANUAL_SECOND_VALUE_KEYS
 * }
 */
function generateReport(payload) {
  var dataDate = parseIsoDate_(payload.dataDate);
  var raw = collectReportValues_(dataDate);

  var fields = {
    reportDate: formatDisplayDate_(parseIsoDate_(payload.reportDate)),
    chatLink: payload.chatLink || '',
    waiting: payload.waiting || '',
    approved: payload.approved || '',
    errorStatus: payload.errorStatus || '',
    problems: payload.problems || ''
  };
  for (var key in raw) {
    fields[key] = formatNum_(raw[key]);
  }
  MANUAL_SECOND_VALUE_KEYS.forEach(function (key) {
    fields[key] = formatNum_(payload[key] || '');
  });

  var text = TEMPLATE;
  for (var k in fields) {
    text = text.replace(new RegExp('{{' + k + '}}', 'g'), fields[k]);
  }
  return text;
}

function parseIsoDate_(iso) {
  var parts = iso.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function formatDisplayDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
}
