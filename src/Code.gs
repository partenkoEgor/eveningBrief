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
    .addItem('Поиск аномалий…', 'openSpikesDialog')
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

/** Сегодняшняя дата (ISO yyyy-MM-dd) — значение по умолчанию для единственного поля даты в диалоге. */
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
// Даты: "Создано тикетов" — это итог за ПРОШЛЫЙ (уже завершённый) день,
// поэтому берётся из колонки yesterdayDate. Все остальные автополя —
// текущий снимок очереди/нагрузки, берутся из колонки todayDate.
//
// В строках вида "API (Team A/B): X / Y" из таблицы берётся только X
// (значение ДО "/"). Y — всегда заполняется вручную в диалоге, см.
// MANUAL_SECOND_VALUE_KEYS ниже и соответствующие поля в Dialog.html.

// Человекочитаемые названия полей — для предупреждения о пустых ячейках.
var FIELD_LABELS = {
  created: 'Создано тикетов',
  vyvody: 'Выводы',
  apiTeamA: 'Нагрузка API (Суммарная нагрузка "API")',
  pspTotal: 'Нагрузка PSP',
  btTotal: 'Нагрузка BT M (Суммарная нагрузка "BT M")',
  smpTotal: 'Нагрузка SMP M',
  l2l1Total: 'L2/L1 депозиты (Mena 1x + Mena Leads 1x)',
  l1Total: 'L1 — Нагрузка (Cуммарное кол-во нагрузки)',
  fraudTotal: 'Fraud — Нагрузка (Cуммарное кол-во нагрузки)',
  zavPspApi: 'Зависшие PSP/API (сумма строк API + PSP)',
  zavBtM: 'Зависшие BT M (сумма PT 24 часа (ПК): Mena 1x + Mena Leads 1x)',
  zavSmp: 'Зависшие SMP M',
  inProgressBt: 'In Progress BT (Mena 1x + Mena Leads 1x + SMP)',
  btSentTotal: 'BT Sent for processing 72h+ (Mena 1x + Mena Leads 1x)',
  btNewTotal: 'BT New request 72h+ (Mena 1x + Mena Leads 1x)',
  smpSent: 'SMP Sent for processing 72h+',
  smpNew: 'SMP New request 72h+'
};

function collectReportValues_(todayDate, yesterdayDate) {
  var v = {};

  v.created = valueByLabel_(SHEETS.SVODNAYA, 'Кол-во созданных тикетов', yesterdayDate);
  v.vyvody = valueByLabel_(SHEETS.VYVODY, 'уммарная нагрузка', todayDate, { contains: true });

  // Нагрузка по процессам — берём только первое (авто) значение,
  // второе после "/" — ручной ввод (apiTeamB, pspSecond, ...).
  v.apiTeamA = valueByLabel_(SHEETS.L2, 'уммарная нагрузка "API"', todayDate, { contains: true });
  v.pspTotal = valueByLabel_(SHEETS.L2, 'уммарная нагрузка "PSP"', todayDate, { contains: true });
  v.btTotal = valueByLabel_(SHEETS.L2, 'уммарная нагрузка "BT M"', todayDate, { contains: true });
  v.smpTotal = valueByLabel_(SHEETS.L2, 'уммарная нагрузка "SMP M"', todayDate, { contains: true });
  var l2l1Mena1xDeposits = valueNearAnchor_(SHEETS.L2, 'L2/L1 Mena 1x', 'Суммарное кол-во Депозиты', todayDate, 3);
  var l2l1MenaLeads1xDeposits = valueNearAnchor_(SHEETS.L2, 'L2/L1 Mena Leads 1x', 'Суммарное кол-во Депозиты', todayDate, 3);
  v.l2l1Total = sumValues_(l2l1Mena1xDeposits, l2l1MenaLeads1xDeposits);

  // L1 / Fraud — тоже только первое значение.
  v.l1Total = valueByLabel_(SHEETS.L1, 'уммарное кол-во нагрузки', todayDate, { contains: true });
  v.fraudTotal = valueByLabel_(SHEETS.FRAUD, 'уммарное кол-во нагрузки', todayDate, { contains: true });

  // Зависшие (24+) — тоже только первое значение.
  var zavApiTotal = valueByLabel_(SHEETS.ZAVISSHIE, 'API', todayDate);
  var zavPspTotal = valueByLabel_(SHEETS.ZAVISSHIE, 'PSP', todayDate);
  v.zavPspApi = sumValues_(zavApiTotal, zavPspTotal);
  var zavBtM24hMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'PT 24 часа (ПК)', todayDate, 10);
  var zavBtM24hMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'PT 24 часа (ПК)', todayDate, 10);
  v.zavBtM = sumValues_(zavBtM24hMena1x, zavBtM24hMenaLeads1x);
  v.zavSmp = valueByLabel_(SHEETS.ZAVISSHIE, 'SMP M', todayDate);

  var btInProgressMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'PT 24 часа In Progress (M)', todayDate, 10);
  var btInProgressMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'PT 24 часа In Progress (M)', todayDate, 10);
  var smpInProgress24h = valueByLabel_(SHEETS.ZAVISSHIE, 'PT 24 часа In Progress', todayDate);
  v.inProgressBt = sumValues_(btInProgressMena1x, btInProgressMenaLeads1x, smpInProgress24h);

  // Значение до "/" — сумма Mena 1x + Mena Leads 1x из таблицы; после "/" — ручной ввод.
  var btSentMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'BT Sent for processing (M) 72h+', todayDate, 10);
  var btSentMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'BT Sent for processing (M) 72h+', todayDate, 10);
  v.btSentTotal = sumValues_(btSentMena1x, btSentMenaLeads1x);
  var btNewMena1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena 1x', 'New request (M) 72h+', todayDate, 10);
  var btNewMenaLeads1x = valueNearAnchor_(SHEETS.ZAVISSHIE, 'Mena Leads 1x', 'New request (M) 72h+', todayDate, 10);
  v.btNewTotal = sumValues_(btNewMena1x, btNewMenaLeads1x);

  v.smpSent = valueByLabel_(SHEETS.ZAVISSHIE, 'SMP Sent for processing 72h+', todayDate);
  v.smpNew = valueByLabel_(SHEETS.ZAVISSHIE, 'New request 72h+', todayDate);

  return v;
}

/** Сумма любого числа значений; пустые/нечисловые считаются нулём.
 *  Если ВСЕ значения пустые — возвращает '' (чтобы поле числилось как "не найдено"). */
function sumValues_() {
  var total = 0;
  var allEmpty = true;
  for (var i = 0; i < arguments.length; i++) {
    var n = toNumber_(arguments[i]);
    if (n !== null) {
      allEmpty = false;
      total += n;
    }
  }
  return allEmpty ? '' : total;
}

// ---------- сборка итогового текста ----------

var TEMPLATE =
  '#отчёт_monitoring\n' +
  '\n' +
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
  'BT M: {{btTotal}} / {{btMenaLeads1x}}\n' +
  'SMP M: {{smpTotal}} / {{smpSecond}}\n' +
  'L2/L1 (депозиты): {{l2l1Total}} / {{l2l1MenaLeads1x}}\n' +
  '\n' +
  'L1:\n' +
  'Нагрузка: {{l1Total}} / {{l1MenaLeads1x}}\n' +
  '\n' +
  'Fraud:\n' +
  'Нагрузка: {{fraudTotal}} / {{fraudMenaLeads1x}}\n' +
  '\n' +
  'Зависшие (24+):\n' +
  'PSP/API  {{zavPspApi}} / {{zavApi}}\n' +
  'BT M: {{zavBtM}} / {{zavBtMenaLeads1x}}\n' +
  'SMP M: {{zavSmp}} / {{zavSmpSecond}}\n' +
  'In Progress (BT/SMP): {{inProgressBt}} / {{inProgressSmp}}\n' +
  '\n' +
  'BT Sent for processing (M) 72h+ MENA 1X / Leads 1X : {{btSentTotal}}  / {{btSentSecond}}\n' +
  'BT New request (M) 72h+  MENA 1X / Leads 1X: {{btNewTotal}} / {{btNewSecond}}\n' +
  'SMP Sent for processing (M) 72h+ : {{smpSent}} / {{smpSentSecond}}\n' +
  'SMP New Request 72h+: {{smpNew}} / {{smpNewSecond}}\n' +
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
  'zavApi', 'zavBtMenaLeads1x', 'zavSmpSecond', 'inProgressSmp',
  'btSentSecond', 'btNewSecond', 'smpSentSecond', 'smpNewSecond'
];

/**
 * Точка входа для диалога: payload = {
 *   date: 'yyyy-MM-dd' (это "сегодня" для отчёта — от него же считается "вчера"),
 *   chatLink, waiting, approved, errorStatus, problems,
 *   ...MANUAL_SECOND_VALUE_KEYS
 * }
 * Возвращает { text, missing } — missing перечисляет автополя, для
 * которых в таблице не нашлась строка или ячейка на нужную дату
 * оказалась пустой (в тексте на их месте останется пустое место).
 */
function generateReport(payload) {
  var todayDate = parseIsoDate_(payload.date);
  var yesterdayDate = new Date(todayDate);
  yesterdayDate.setDate(yesterdayDate.getDate() - 1);

  var raw = collectReportValues_(todayDate, yesterdayDate);

  var missing = [];
  for (var rk in raw) {
    if (raw[rk] === '' && FIELD_LABELS[rk]) missing.push(FIELD_LABELS[rk]);
  }

  var fields = {
    reportDate: formatDisplayDate_(todayDate),
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
  return { text: text, missing: missing };
}

function parseIsoDate_(iso) {
  var parts = iso.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function formatDisplayDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
}

// ========== ПОИСК АНОМАЛИЙ (ВСПЛЕСКОВ) ==========
//
// Ищем процентные метрики, которые в выбранный день заметно выбились из
// собственной нормы за предыдущие дни месяца.
//
// Номера строк НЕ прописаны жёстко: книга заводится заново каждый месяц,
// и вставка одного блока сдвинула бы все номера. Инвариант — числовой
// формат ячейки (0.00% / 0.0% / 0%), по нему строки и находятся.
//
// Столбец дня тоже не хардкодим: в строке 2 у всех дневных листов лежит
// число вида "день + месяц/100" (1.08, 2.08, ... 31.08). Поиск по этой
// шапке разом покрывает все раскладки: C..AG, D..AH, E..AI и парные
// столбцы на листе "Трафик + Конверсия" (трафик + конверсия на день).

var SPIKE_PRESETS = {
  low: { zMin: 5.0, minAbsDelta: 0.15, minAbsLevel: 0.10 },
  medium: { zMin: 3.5, minAbsDelta: 0.10, minAbsLevel: 0.05 },
  high: { zMin: 2.5, minAbsDelta: 0.05, minAbsLevel: 0.03 }
};

var SPIKE_MIN_HISTORY = 5;       // меньше точек — по строке судить рано
var SPIKE_MIN_DAY_COLUMNS = 5;   // меньше столбцов-дней — лист не дневной
var SPIKE_PERCENT_ROW_RATIO = 0.6;
var SPIKE_COLOR = '#f4c7c3';
var SPIKE_MAX_CELLS = 400;
var HIGHLIGHT_PROP = 'spikeHighlight';

function openSpikesDialog() {
  var html = HtmlService.createTemplateFromFile('SpikesDialog')
    .evaluate()
    .setWidth(720)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Поиск аномалий');
}

// ---------- геометрия листа ----------

/** Столбцы дней по шапке (строка 2). Возвращает null, если дневной оси нет —
 *  так из скана сами собой выпадают недельные и сводные листы. */
function resolveDayColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 2 || sheet.getLastRow() < 3) return null;

  var header = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var marks = [];
  for (var i = 0; i < header.length; i++) {
    var v = header[i];
    if (typeof v !== 'number' || !isFinite(v) || v < 1 || v >= 32) continue;
    var day = Math.floor(v);
    marks.push({ col: i + 1, day: day, month: Math.round((v - day) * 100) });
  }
  if (marks.length < SPIKE_MIN_DAY_COLUMNS) return null;

  var byDay = {};
  var months = {};
  for (var m = 0; m < marks.length; m++) {
    var mark = marks[m];
    var nextCol = (m + 1 < marks.length) ? marks[m + 1].col : mark.col + 1;
    // Ширина дня — до следующего заголовка, но не больше двух столбцов:
    // на "Трафик + Конверсия" это как раз пара "трафик / конверсия".
    var span = Math.max(1, Math.min(nextCol - mark.col, 2));
    var cols = byDay[mark.day] || [];
    for (var s = 0; s < span; s++) cols.push(mark.col + s);
    byDay[mark.day] = cols;
    if (mark.month) months[mark.month] = (months[mark.month] || 0) + 1;
  }
  return { byDay: byDay, firstDayCol: marks[0].col, months: months };
}

/** Месяц листа — самый часто встречающийся в шапке. */
function modalMonth_(months) {
  var best = null;
  var bestCount = 0;
  for (var m in months) {
    if (months[m] > bestCount) {
      bestCount = months[m];
      best = Number(m);
    }
  }
  return best;
}

/** Для одной строки выбирает в каждом дне ту колонку, что отформатирована
 *  как проценты. На парных столбцах это отсекает колонку трафика. */
function percentColsForRow_(formatRow, valueRow, byDay) {
  var colByDay = {};
  var pctDays = 0;
  var totalDays = 0;
  var numeric = 0;

  for (var day in byDay) {
    totalDays++;
    var cols = byDay[day];
    var chosen = null;
    for (var i = 0; i < cols.length; i++) {
      var fmt = formatRow[cols[i] - 1];
      if (fmt && fmt.toString().indexOf('%') !== -1) {
        chosen = cols[i];
        break;
      }
    }
    if (chosen === null) continue;
    pctDays++;
    colByDay[day] = chosen;
    if (toNumber_(valueRow[chosen - 1]) !== null) numeric++;
  }
  return { colByDay: colByDay, pctDays: pctDays, totalDays: totalDays, numeric: numeric };
}

/** Строка процентная, если большинство дневных ячеек в процентном формате
 *  и в ней есть живые числа. Числовая проверка отсекает строки со
 *  значками ▲/▼/○ — они текстовые. */
function isPercentRow_(info) {
  if (info.totalDays === 0) return false;
  if (info.numeric < SPIKE_MIN_HISTORY) return false;
  return info.pctDays >= info.totalDays * SPIKE_PERCENT_ROW_RATIO;
}

/** Подпись строки — первая непустая ячейка слева от первого дня.
 *  На "ЗТ по ГЕО API" дни начинаются с E, поэтому сюда попадают подписи
 *  из столбца D, а не только из B. */
function rowLabel_(valueRow, firstDayCol) {
  for (var col = firstDayCol - 1; col >= 1; col--) {
    var v = normalize_(valueRow[col - 1]);
    if (v) return v;
  }
  return '';
}

/** Тип метрики определяем по тексту подписи, а не по номеру строки.
 *  'зменени' покрывает и "Изменение", и "изменения", и "Изменения". */
function metricKind_(sheetName, label) {
  if (sheetName.indexOf('Трафик') !== -1) return 'conv';
  if (label.indexOf('От нагрузки') !== -1) return 'load';
  return 'delta';
}

// ---------- статистика ----------

function median_(values) {
  if (!values.length) return null;
  var arr = values.slice().sort(function (a, b) { return a - b; });
  var mid = Math.floor(arr.length / 2);
  return arr.length % 2 ? arr[mid] : (arr[mid - 1] + arr[mid]) / 2;
}

/** Медиана и MAD — в отличие от среднего и σ, не разъезжаются от одного
 *  выброса в истории (а выброс мы как раз и ищем). */
function robustStats_(series) {
  var med = median_(series);
  var dev = [];
  for (var i = 0; i < series.length; i++) dev.push(Math.abs(series[i] - med));
  var mad = median_(dev);
  return { median: med, mad: mad, sigma: 1.4826 * mad };
}

/** Всплеск = отклонение больше минимального порога И больше типичного
 *  разброса этой же строки. Возвращает null, если это не всплеск.
 *
 *  sigma === 0 значит "история строки константна" (частое на тихих
 *  строках, простоявших весь месяц в нуле). z тогда не определён, и мы
 *  считаем условие выполненным: иначе прыжок с 0% до 40% остался бы
 *  незамеченным — а это худшее, что здесь можно пропустить. */
function evaluateSpike_(x, stats, kind, preset) {
  var diff = x - stats.median;
  var absDiff = Math.abs(diff);
  var minAbs = (kind === 'delta') ? preset.minAbsDelta : preset.minAbsLevel;
  if (absDiff < minAbs) return null;

  var z = null;
  if (stats.sigma > 0) {
    z = absDiff / stats.sigma;
    if (z < preset.zMin) return null;
  }

  return {
    z: z,
    diff: diff,
    direction: diff > 0 ? 'рост' : 'падение',
    // bad влияет только на иконку и порядок сортировки, находку не скрывает
    bad: kind === 'conv' ? diff < 0 : (kind === 'load' ? diff > 0 : true)
  };
}

// ---------- скан ----------

function scanSheetForSpikes_(sheet, targetDate, preset, out) {
  var name = sheet.getName();
  var geom = resolveDayColumns_(sheet);
  if (!geom) return;

  var day = targetDate.getDate();
  var month = modalMonth_(geom.months);
  if (month && month !== targetDate.getMonth() + 1) {
    out.warnings.push('«' + name + '» — данные за другой месяц, лист пропущен');
    return;
  }
  if (!geom.byDay[day]) {
    out.warnings.push('«' + name + '» — нет столбца для этого дня, лист пропущен');
    return;
  }

  out.stats.sheetsScanned++;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var values = range.getValues();
  var formats = range.getNumberFormats();

  for (var r = 3; r <= lastRow; r++) {
    var valueRow = values[r - 1];
    var info = percentColsForRow_(formats[r - 1], valueRow, geom.byDay);
    if (!isPercentRow_(info)) continue;

    out.stats.rowsChecked++;

    var todayCol = info.colByDay[day];
    if (!todayCol) continue;
    var x = toNumber_(valueRow[todayCol - 1]);
    if (x === null) continue;

    // История: те же дни этой же строки. Пустые ячейки просто не попадают
    // в выборку — это верно и для дыр в середине строки, и для стран,
    // подключённых в середине месяца. Пропуск значит "не наблюдалось".
    var series = [];
    for (var d = 1; d < day; d++) {
      var col = info.colByDay[d];
      if (!col) continue;
      var n = toNumber_(valueRow[col - 1]);
      if (n !== null) series.push(n);
    }
    if (series.length < SPIKE_MIN_HISTORY) {
      out.stats.rowsSkipped++;
      continue;
    }

    var label = rowLabel_(valueRow, geom.firstDayCol);
    var kind = metricKind_(name, label);
    var stats = robustStats_(series);
    var hit = evaluateSpike_(x, stats, kind, preset);
    if (!hit) continue;

    out.spikes.push({
      sheet: name,
      row: r,
      col: todayCol,
      label: label,
      kind: kind,
      value: x,
      median: stats.median,
      z: hit.z,
      diff: hit.diff,
      direction: hit.direction,
      bad: hit.bad,
      day: day
    });
  }
}

/**
 * Точка входа для диалога. payload = { date: 'yyyy-MM-dd', sensitivity }.
 * Ничего в таблице не меняет — только читает.
 */
function findSpikes(payload) {
  var targetDate = parseIsoDate_(payload.date);
  var preset = SPIKE_PRESETS[payload.sensitivity] || SPIKE_PRESETS.medium;

  var out = {
    date: formatDisplayDate_(targetDate),
    spikes: [],
    warnings: [],
    stats: { sheetsScanned: 0, rowsChecked: 0, rowsSkipped: 0 }
  };

  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    scanSheetForSpikes_(sheets[i], targetDate, preset, out);
  }

  // Сильнее всего выбившиеся — наверх; константная история (z = ∞) первой.
  out.spikes.sort(function (a, b) {
    var az = a.z === null ? Infinity : a.z;
    var bz = b.z === null ? Infinity : b.z;
    return bz - az;
  });

  return out;
}

// ---------- подсветка ----------

function columnLetter_(col) {
  var s = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

function formatPercent_(fraction) {
  var n = toNumber_(fraction);
  return n === null ? '' : (n * 100).toFixed(1) + ' %';
}

function sameColor_(a, b) {
  return String(a).toLowerCase() === String(b).toLowerCase();
}

function saveHighlightRecord_(record) {
  PropertiesService.getDocumentProperties()
    .setProperty(HIGHLIGHT_PROP, JSON.stringify(record));
}

function loadHighlightRecord_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(HIGHLIGHT_PROP);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function clearHighlightRecord_() {
  PropertiesService.getDocumentProperties().deleteProperty(HIGHLIGHT_PROP);
}

/** Снимает подсветку и возвращает число восстановленных ячеек.
 *  Ячейку, перекрашенную человеком вручную, не трогает. */
function clearHighlightInternal_() {
  var record = loadHighlightRecord_();
  if (!record || !record.sheets) return 0;

  var ss = SpreadsheetApp.getActive();
  var restored = 0;

  for (var i = 0; i < record.sheets.length; i++) {
    var entry = record.sheets[i];
    var sheet = ss.getSheetByName(entry.s);
    if (!sheet) continue;

    var minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
    for (var j = 0; j < entry.c.length; j++) {
      minRow = Math.min(minRow, entry.c[j][0]);
      maxRow = Math.max(maxRow, entry.c[j][0]);
      minCol = Math.min(minCol, entry.c[j][1]);
      maxCol = Math.max(maxCol, entry.c[j][1]);
    }
    var backgrounds = sheet
      .getRange(minRow, minCol, maxRow - minRow + 1, maxCol - minCol + 1)
      .getBackgrounds();

    var groups = {};
    for (var k = 0; k < entry.c.length; k++) {
      var row = entry.c[k][0];
      var col = entry.c[k][1];
      var current = backgrounds[row - minRow][col - minCol];
      if (!sameColor_(current, SPIKE_COLOR)) continue;
      var prev = entry.p[entry.c[k][2]];
      if (!groups[prev]) groups[prev] = [];
      groups[prev].push(columnLetter_(col) + row);
      restored++;
    }
    for (var color in groups) {
      sheet.getRangeList(groups[color]).setBackground(color);
    }
  }

  clearHighlightRecord_();
  return restored;
}

/**
 * Точка входа для диалога: красит ячейки найденных всплесков.
 * МЕНЯЕТ ОБЩУЮ ТАБЛИЦУ — вызывается только по явному нажатию кнопки.
 */
function highlightSpikes(cells, dateLabel) {
  if (!cells || !cells.length) {
    return { ok: false, message: 'Нечего подсвечивать' };
  }
  if (cells.length > SPIKE_MAX_CELLS) {
    return {
      ok: false,
      message: 'Слишком много ячеек (' + cells.length + '). Понизьте чувствительность.'
    };
  }

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, message: 'Кто-то уже меняет подсветку, попробуйте через несколько секунд' };
  }

  try {
    // Старую подсветку снимаем всегда — так повторный запуск, смена даты
    // и смена чувствительности идемпотентны, и живой набор всегда один.
    clearHighlightInternal_();

    var ss = SpreadsheetApp.getActive();
    var bySheet = {};
    for (var i = 0; i < cells.length; i++) {
      var name = cells[i].sheet;
      if (!bySheet[name]) bySheet[name] = [];
      bySheet[name].push(cells[i]);
    }

    var record = { date: dateLabel || '', sheets: [] };

    for (var sheetName in bySheet) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      var list = bySheet[sheetName];

      var minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;
      for (var j = 0; j < list.length; j++) {
        minRow = Math.min(minRow, list[j].row);
        maxRow = Math.max(maxRow, list[j].row);
        minCol = Math.min(minCol, list[j].col);
        maxCol = Math.max(maxCol, list[j].col);
      }
      var backgrounds = sheet
        .getRange(minRow, minCol, maxRow - minRow + 1, maxCol - minCol + 1)
        .getBackgrounds();

      var palette = [];
      var entries = [];
      var a1 = [];
      for (var k = 0; k < list.length; k++) {
        var color = backgrounds[list[k].row - minRow][list[k].col - minCol];
        var pi = palette.indexOf(color);
        if (pi === -1) {
          palette.push(color);
          pi = palette.length - 1;
        }
        entries.push([list[k].row, list[k].col, pi]);
        a1.push(columnLetter_(list[k].col) + list[k].row);
      }

      sheet.getRangeList(a1).setBackground(SPIKE_COLOR);
      record.sheets.push({ s: sheetName, p: palette, c: entries });
    }

    var json = JSON.stringify(record);
    if (json.length > 8500) {
      return { ok: false, message: 'Слишком много ячеек для запоминания. Понизьте чувствительность.' };
    }
    saveHighlightRecord_(record);
    SpreadsheetApp.flush();
    return { ok: true, painted: cells.length };
  } finally {
    lock.releaseLock();
  }
}

/** Точка входа для диалога: снимает подсветку, возвращая прежние цвета. */
function clearSpikeHighlight() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, message: 'Кто-то уже меняет подсветку, попробуйте через несколько секунд' };
  }
  try {
    var record = loadHighlightRecord_();
    if (!record) {
      return { ok: true, restored: 0, message: 'Подсветка не найдена — нечего убирать' };
    }
    var restored = clearHighlightInternal_();
    SpreadsheetApp.flush();
    var from = record.date ? ' от ' + record.date : '';
    return { ok: true, restored: restored, message: 'Убрана подсветка' + from + ', ячеек: ' + restored };
  } finally {
    lock.releaseLock();
  }
}

/** Служебная: печатает найденный набор процентных строк в лог.
 *  Запускать при смене месяца, чтобы сверить, что автопоиск не поехал. */
function debugListPercentRows() {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var geom = resolveDayColumns_(sheet);
    if (!geom) {
      Logger.log('— %s: дневной оси нет, пропущен', sheet.getName());
      continue;
    }
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    var values = range.getValues();
    var formats = range.getNumberFormats();
    var found = [];
    for (var r = 3; r <= lastRow; r++) {
      var info = percentColsForRow_(formats[r - 1], values[r - 1], geom.byDay);
      if (!isPercentRow_(info)) continue;
      found.push(r + ' «' + rowLabel_(values[r - 1], geom.firstDayCol) + '»');
    }
    Logger.log('%s (%s строк): %s', sheet.getName(), found.length, found.join('; '));
  }
}
