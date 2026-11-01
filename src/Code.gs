/**
 * Вечерний отчёт — собирает значения с вкладок мониторинга
 * и формирует готовый текст отчёта.
 *
 * Выводы: каждая строка с выбранной датой фиксации в месячных листах
 * таблицы «Выводы L2 PS | 2026» считается отдельно, включая дубли.
 * К полученному количеству прибавляется ручное поле «Выводы L1-L2».
 * Для установки замените основной код Apps Script целиком этим файлом
 * и вместе с ним обновите файл Dialog: дата отчёта теперь считается в
 * диалоге, а не на сервере. При первом запуске может потребоваться
 * разрешение на чтение другой Google Таблицы; нужны права доступа к ней.
 * Для отмены верните предыдущий основной код.
 *
 * Логика дат:
 * - все текущие показатели берутся за выбранную дату;
 * - «Создано тикетов» берётся за предыдущий день только внутри текущего месяца;
 * - если выбрано 1-е число месяца, в предыдущую месячную таблицу скрипт не обращается,
 *   значение «Создано тикетов» остаётся пустым и выводится предупреждение.
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
    .setWidth(680)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, 'Вечерний отчёт');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- низкоуровневый поиск по подписям ----------

function normalize_(v) {
  return v === null || v === undefined
    ? ''
    : v.toString().replace(/\s+/g, ' ').trim();
}

/** Первая строка в столбце B, где текст точно совпадает с label. */
function findRowExact_(sheet, label, startRow, endRow) {
  startRow = startRow || 1;
  endRow = endRow || sheet.getLastRow();

  startRow = Math.max(1, startRow);
  endRow = Math.min(sheet.getLastRow(), endRow);

  if (endRow < startRow) return null;

  var needle = normalize_(label);
  var values = sheet.getRange(startRow, 2, endRow - startRow + 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    if (normalize_(values[i][0]) === needle) return startRow + i;
  }

  return null;
}

/** Первая строка в столбце B, чей текст содержит needle. */
function findRowContains_(sheet, needle, startRow, endRow) {
  startRow = startRow || 1;
  endRow = endRow || sheet.getLastRow();

  startRow = Math.max(1, startRow);
  endRow = Math.min(sheet.getLastRow(), endRow);

  if (endRow < startRow) return null;

  needle = normalize_(needle);
  var values = sheet.getRange(startRow, 2, endRow - startRow + 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    if (normalize_(values[i][0]).indexOf(needle) !== -1) return startRow + i;
  }

  return null;
}

/** Столбец для даты: C = 1-е число, D = 2-е, ... AG = 31-е. */
function dateColumn_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error('Некорректная дата отчёта.');
  }

  var day = date.getDate();
  if (day < 1 || day > 31) {
    throw new Error('Некорректный день месяца: ' + day);
  }

  return day + 2;
}

function toNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;

  var cleaned = v.toString()
    .replace(/\s+/g, '')
    .replace(',', '.');

  if (cleaned === '') return null;

  var n = Number(cleaned);
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

/** Значение по подписи в столбце B данного листа для нужной даты. */
function valueByLabel_(sheetName, label, date, opts) {
  opts = opts || {};

  var sheet = getSheet_(sheetName);
  var row = opts.contains
    ? findRowContains_(sheet, label, opts.startRow, opts.endRow)
    : findRowExact_(sheet, label, opts.startRow, opts.endRow);

  if (!row) return '';

  return sheet.getRange(row, dateColumn_(date)).getValue();
}

/**
 * Значение подписи label в ограниченном окне ниже anchorLabel.
 * Используется для повторяющихся блоков Mena 1x / Mena Leads 1x.
 */
function valueNearAnchor_(sheetName, anchorLabel, label, date, maxOffset) {
  var sheet = getSheet_(sheetName);
  var anchorRow = findRowExact_(sheet, anchorLabel);

  if (!anchorRow) return '';

  var row = findRowExact_(
    sheet,
    label,
    anchorRow + 1,
    anchorRow + (maxOffset || 20)
  );

  if (!row) return '';

  return sheet.getRange(row, dateColumn_(date)).getValue();
}

// Источник записей по выводам. Из него читаются только даты фиксации.
var WITHDRAWALS_SOURCE = {
  spreadsheetId: '1oGiFgCzN1iOia-AmfRD6H9wEo0-zbxXZbTAKCOHs5vk',
  monthSheets: [
    'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
    'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'
  ]
};

/**
 * Считает строки во всех месячных листах источника за дату отчёта.
 * Каждая строка с этой датой фиксации даёт +1, включая дубли.
 * Даты создания вывода, статусы, суммы и уникальность ID не используются.
 * Исходную таблицу не изменяет.
 */
function countWithdrawalsByDate_(reportDate) {
  var source;
  try {
    source = SpreadsheetApp.openById(WITHDRAWALS_SOURCE.spreadsheetId);
  } catch (error) {
    throw new Error('Не удалось открыть таблицу «Выводы L2 PS | 2026». Проверьте доступ и разрешение скрипта на чтение этой таблицы.');
  }

  var targetKey = withdrawalsCalendarKey_(
    reportDate.getFullYear(), reportDate.getMonth() + 1, reportDate.getDate()
  );
  var timezone = source.getSpreadsheetTimeZone();
  var locale = source.getSpreadsheetLocale();
  var sheets = source.getSheets();
  var foundSheets = 0;
  var total = 0;

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = normalize_(sheet.getName()).toLowerCase();
    if (WITHDRAWALS_SOURCE.monthSheets.indexOf(sheetName) === -1) continue;
    foundSheets++;

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow < 3 || lastColumn < 1) continue;

    var headerRows = sheet.getRange(1, 1, Math.min(10, lastRow), lastColumn).getValues();
    var headerRow = null;
    var dateColumn = null;
    for (var r = 0; r < headerRows.length; r++) {
      for (var c = 0; c < headerRows[r].length; c++) {
        if (normalize_(headerRows[r][c]).toLowerCase() !== 'дата фиксации') continue;
        if (dateColumn !== null) {
          throw new Error('На листе «' + sheet.getName() + '» несколько заголовков «Дата фиксации».');
        }
        headerRow = r + 1;
        dateColumn = c + 1;
      }
    }
    if (dateColumn === null) {
      throw new Error('На листе «' + sheet.getName() + '» не найден столбец «Дата фиксации».');
    }
    if (lastRow <= headerRow) continue;

    var dates = sheet.getRange(headerRow + 1, dateColumn, lastRow - headerRow, 1).getValues();
    for (var j = 0; j < dates.length; j++) {
      var value = dates[j][0];
      if (value === '' || value === null || value === undefined) continue;
      var key = withdrawalsDateKey_(value, timezone, locale);
      if (key === null) {
        throw new Error('На листе «' + sheet.getName() + '», строка ' + (headerRow + 1 + j) + ': некорректная «Дата фиксации». Исправьте дату и повторите формирование отчёта.');
      }
      if (key === targetKey) total++;
    }
  }

  if (foundSheets === 0) {
    throw new Error('В таблице выводов не найдены месячные листы: Январь, Февраль и т. д.');
  }
  return total;
}

/** Ключ календарной даты без времени. Некорректная дата возвращает null. */
function withdrawalsCalendarKey_(year, month, day) {
  var check = new Date(Date.UTC(year, month - 1, day));
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null;
  return String(year) + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2);
}

/** Даты Google Таблиц сравниваются в часовом поясе таблицы-источника. */
function withdrawalsDateKey_(value, timezone, locale) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }
  if (typeof value === 'number') {
    if (!isFinite(value) || value < 1 || value > 2958465) return null;
    var serialDate = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return withdrawalsCalendarKey_(serialDate.getUTCFullYear(), serialDate.getUTCMonth() + 1, serialDate.getUTCDate());
  }
  if (typeof value !== 'string') return null;
  var text = value.trim();
  if (text === '') return '';
  var parts = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T,].*)?$/);
  if (parts) return withdrawalsCalendarKey_(Number(parts[1]), Number(parts[2]), Number(parts[3]));
  parts = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T,].*)?$/);
  if (parts) return withdrawalsCalendarKey_(Number(parts[3]), Number(parts[2]), Number(parts[1]));
  parts = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T,].*)?$/);
  if (parts) {
    var monthFirst = /^en[_-]US$/i.test(locale || '');
    return withdrawalsCalendarKey_(Number(parts[3]), Number(parts[monthFirst ? 1 : 2]), Number(parts[monthFirst ? 2 : 1]));
  }
  return null;
}

// ---------- сбор всех автоматических полей отчёта ----------

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

function collectReportValues_(todayDate) {
  var v = {};

  // «Создано тикетов» берём за вчера только если вчера находится
  // в этой же месячной таблице. На 1-е число месяца не ищем прошлый месяц.
  if (todayDate.getDate() === 1) {
    v.created = '';
  } else {
    var yesterdayDate = new Date(todayDate);
    yesterdayDate.setDate(yesterdayDate.getDate() - 1);
    v.created = valueByLabel_(
      SHEETS.SVODNAYA,
      'Кол-во созданных тикетов',
      yesterdayDate
    );
  }

  v.vyvody = countWithdrawalsByDate_(todayDate);

  // Нагрузка по процессам. Первое значение подтягивает скрипт,
  // вечернее значение вводится в диалоге.
  v.apiTeamA = valueByLabel_(
    SHEETS.L2,
    'уммарная нагрузка "API"',
    todayDate,
    { contains: true }
  );

  v.pspTotal = valueByLabel_(
    SHEETS.L2,
    'уммарная нагрузка "PSP"',
    todayDate,
    { contains: true }
  );

  v.btTotal = valueByLabel_(
    SHEETS.L2,
    'уммарная нагрузка "BT M"',
    todayDate,
    { contains: true }
  );

  v.smpTotal = valueByLabel_(
    SHEETS.L2,
    'уммарная нагрузка "SMP M"',
    todayDate,
    { contains: true }
  );

  var l2l1Mena1xDeposits = valueNearAnchor_(
    SHEETS.L2,
    'L2/L1 Mena 1x',
    'Суммарное кол-во Депозиты',
    todayDate,
    3
  );

  var l2l1MenaLeads1xDeposits = valueNearAnchor_(
    SHEETS.L2,
    'L2/L1 Mena Leads 1x',
    'Суммарное кол-во Депозиты',
    todayDate,
    3
  );

  v.l2l1Total = sumValues_(
    l2l1Mena1xDeposits,
    l2l1MenaLeads1xDeposits
  );

  v.l1Total = valueByLabel_(
    SHEETS.L1,
    'уммарное кол-во нагрузки',
    todayDate,
    { contains: true }
  );

  v.fraudTotal = valueByLabel_(
    SHEETS.FRAUD,
    'уммарное кол-во нагрузки',
    todayDate,
    { contains: true }
  );

  // Зависшие (24+).
  var zavApiTotal = valueByLabel_(
    SHEETS.ZAVISSHIE,
    'API',
    todayDate
  );

  var zavPspTotal = valueByLabel_(
    SHEETS.ZAVISSHIE,
    'PSP',
    todayDate
  );

  v.zavPspApi = sumValues_(zavApiTotal, zavPspTotal);

  var zavBtM24hMena1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena 1x',
    'PT 24 часа (ПК)',
    todayDate,
    10
  );

  var zavBtM24hMenaLeads1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena Leads 1x',
    'PT 24 часа (ПК)',
    todayDate,
    10
  );

  v.zavBtM = sumValues_(
    zavBtM24hMena1x,
    zavBtM24hMenaLeads1x
  );

  v.zavSmp = valueByLabel_(
    SHEETS.ZAVISSHIE,
    'SMP M',
    todayDate
  );

  var btInProgressMena1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena 1x',
    'PT 24 часа In Progress (M)',
    todayDate,
    10
  );

  var btInProgressMenaLeads1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena Leads 1x',
    'PT 24 часа In Progress (M)',
    todayDate,
    10
  );

  var smpInProgress24h = valueByLabel_(
    SHEETS.ZAVISSHIE,
    'PT 24 часа In Progress',
    todayDate
  );

  v.inProgressBt = sumValues_(
    btInProgressMena1x,
    btInProgressMenaLeads1x,
    smpInProgress24h
  );

  var btSentMena1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena 1x',
    'BT Sent for processing (M) 72h+',
    todayDate,
    10
  );

  var btSentMenaLeads1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena Leads 1x',
    'BT Sent for processing (M) 72h+',
    todayDate,
    10
  );

  v.btSentTotal = sumValues_(btSentMena1x, btSentMenaLeads1x);

  var btNewMena1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena 1x',
    'New request (M) 72h+',
    todayDate,
    10
  );

  var btNewMenaLeads1x = valueNearAnchor_(
    SHEETS.ZAVISSHIE,
    'Mena Leads 1x',
    'New request (M) 72h+',
    todayDate,
    10
  );

  v.btNewTotal = sumValues_(btNewMena1x, btNewMenaLeads1x);

  v.smpSent = valueByLabel_(
    SHEETS.ZAVISSHIE,
    'SMP Sent for processing 72h+',
    todayDate
  );

  v.smpNew = valueByLabel_(
    SHEETS.ZAVISSHIE,
    'New request 72h+',
    todayDate
  );

  return v;
}

/**
 * Сумма любого числа значений.
 * Пустые и нечисловые значения игнорируются.
 * Если все значения пустые, возвращается пустая строка.
 */
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

/** Добавляет ручное числовое значение к автоматическому. */
function addManualValue_(autoValue, manualValue) {
  var manual = toNumber_(manualValue);
  if (manual === null) return autoValue;

  var auto = toNumber_(autoValue);
  return (auto === null ? 0 : auto) + manual;
}

// ---------- сборка итогового текста ----------

var TEMPLATE =
  '@leads\n' +
  '\n' +
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
  'API (Team A/B): {{apiTeamA}} / {{apiTeamB}}\n' +
  'PSP: {{pspTotal}} / {{pspSecond}}\n' +
  'BT M: {{btTotal}} / {{btMenaLeads1x}}\n' +
  'SMP M: {{smpTotal}} / {{smpSecond}}\n' +
  'L2/L1 (депозиты): {{l2l1Total}} / {{l2l1MenaLeads1x}}\n' +
  '\n' +
  'Fraud:\n' +
  'Нагрузка: {{fraudTotal}} / {{fraudMenaLeads1x}}\n' +
  '\n' +
  'L1:\n' +
  'Нагрузка: {{l1Total}} / {{l1MenaLeads1x}}\n' +
  '\n' +
  'Зависшие (24+):\n' +
  'PSP/API {{zavPspApi}} / {{zavApi}}\n' +
  'BT M: {{zavBtM}} / {{zavBtMenaLeads1x}}\n' +
  'SMP M: {{zavSmp}} / {{zavSmpSecond}}\n' +
  'In Progress (BT/SMP): {{inProgressBt}} / {{inProgressSmp}}\n' +
  '\n' +
  'BT Sent for processing (M) 72h+ MENA 1X / Leads 1X: {{btSentTotal}} / {{btSentSecond}}\n' +
  'BT New request (M) 72h+ MENA 1X / Leads 1X: {{btNewTotal}} / {{btNewSecond}}\n' +
  'SMP Sent for processing (M) 72h+: {{smpSent}} / {{smpSentSecond}}\n' +
  'SMP New Request 72h+: {{smpNew}} / {{smpNewSecond}}\n' +
  '\n' +
  'Чат 72: {{chatLink}}\n' +
  'Ожидают ответа: {{waiting}}\n' +
  '{{waitingLinks}}' +
  '\n' +
  'Массовый Approved: {{approved}}\n' +
  'Ошибка - Статус: {{errorStatus}}\n' +
  '\n' +
  'Проблемные области:\n' +
  '{{problems}}';

// Поля вечерних данных. Они формируют второе значение в строках с «/».
var EVENING_VALUE_KEYS = [
  'apiTeamB',
  'pspSecond',
  'btMenaLeads1x',
  'smpSecond',
  'l2l1MenaLeads1x',
  'l1MenaLeads1x',
  'fraudMenaLeads1x',
  'zavApi',
  'zavBtMenaLeads1x',
  'zavSmpSecond',
  'inProgressSmp',
  'btSentSecond',
  'btNewSecond',
  'smpSentSecond',
  'smpNewSecond'
];

/**
 * Формирует отчёт из автоматических значений таблицы и ручных полей диалога.
 */
function generateReport(payload) {
  payload = payload || {};

  var todayDate = parseIsoDate_(payload.date);
  var raw = collectReportValues_(todayDate);
  var warnings = [];

  // На 1-е число не обращаемся к таблице предыдущего месяца.
  if (todayDate.getDate() === 1) {
    warnings.push(
      'Создано тикетов: выбрано 1-е число месяца. Значение за прошлый день находится в предыдущей месячной таблице и автоматически не подставляется.'
    );
  }

  // Проверяем автоматические значения до прибавления ручных выводов,
  // чтобы ручной ввод не скрывал проблему поиска значения в таблице.
  for (var rk in raw) {
    if (rk === 'created' && todayDate.getDate() === 1) continue;

    if (raw[rk] === '' && FIELD_LABELS[rk]) {
      warnings.push('Не найдено в таблице: ' + FIELD_LABELS[rk]);
    }
  }

  // Выводы L1-L2: ручное значение прибавляется к тому,
  // что скрипт посчитал по датам фиксации в таблице выводов.
  raw.vyvody = addManualValue_(raw.vyvody, payload.vyvodyL1L2);

  var fields = {
    reportDate: formatDisplayDate_(todayDate),
    chatLink: payload.chatLink || '',
    waiting: payload.waiting || '',
    waitingLinks: String(payload.waitingLinks || '').trim()
      ? String(payload.waitingLinks).trim() + '\n' : '',
    approved: payload.approved || '',
    errorStatus: payload.errorStatus || '',
    problems: payload.problems || ''
  };

  for (var key in raw) {
    fields[key] = formatNum_(raw[key]);
  }

  EVENING_VALUE_KEYS.forEach(function (key) {
    fields[key] = formatNum_(payload[key] || '');
  });

  var text = TEMPLATE;

  for (var k in fields) {
    text = text.replace(
      new RegExp('{{' + k + '}}', 'g'),
      fields[k]
    );
  }

  return {
    text: text,
    warnings: warnings
  };
}

function parseIsoDate_(iso) {
  if (typeof iso !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
    throw new Error('Выберите корректную дату отчёта.');
  }

  var parts = iso.split('-');
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  var day = Number(parts[2]);
  var date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error('Выберите корректную дату отчёта.');
  }

  return date;
}

function formatDisplayDate_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd.MM.yyyy'
  );
}
