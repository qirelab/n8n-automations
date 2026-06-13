/**
 * ============================================================
 *  ДНЕВНИК КОМПАНИИ — Apps Script (Web App)
 * ============================================================
 *
 *  НАЗНАЧЕНИЕ
 *  ----------
 *  Принимает от n8n (после расшифровки голосового в Telegram) запись дня и:
 *   1. Сохраняет расшифровку в Google Drive по структуре  Год / Месяц / Неделя.
 *   2. Добавляет краткое AI-резюме дня в месячный Google Doc — одна страница
 *      на один день (новый день = новая страница).
 *   3. Обновляет реестр Google Sheets: вкладка на каждый год, строки = месяцы,
 *      со ссылками на месячный документ и папку.
 *
 *  УСТАНОВКА (делается один раз)
 *  -----------------------------
 *  1. Откройте https://script.google.com → «Новый проект».
 *  2. Удалите код в файле Code.gs и вставьте ВЕСЬ этот файл.
 *  3. Project Settings (⚙) → Script Properties → добавьте:
 *        DIARY_TOKEN            — любой секрет (тот же впишется в n8n).
 *        DIARY_SPREADSHEET_ID   — можно оставить пустым: создастся сам
 *                                 при первом запуске setup() (см. ниже).
 *  4. Запустите один раз функцию  setup()  (Run → setup). Она:
 *        • выдаст разрешения (нажмите Allow);
 *        • создаст реестр-таблицу, если её ещё нет, и впишет её ID
 *          обратно в Script Properties;
 *        • сгенерирует DIARY_TOKEN, если он пуст;
 *        • выведет в лог токен и ссылку на таблицу.
 *  5. Deploy → New deployment → тип «Web app»:
 *        Execute as: Me
 *        Who has access: Anyone
 *     Скопируйте Web app URL.
 *  6. В n8n (workflow «WF-01 Diary Capture»):
 *        • в ноде «POST to Apps Script» вставьте Web app URL;
 *        • в ноде «Build Payload» впишите тот же DIARY_TOKEN.
 *  7. Корневую папку Drive (ROOT_FOLDER_ID ниже) расшарьте на аккаунт,
 *     под которым деплоится скрипт (доступ «Редактор»).
 *
 * ============================================================
 */

// ======================== CONFIG =============================
var CONFIG = {
  // Корневая папка дневника на Google Drive (из URL папки).
  ROOT_FOLDER_ID: '1_MLUQD89nBAMKUXg1BuvJTgBVzNeIR2q',

  // Таймзона для вычисления даты/недели. Должна совпадать с TZ в ноде
  // «Build Payload» в n8n.
  TIMEZONE: 'Europe/Warsaw',

  // Секреты берутся из Script Properties (Project Settings → Script Properties).
  TOKEN_PROP: 'DIARY_TOKEN',
  SPREADSHEET_ID_PROP: 'DIARY_SPREADSHEET_ID',

  // Реестр-таблица. Используется существующая таблица (ID ниже). Приоритет:
  // Script Property DIARY_SPREADSHEET_ID → SPREADSHEET_ID_DEFAULT → создать новую.
  SPREADSHEET_ID_DEFAULT: '1G5vFw0inl0HR7xRPHJ4O86t4fB5xcFlNT0O7GJdE8iQ',
  // Название реестр-таблицы, если её всё же придётся создавать автоматически.
  SPREADSHEET_NAME: 'Дневник компании — Реестр',

  MONTHS_RU: ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
              'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'],
  WEEKDAYS_RU: ['воскресенье', 'понедельник', 'вторник', 'среда',
                'четверг', 'пятница', 'суббота'],

  // Заголовки колонок в реестре (строка = месяц).
  REGISTRY_HEADERS: ['Месяц', 'Дней с записями', 'Документ месяца',
                     'Папка Drive', 'Последний заголовок', 'Обновлено', '_dates']
};

// ======================== WEB APP ENTRY ======================

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    var expected = props_().getProperty(CONFIG.TOKEN_PROP);
    if (!expected || payload.token !== expected) {
      return json_({ ok: false, error: 'unauthorized' });
    }

    var result = processEntry_(payload);
    return json_(Object.assign({ ok: true }, result));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'company-diary', time: now_() });
}

// ======================== CORE ===============================

/**
 * payload = { token, date:'YYYY-MM-DD', transcript:'...',
 *             summary:{ Заголовок, Итоги[], Инсайты[], Проблемы[], 'Следующие шаги'[] } }
 */
function processEntry_(payload) {
  var date = payload.date;                      // YYYY-MM-DD
  var parts = date.split('-');
  var year = parts[0];
  var monthIdx = parseInt(parts[1], 10) - 1;    // 0..11
  var day = parseInt(parts[2], 10);
  var mm = parts[1];
  var monthName = CONFIG.MONTHS_RU[monthIdx];
  var weekNum = Math.ceil(day / 7);             // Неделя 1..5
  var summary = payload.summary || {};

  // --- 1. Папки:  Root / 2026 / 06 Июнь / Неделя 3 ---
  var root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  var yearFolder = childFolder_(root, year);
  var monthFolder = childFolder_(yearFolder, mm + ' ' + monthName);
  var weekFolder = childFolder_(monthFolder, 'Неделя ' + weekNum);

  // --- 2. Сохраняем расшифровку в папку недели ---
  var transcriptUrl = saveTranscript_(weekFolder, date, payload.transcript || '');

  // --- 3. Месячный документ: одна страница на день ---
  var doc = getOrCreateMonthlyDoc_(monthFolder, year, mm, monthName);
  appendDaySummary_(doc, year, monthIdx, day, summary);
  var monthlyDocUrl = doc.getUrl();

  // --- 4. Реестр ---
  updateRegistry_(year, mm, monthName, date, summary, monthlyDocUrl, monthFolder.getUrl());

  return {
    date: date,
    transcriptUrl: transcriptUrl,
    monthlyDocUrl: monthlyDocUrl,
    folderUrl: weekFolder.getUrl()
  };
}

// ----- Drive helpers -----

function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function saveTranscript_(folder, date, text) {
  // Имя: «Расшифровка 2026-06-12 (14-30)» — время добавляем, чтобы не затирать
  // несколько записей за один день.
  var time = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH-mm');
  var name = 'Расшифровка ' + date + ' (' + time + ')';
  var doc = DocumentApp.create(name);
  doc.getBody().setText(text || '(пустая расшифровка)');
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);
  return doc.getUrl();
}

// ----- Monthly Doc helpers -----

function getOrCreateMonthlyDoc_(monthFolder, year, mm, monthName) {
  var name = 'Дневник компании — ' + year + '-' + mm + ' (' + monthName + ')';
  var it = monthFolder.getFilesByName(name);
  if (it.hasNext()) {
    return DocumentApp.openById(it.next().getId());
  }
  var doc = DocumentApp.create(name);
  var body = doc.getBody();
  body.clear();
  body.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.TITLE);
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(monthFolder);
  return DocumentApp.openById(doc.getId());
}

var DAY_PREFIX = '🗓 ';

/**
 * Добавляет резюме дня. Новый день = новая страница (page break).
 * Повторная запись за тот же день добавляется в конец секции этого дня.
 */
function appendDaySummary_(doc, year, monthIdx, day, summary) {
  var body = doc.getBody();
  var dateKey = year + '-' + pad2_(monthIdx + 1) + '-' + pad2_(day);
  var weekday = CONFIG.WEEKDAYS_RU[new Date(year, monthIdx, day).getDay()];
  var headingText = DAY_PREFIX + dateKey + ' (' + weekday + ')';
  var title = (summary['Заголовок'] || '').trim();
  if (title) headingText += ' — ' + title;

  var existingIdx = findDayHeadingIndex_(body, DAY_PREFIX + dateKey);

  if (existingIdx === -1) {
    // Новый день -> новая страница в конце документа.
    if (body.getNumChildren() > 1) body.appendPageBreak();
    body.appendParagraph(headingText).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    appendSummaryBody_(body, summary, null);
  } else {
    // Тот же день -> дополнение в конце секции (перед след. днём / концом).
    var insertAt = nextDayOrEnd_(body, existingIdx);
    var time = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm');
    insertAt = insertHeading_(body, insertAt, '➕ Дополнение ' + time,
                              DocumentApp.ParagraphHeading.HEADING3);
    appendSummaryBody_(body, summary, insertAt);
  }
  doc.saveAndClose();
}

function findDayHeadingIndex_(body, prefix) {
  var n = body.getNumChildren();
  for (var i = 0; i < n; i++) {
    var el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
          p.getText().indexOf(prefix) === 0) {
        return i;
      }
    }
  }
  return -1;
}

// Индекс начала следующего дня (его page break) или конец документа.
function nextDayOrEnd_(body, fromIdx) {
  var n = body.getNumChildren();
  for (var i = fromIdx + 1; i < n; i++) {
    var el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH &&
        el.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING1 &&
        el.asParagraph().getText().indexOf(DAY_PREFIX) === 0) {
      // Перед заголовком дня стоит page break — вставляем перед ним.
      var prev = body.getChild(i - 1);
      if (prev.getType() === DocumentApp.ElementType.PAGE_BREAK) return i - 1;
      return i;
    }
  }
  return n; // конец документа
}

/**
 * Печатает тело резюме (секции Итоги/Инсайты/Проблемы/Следующие шаги).
 * Если insertAt === null — добавляет в конец; иначе вставляет с этого индекса
 * и возвращает следующий свободный индекс.
 */
function appendSummaryBody_(body, summary, insertAt) {
  var sections = [
    ['Итоги', summary['Итоги']],
    ['Инсайты', summary['Инсайты']],
    ['Проблемы', summary['Проблемы']],
    ['Следующие шаги', summary['Следующие шаги']]
  ];
  var any = false;
  for (var s = 0; s < sections.length; s++) {
    var label = sections[s][0];
    var items = sections[s][1];
    if (!Array.isArray(items) || items.length === 0) continue;
    any = true;
    if (insertAt === null) {
      body.appendParagraph(label).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      for (var j = 0; j < items.length; j++) {
        body.appendParagraph('•  ' + items[j]).setHeading(DocumentApp.ParagraphHeading.NORMAL);
      }
    } else {
      insertAt = insertHeading_(body, insertAt, label, DocumentApp.ParagraphHeading.HEADING3);
      for (var k = 0; k < items.length; k++) {
        body.insertParagraph(insertAt++, '•  ' + items[k])
            .setHeading(DocumentApp.ParagraphHeading.NORMAL);
      }
    }
  }
  if (!any) {
    var txt = '(нет структурированных пунктов)';
    if (insertAt === null) body.appendParagraph(txt);
    else body.insertParagraph(insertAt++, txt);
  }
  return insertAt;
}

function insertHeading_(body, idx, text, heading) {
  body.insertParagraph(idx, text).setHeading(heading);
  return idx + 1;
}

// ----- Registry (Sheets) helpers -----

function getSpreadsheet_() {
  var id = props_().getProperty(CONFIG.SPREADSHEET_ID_PROP) || CONFIG.SPREADSHEET_ID_DEFAULT;
  if (id) return SpreadsheetApp.openById(id);
  // Запасной вариант: создаём и кладём в корневую папку дневника.
  var ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  var file = DriveApp.getFileById(ss.getId());
  file.moveTo(DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID));
  props_().setProperty(CONFIG.SPREADSHEET_ID_PROP, ss.getId());
  return ss;
}

function getOrCreateYearSheet_(ss, year) {
  var sheet = ss.getSheetByName(year);
  if (!sheet) {
    sheet = ss.insertSheet(year);
    sheet.appendRow(CONFIG.REGISTRY_HEADERS);
    sheet.getRange(1, 1, 1, CONFIG.REGISTRY_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideColumns(CONFIG.REGISTRY_HEADERS.indexOf('_dates') + 1);
    // Удаляем стартовую вкладку по умолчанию ТОЛЬКО если она пустая —
    // существующие данные в таблице не трогаем.
    var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Лист1');
    if (def && def.getName() !== year && ss.getSheets().length > 1 &&
        def.getLastRow() === 0 && def.getLastColumn() === 0) {
      ss.deleteSheet(def);
    }
  }
  return sheet;
}

function updateRegistry_(year, mm, monthName, date, summary, docUrl, folderUrl) {
  var ss = getSpreadsheet_();
  var sheet = getOrCreateYearSheet_(ss, year);
  var monthLabel = mm + ' — ' + monthName;
  var datesCol = CONFIG.REGISTRY_HEADERS.indexOf('_dates') + 1;

  // Ищем строку месяца.
  var last = sheet.getLastRow();
  var rowIdx = -1;
  if (last >= 2) {
    var labels = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < labels.length; i++) {
      if (labels[i][0] === monthLabel) { rowIdx = i + 2; break; }
    }
  }

  var title = (summary['Заголовок'] || '').toString();
  var docCell = '=HYPERLINK("' + docUrl + '";"Открыть")';
  var folderCell = '=HYPERLINK("' + folderUrl + '";"Папка")';
  var stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');

  if (rowIdx === -1) {
    sheet.appendRow([monthLabel, 1, docCell, folderCell, title, stamp, date]);
    sortYearSheet_(sheet);
  } else {
    // Учитываем уникальные дни записей.
    var prevDates = (sheet.getRange(rowIdx, datesCol).getValue() || '').toString();
    var set = {};
    prevDates.split(',').forEach(function (d) { if (d) set[d.trim()] = 1; });
    set[date] = 1;
    var allDates = Object.keys(set).sort().join(',');
    var count = Object.keys(set).length;
    sheet.getRange(rowIdx, 1, 1, CONFIG.REGISTRY_HEADERS.length)
         .setValues([[monthLabel, count, docCell, folderCell, title, stamp, allDates]]);
  }
}

function sortYearSheet_(sheet) {
  var last = sheet.getLastRow();
  if (last > 2) sheet.getRange(2, 1, last - 1, CONFIG.REGISTRY_HEADERS.length).sort(1);
}

// ======================== UTIL ===============================

function props_() { return PropertiesService.getScriptProperties(); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function now_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================== SETUP / TEST =======================

/** Запустите один раз вручную после вставки кода. */
function setup() {
  var p = props_();
  if (!p.getProperty(CONFIG.TOKEN_PROP)) {
    p.setProperty(CONFIG.TOKEN_PROP, Utilities.getUuid());
  }
  // Проверяем доступ к корневой папке (выбросит ошибку, если нет доступа).
  var root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  var ss = getSpreadsheet_();

  Logger.log('DIARY_TOKEN: ' + p.getProperty(CONFIG.TOKEN_PROP));
  Logger.log('Реестр: ' + ss.getUrl());
  Logger.log('Корневая папка: ' + root.getName() + ' (' + root.getUrl() + ')');
  Logger.log('Готово. Скопируйте токен в n8n и задеплойте Web App.');
}

/** Тестовая запись — проверка всей цепочки без n8n. */
function testEntry() {
  var res = processEntry_({
    token: props_().getProperty(CONFIG.TOKEN_PROP),
    date: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    transcript: 'Тестовая расшифровка дня.',
    summary: {
      'Заголовок': 'Тестовый день',
      'Итоги': ['Проверили автоматизацию дневника'],
      'Инсайты': ['Связка Telegram → Whisper → Drive работает'],
      'Проблемы': [],
      'Следующие шаги': ['Записать первую реальную запись']
    }
  });
  Logger.log(JSON.stringify(res, null, 2));
}
