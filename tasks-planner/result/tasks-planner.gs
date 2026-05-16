/**
 * ============================================================
 *  AUTOMATIC TASK PLANNER — Apps Script
 * ============================================================
 *
 *  ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ / HOW TO USE
 *  ------------------------------------------
 *
 *  1. Откройте Google Таблицу с задачами:
 *     https://docs.google.com/spreadsheets/d/1ZHpmKY-e4IphN6z1_1Md2nsfUKfTthUp-RD8ggBU0ps/edit
 *
 *  2. В меню таблицы нажмите:  Расширения → Apps Script
 *     (Extensions → Apps Script)
 *
 *  3. Удалите весь код в открывшемся редакторе (файл Code.gs).
 *
 *  4. Скопируйте и вставьте ВЕСЬ этот файл целиком в редактор.
 *
 *  5. Нажмите «Сохранить» (Ctrl+S).
 *
 *  6. Вернитесь в Google Таблицу — через 2-3 секунды в меню
 *     появится новый пункт «Планировщик задач».
 *
 *  7. Нажмите «Планировщик задач → Запланировать задачи на завтра»
 *     (или «на сегодня»).
 *
 *  8. При первом запуске скрипт запросит разрешения на доступ
 *     к Таблицам и Календарю — нажмите «Разрешить» (Allow).
 *
 *  ПОВЕДЕНИЕ ПРИ ЗАПУСКЕ "НА СЕГОДНЯ":
 *  Если запустить планировщик в течение рабочего дня (например, в 11:00
 *  или в 14:00), задачи будут запланированы от текущего времени до конца
 *  рабочего дня (WORK_END_HOUR, по умолчанию 18:00). В прошлое ничего
 *  не ставится. Если запуск произошёл после WORK_END_HOUR, скрипт
 *  сообщит, что рабочий день уже закончился.
 *
 *  НАСТРОЙКИ (можно изменить в секции CONFIG ниже):
 *  - SHEET_NAME        — имя вкладки с задачами (по умолчанию "QIRElab")
 *  - CALENDAR_ID       — ID календаря (по умолчанию "primary")
 *  - WORK_START_HOUR   — начало рабочего дня (9)
 *  - WORK_END_HOUR     — конец рабочего дня (18)
 *  - BREAK_DURATION_MIN— длительность отдыха в минутах (60)
 *  - BREAK_AFTER_HOURS — после скольких часов работы ставить отдых (4)
 *  - SHORT_TASK_MIN    — длительность краткосрочной задачи (6 мин)
 *  - EVENT_COLOR       — цвет событий в календаре (по желанию)
 *  - GEMINI_API_KEY    — ключ для Gemini AI (анализ комментариев)
 *
 *  ВАЖНО: Для анализа комментариев нужен бесплатный Gemini API ключ.
 *  Получите его на https://aistudio.google.com/app/apikey
 *  и вставьте в CONFIG.GEMINI_API_KEY
 *
 * ============================================================
 */

// ======================== CONFIG =============================
var CONFIG = {
  SHEET_NAME: 'QIRElab',
  CALENDAR_ID: 'primary',
  WORK_START_HOUR: 9,
  WORK_END_HOUR: 18,
  BREAK_DURATION_MIN: 60,
  BREAK_AFTER_HOURS: 4,
  SHORT_TASK_MIN: 6,
  EVENT_COLOR: CalendarApp.EventColor.CYAN,
  EVENT_PREFIX: '[Auto] ',
  GEMINI_API_KEY: '***REMOVED_GEMINI_KEY***'  // Вставьте свой API-ключ Gemini (бесплатный): https://aistudio.google.com/app/apikey
};

// ======================== MENU ===============================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Планировщик задач')
    .addItem('Запланировать задачи на сегодня', 'planForToday')
    .addItem('Запланировать задачи на завтра', 'planForTomorrow')
    .addItem('Запланировать задачи на дату...', 'planForCustomDate')
    .addItem('Запланировать выбранные задачи по номерам...', 'planSelectedTasksByNumbers')
    .addSeparator()
    .addItem('Удалить автозадачи на сегодня', 'clearAutoEventsToday')
    .addItem('Удалить автозадачи на завтра', 'clearAutoEventsTomorrow')
    .addSeparator()
    .addItem('Архивировать запланированные задачи', 'archivePlannedTasks')
    .addToUi();
}

// ======================== ENTRY POINTS =======================

function planForToday() {
  var today = new Date();
  planTasksForDate(today);
}

function planForTomorrow() {
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  planTasksForDate(tomorrow);
}

function planForCustomDate() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Выберите дату',
    'Введите дату в формате ДД.ММ.ГГГГ (например, 25.03.2026):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var parts = response.getResponseText().trim().split('.');
  if (parts.length !== 3) {
    ui.alert('Неверный формат даты. Используйте ДД.ММ.ГГГГ');
    return;
  }
  var date = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
  if (isNaN(date.getTime())) {
    ui.alert('Неверная дата.');
    return;
  }
  planTasksForDate(date);
}

function planSelectedTasksByNumbers() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Запланировать выбранные задачи',
    'Введите номера задач через запятую или диапазоном.\n' +
    'Например: 1, 3, 5  или  1-3, 7\n\n' +
    'Задачи будут запланированы на сегодня, начиная с текущего времени, ' +
    'с учётом уже занятых интервалов в календаре.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  if (!input) {
    ui.alert('Номера задач не указаны.');
    return;
  }

  var requestedNumbers = parseNumberList(input);
  if (requestedNumbers.length === 0) {
    ui.alert('Не удалось распознать номера задач.\nИспользуйте формат "1, 3, 5" или "1-3, 7".');
    return;
  }

  planTasksForDate(new Date(), requestedNumbers);
}

/**
 * Parses a string like "1, 3, 5" or "1-3, 7" into a sorted, deduplicated array of integers.
 */
function parseNumberList(input) {
  var numbers = [];
  var parts = input.split(/[,;]/);
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;

    var rangeMatch = part.match(/^(\d+)\s*[-–—]\s*(\d+)$/);
    if (rangeMatch) {
      var lo = Math.min(parseInt(rangeMatch[1], 10), parseInt(rangeMatch[2], 10));
      var hi = Math.max(parseInt(rangeMatch[1], 10), parseInt(rangeMatch[2], 10));
      for (var n = lo; n <= hi; n++) {
        if (numbers.indexOf(n) < 0) numbers.push(n);
      }
    } else {
      var num = parseInt(part, 10);
      if (!isNaN(num) && numbers.indexOf(num) < 0) numbers.push(num);
    }
  }
  return numbers;
}

function clearAutoEventsToday() {
  clearAutoEvents(new Date());
}

function clearAutoEventsTomorrow() {
  var d = new Date();
  d.setDate(d.getDate() + 1);
  clearAutoEvents(d);
}

/**
 * Moves all tasks with the "Запланировано" checkbox enabled to the "Archive" tab.
 * The Archive tab preserves the same columns and structure except for "Запланировано".
 */
function archivePlannedTasks() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    ui.alert('Вкладка "' + CONFIG.SHEET_NAME + '" не найдена.');
    return;
  }

  var layout = getSheetLayout(sheet);
  if (!layout || layout.colMap.name < 0) {
    ui.alert('Не удалось определить структуру таблицы.');
    return;
  }

  var plannedCol = layout.colMap.planned;
  if (plannedCol < 0) {
    ui.alert('Колонка "Запланировано" не найдена.');
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < layout.dataStartRow) {
    ui.alert('Нет задач для архивации.');
    return;
  }

  var data = sheet.getRange(layout.dataStartRow, 1, lastRow - layout.dataStartRow + 1, lastCol).getValues();

  // Indices of columns to keep in Archive (every column except "Запланировано")
  var keepCols = [];
  for (var c = 0; c < lastCol; c++) {
    if (c !== plannedCol) keepCols.push(c);
  }

  // Collect planned rows (and their row indices in the source sheet)
  var rowsToArchive = [];
  var rowIndicesToClear = [];
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][layout.colMap.name]).trim();
    if (!name) continue;
    if (data[i][plannedCol] === true) {
      var archivedRow = [];
      for (var kc = 0; kc < keepCols.length; kc++) {
        archivedRow.push(data[i][keepCols[kc]]);
      }
      rowsToArchive.push(archivedRow);
      rowIndicesToClear.push(i + layout.dataStartRow);
    }
  }

  if (rowsToArchive.length === 0) {
    ui.alert('Нет задач с отметкой "Запланировано" для архивации.');
    return;
  }

  // Build Archive headers (without "Запланировано")
  var archiveHeaders = [];
  for (var hc = 0; hc < keepCols.length; hc++) {
    archiveHeaders.push(layout.headers[keepCols[hc]]);
  }

  // Get or create the Archive sheet with proper headers
  var archiveSheet = ss.getSheetByName('Archive');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Archive');
    archiveSheet.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders]);
  } else if (archiveSheet.getLastRow() === 0) {
    archiveSheet.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders]);
  }

  // Archive has its own numeration — locate the Number column inside the archive layout
  // and renumber rows being moved so they continue from the current max in Archive.
  var numberColInArchive = -1;
  if (layout.colMap.number >= 0) {
    for (var nc = 0; nc < keepCols.length; nc++) {
      if (keepCols[nc] === layout.colMap.number) {
        numberColInArchive = nc;
        break;
      }
    }
  }

  if (numberColInArchive >= 0) {
    var nextNumber = 1;
    if (archiveSheet.getLastRow() > 1) {
      var existing = archiveSheet.getRange(2, numberColInArchive + 1,
        archiveSheet.getLastRow() - 1, 1).getValues();
      var maxNum = 0;
      for (var en = 0; en < existing.length; en++) {
        var n = Number(existing[en][0]);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
      nextNumber = maxNum + 1;
    }
    for (var rn = 0; rn < rowsToArchive.length; rn++) {
      rowsToArchive[rn][numberColInArchive] = nextNumber++;
    }
  }

  // Append archived rows below the existing content
  var startRow = archiveSheet.getLastRow() + 1;
  archiveSheet.getRange(startRow, 1, rowsToArchive.length, archiveHeaders.length).setValues(rowsToArchive);

  // Clear archived rows from the main sheet, then compact and renumber
  for (var r = 0; r < rowIndicesToClear.length; r++) {
    sheet.getRange(rowIndicesToClear[r], 1, 1, lastCol).clearContent();
  }
  compactRows();
  renumberSheet(sheet);

  ui.alert(
    'Архивация завершена',
    'Перенесено задач в Archive: ' + rowsToArchive.length,
    ui.ButtonSet.OK
  );
}

/**
 * Re-assigns sequential numbers (1..N) in the "Номер" column of a sheet,
 * but only if the current numeration is not already correct. No-op if
 * the sheet has no number column.
 */
function renumberSheet(sheet) {
  var layout = getSheetLayout(sheet);
  if (!layout || layout.colMap.number < 0 || layout.colMap.name < 0) return;

  var numberCol = layout.colMap.number;
  var nameCol = layout.colMap.name;
  var lastRow = sheet.getLastRow();
  if (lastRow < layout.dataStartRow) return;

  var rowCount = lastRow - layout.dataStartRow + 1;
  var nameValues = sheet.getRange(layout.dataStartRow, nameCol + 1, rowCount, 1).getValues();
  var numberValues = sheet.getRange(layout.dataStartRow, numberCol + 1, rowCount, 1).getValues();

  var expected = 1;
  var corrections = []; // { row (1-based), value }
  for (var i = 0; i < rowCount; i++) {
    if (!String(nameValues[i][0]).trim()) continue;
    if (Number(numberValues[i][0]) !== expected) {
      corrections.push({ row: layout.dataStartRow + i, value: expected });
    }
    expected++;
  }

  for (var c = 0; c < corrections.length; c++) {
    sheet.getRange(corrections[c].row, numberCol + 1).setValue(corrections[c].value);
  }
}

// ======================== MAIN LOGIC =========================

/**
 * Plans tasks from the spreadsheet into the calendar for a given date.
 */
function planTasksForDate(targetDate, filterNumbers) {
  var ui = SpreadsheetApp.getUi();
  var isFilteredMode = filterNumbers && filterNumbers.length > 0;

  // Pre-check: if planning for today after the workday end, refuse early.
  // (Skipped in filtered mode, which spans up to 7 days from now.)
  var now = new Date();
  var isToday = targetDate.toDateString() === now.toDateString();
  if (!isFilteredMode && isToday && now.getHours() >= CONFIG.WORK_END_HOUR) {
    ui.alert(
      'Рабочий день уже закончился',
      'Сейчас ' + formatTime(now) + ', а рабочий день заканчивается в ' +
        CONFIG.WORK_END_HOUR + ':00.\n\n' +
        'Запланируйте задачи на завтра или используйте пункт «на дату...».',
      ui.ButtonSet.OK
    );
    return;
  }

  // 0. Compact empty rows
  compactRows();

  // 1. Read tasks from the sheet
  var tasks = readTasks();
  if (tasks.length === 0) {
    var log = Logger.getLog();
    ui.alert('Нет задач для планирования',
      'Вкладка: "' + CONFIG.SHEET_NAME + '"\n\n' +
      'Диагностика:\n' + (log || 'нет данных'),
      ui.ButtonSet.OK);
    return;
  }

  var skipped = 0;
  var notFoundNumbers = [];

  if (isFilteredMode) {
    // 1a. Filter to only the requested task numbers (preserving requested order)
    var matched = [];
    for (var f = 0; f < filterNumbers.length; f++) {
      var found = null;
      for (var t = 0; t < tasks.length; t++) {
        if (!isNaN(tasks[t].taskNumber) && tasks[t].taskNumber === filterNumbers[f]) {
          found = tasks[t];
          break;
        }
      }
      if (found) matched.push(found);
      else notFoundNumbers.push(filterNumbers[f]);
    }

    if (matched.length === 0) {
      ui.alert('Задачи не найдены',
        'В таблице нет задач с номерами: ' + filterNumbers.join(', '),
        ui.ButtonSet.OK);
      return;
    }
    tasks = matched;
  } else {
    // 1b. Filter out tasks already scheduled on other days (default mode)
    var alreadyScheduled = getAlreadyScheduledTaskNames(targetDate);
    tasks = tasks.filter(function(task) {
      if (alreadyScheduled.indexOf(task.name) >= 0) {
        skipped++;
        return false;
      }
      return true;
    });

    if (tasks.length === 0) {
      ui.alert('Все задачи уже запланированы на другие дни.');
      return;
    }
  }

  // 2. Analyze comments with Gemini AI and separate pinned vs flexible tasks
  var scheduleHints = analyzeCommentsWithGemini(tasks, targetDate);
  for (var t = 0; t < tasks.length; t++) {
    tasks[t].schedule = scheduleHints[t] || { matchesDate: false, preferredStartHour: null };
  }

  // Pinned tasks: Gemini determined a specific time on the target date
  var pinnedTasks = [];
  var flexibleTasks = [];
  for (var p = 0; p < tasks.length; p++) {
    var sch = tasks[p].schedule;
    if (sch && sch.matchesDate && sch.preferredStartHour !== null) {
      pinnedTasks.push(tasks[p]);
    } else {
      flexibleTasks.push(tasks[p]);
    }
  }

  // Sort pinned by their preferred start time
  pinnedTasks.sort(function(a, b) {
    return a.schedule.preferredStartHour - b.schedule.preferredStartHour;
  });

  // Sort flexible by priority
  flexibleTasks = sortByPriority(flexibleTasks);

  // 3. Build free time slots
  var slots;
  if (isFilteredMode) {
    // For selected tasks: span up to 7 days starting from now, no break.
    slots = getFreeSlotsRange(targetDate, 7);
  } else {
    slots = getFreeSlots(targetDate);
    // 4. Insert break
    slots = insertBreak(slots);
  }

  // 5a. Schedule pinned tasks at their preferred times first
  var pinnedResult = schedulePinnedTasks(pinnedTasks, slots, targetDate);
  slots = pinnedResult.remainingSlots;

  // 5b. Schedule flexible tasks into remaining slots
  var flexResult = scheduleTasks(flexibleTasks, slots, targetDate);

  var allScheduledRows = pinnedResult.scheduledRows.concat(flexResult.scheduledRows);
  var totalScheduled = pinnedResult.scheduled + flexResult.scheduled;
  var totalUnscheduled = (pinnedTasks.length - pinnedResult.scheduled) +
                         (flexibleTasks.length - flexResult.scheduled);

  // 6. Set checkboxes for scheduled tasks
  markPlanned(allScheduledRows, true);

  // 7. Show summary
  var windowInfo = '';
  if (!isFilteredMode) {
    var effectiveStart = getEffectiveStartTime(targetDate, CONFIG.WORK_START_HOUR);
    windowInfo = 'Окно планирования: ' + formatTime(effectiveStart) +
      ' – ' + CONFIG.WORK_END_HOUR + ':00';
    if (isToday && effectiveStart.getHours() > CONFIG.WORK_START_HOUR) {
      windowInfo += ' (сегодня, начиная с текущего времени)';
    }
    windowInfo += '\n';
  }

  ui.alert(
    'Планирование завершено',
    (isFilteredMode
      ? 'Начиная с: ' + formatDate(targetDate) + ' (текущего времени, до 7 дней вперёд)\nРежим: выбранные задачи по номерам\n'
      : 'Дата: ' + formatDate(targetDate) + '\n') +
    windowInfo +
    'Запланировано задач: ' + totalScheduled + ' из ' + tasks.length + '\n' +
    (pinnedResult.scheduled > 0 ? 'Из них привязано ко времени (из комментария): ' + pinnedResult.scheduled + '\n' : '') +
    (skipped > 0 ? 'Пропущено (уже в календаре на другие дни): ' + skipped + '\n' : '') +
    (notFoundNumbers.length > 0 ? 'Не найдены номера: ' + notFoundNumbers.join(', ') + '\n' : '') +
    (totalUnscheduled > 0
      ? 'Не хватило времени для ' + totalUnscheduled + ' задач(и).'
      : 'Все задачи размещены.'),
    ui.ButtonSet.OK
  );
}

/**
 * Reads tasks from the configured sheet tab.
 * Expected columns: Номер | Наименование | Приоритет | Характер задач | Время | Проект | Сфера
 */
function readTasks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log('Вкладка "' + CONFIG.SHEET_NAME + '" не найдена.');
    return [];
  }

  var layout = getSheetLayout(sheet);
  if (!layout) {
    var data = sheet.getDataRange().getValues();
    var diagRows = data.slice(0, 5).map(function(r, idx) {
      return 'Строка ' + (idx + 1) + ': ' + JSON.stringify(r);
    });
    Logger.log('Не удалось найти строку заголовков.\n' +
      'Всего строк: ' + data.length + ', колонок: ' + (data[0] ? data[0].length : 0) + '\n' +
      diagRows.join('\n'));
    return [];
  }

  var colMap = layout.colMap;
  Logger.log('Заголовки в строке ' + layout.headerRowIndex + ': ' + JSON.stringify(layout.headerNames));
  Logger.log('Колонки: ' + JSON.stringify(colMap));

  if (colMap.name < 0) {
    Logger.log('ОШИБКА: колонка "Наименование" не найдена среди заголовков.');
    return [];
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < layout.dataStartRow) return [];

  var data = sheet.getRange(layout.dataStartRow, 1, lastRow - layout.dataStartRow + 1, lastCol).getValues();

  var tasks = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = colMap.name >= 0 ? String(row[colMap.name]).trim() : '';
    if (!name) continue;

    var taskNumber = colMap.number >= 0 ? Number(row[colMap.number]) : NaN;
    var priorityLetter = colMap.priority >= 0 ? String(row[colMap.priority]).trim().toUpperCase() : 'D';
    var taskType = colMap.taskType >= 0 ? normalizeTaskType(String(row[colMap.taskType]).trim()) : 'budgeted';
    var timeVal = colMap.time >= 0 ? parseTime(row[colMap.time]) : 0;
    var project = colMap.project >= 0 ? String(row[colMap.project]).trim() : '';
    var sphere = colMap.sphere >= 0 ? String(row[colMap.sphere]).trim() : '';
    var comment = colMap.comment >= 0 ? String(row[colMap.comment]).trim() : '';

    // Determine duration in minutes
    var durationMin;
    if (taskType === 'short') {
      durationMin = CONFIG.SHORT_TASK_MIN;
    } else if (taskType === 'budgeted' && timeVal > 0) {
      durationMin = timeVal;
    } else if (taskType === 'timebound') {
      // Time-bound tasks: use specified time or default 30 min
      durationMin = timeVal > 0 ? timeVal : 30;
    } else {
      durationMin = 30; // fallback
    }

    tasks.push({
      rowIndex: i + layout.dataStartRow,
      taskNumber: taskNumber,
      name: name,
      priority: priorityLetter,
      taskType: taskType,
      durationMin: durationMin,
      project: project,
      sphere: sphere,
      comment: comment
    });
  }

  return tasks;
}

/**
 * Sorts tasks by: 1) Priority letter (A > B > C > D),
 * 2) Task type (timebound > short > budgeted).
 */
function sortByPriority(tasks) {
  var priorityOrder = { 'A': 0, 'B': 1, 'C': 2, 'D': 3 };
  var typeOrder = { 'timebound': 0, 'short': 1, 'budgeted': 2 };

  tasks.sort(function(a, b) {
    var pa = priorityOrder[a.priority] !== undefined ? priorityOrder[a.priority] : 4;
    var pb = priorityOrder[b.priority] !== undefined ? priorityOrder[b.priority] : 4;
    if (pa !== pb) return pa - pb;

    var ta = typeOrder[a.taskType] !== undefined ? typeOrder[a.taskType] : 3;
    var tb = typeOrder[b.taskType] !== undefined ? typeOrder[b.taskType] : 3;
    return ta - tb;
  });

  return tasks;
}

/**
 * Returns the effective start time for the working window on the given date.
 * If targetDate is today and the current time is already past workStartHour,
 * starts from "now" (rounded up to the next 5-minute mark) instead — so tasks
 * are never scheduled in the past.
 * If targetDate is in the future or current time is still before workStartHour,
 * returns workStartHour on that date.
 */
function getEffectiveStartTime(targetDate, workStartHour) {
  var workdayStart = new Date(targetDate);
  workdayStart.setHours(workStartHour, 0, 0, 0);

  var now = new Date();
  if (workdayStart.toDateString() !== now.toDateString()) return workdayStart;
  if (now <= workdayStart) return workdayStart;

  // Round up to the next 5-minute mark for cleaner slots
  var mins = now.getMinutes();
  var roundedMins = Math.ceil(mins / 5) * 5;
  var effective = new Date(now);
  effective.setMinutes(roundedMins, 0, 0);
  return effective;
}

/**
 * Gets free time slots on the given date between startHour and endHour,
 * considering existing calendar events.
 * Defaults: WORK_START_HOUR..WORK_END_HOUR.
 * When targetDate is today and current time is inside the working window,
 * the start of the window is shifted to "now" so nothing is scheduled in the past.
 */
function getFreeSlots(targetDate, startHour, endHour) {
  if (startHour === undefined) startHour = CONFIG.WORK_START_HOUR;
  if (endHour === undefined) endHour = CONFIG.WORK_END_HOUR;

  var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) cal = CalendarApp.getDefaultCalendar();

  var dayStart = getEffectiveStartTime(targetDate, startHour);

  var dayEnd = new Date(targetDate);
  if (endHour >= 24) {
    dayEnd.setDate(dayEnd.getDate() + 1);
    dayEnd.setHours(0, 0, 0, 0);
  } else {
    dayEnd.setHours(endHour, 0, 0, 0);
  }

  // If current time is already past the window, no slots available
  if (dayStart >= dayEnd) return [];

  // Get existing events for that day within working hours
  var events = cal.getEvents(dayStart, dayEnd);

  // Collect busy intervals (only non-all-day events)
  var busy = [];
  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (ev.isAllDayEvent()) continue;
    var evStart = ev.getStartTime();
    var evEnd = ev.getEndTime();
    // Clamp to working hours
    if (evStart < dayStart) evStart = new Date(dayStart);
    if (evEnd > dayEnd) evEnd = new Date(dayEnd);
    if (evStart < evEnd) {
      busy.push({ start: evStart, end: evEnd });
    }
  }

  // Sort busy intervals by start time
  busy.sort(function(a, b) { return a.start - b.start; });

  // Merge overlapping intervals
  var merged = [];
  for (var j = 0; j < busy.length; j++) {
    if (merged.length === 0 || merged[merged.length - 1].end <= busy[j].start) {
      merged.push({ start: new Date(busy[j].start), end: new Date(busy[j].end) });
    } else {
      if (busy[j].end > merged[merged.length - 1].end) {
        merged[merged.length - 1].end = new Date(busy[j].end);
      }
    }
  }

  // Build free slots
  var freeSlots = [];
  var cursor = new Date(dayStart);

  for (var k = 0; k < merged.length; k++) {
    if (cursor < merged[k].start) {
      freeSlots.push({ start: new Date(cursor), end: new Date(merged[k].start) });
    }
    cursor = new Date(merged[k].end);
  }
  if (cursor < dayEnd) {
    freeSlots.push({ start: new Date(cursor), end: new Date(dayEnd) });
  }

  return freeSlots;
}

/**
 * Returns free slots from `startDate` over the next `daysAhead` days.
 * On day 0 (today) the window stretches to midnight so after-hours time
 * is usable; subsequent days use normal work hours.
 */
function getFreeSlotsRange(startDate, daysAhead) {
  var allSlots = [];
  for (var d = 0; d < daysAhead; d++) {
    var date = new Date(startDate);
    date.setDate(date.getDate() + d);
    var endHour = (d === 0) ? 24 : CONFIG.WORK_END_HOUR;
    var slots = getFreeSlots(date, CONFIG.WORK_START_HOUR, endHour);
    allSlots = allSlots.concat(slots);
  }
  return allSlots;
}

/**
 * Inserts a break (1.5h) after BREAK_AFTER_HOURS of accumulated work time.
 * Splits free slots accordingly.
 */
function insertBreak(slots) {
  var breakDurationMs = CONFIG.BREAK_DURATION_MIN * 60 * 1000;
  var breakAfterMs = CONFIG.BREAK_AFTER_HOURS * 60 * 60 * 1000;

  // Calculate total free time before the break point
  var accumulatedFree = 0;
  var breakSlotIndex = -1;
  var breakPointInSlot = null;

  for (var i = 0; i < slots.length; i++) {
    var slotDuration = slots[i].end - slots[i].start;
    if (accumulatedFree + slotDuration >= breakAfterMs) {
      breakSlotIndex = i;
      var remaining = breakAfterMs - accumulatedFree;
      breakPointInSlot = new Date(slots[i].start.getTime() + remaining);
      break;
    }
    accumulatedFree += slotDuration;
  }

  if (breakSlotIndex < 0) {
    // Not enough free time to place a break — just return as is
    return slots;
  }

  // Split the slot at the break point
  var slot = slots[breakSlotIndex];
  var breakEnd = new Date(breakPointInSlot.getTime() + breakDurationMs);

  var newSlots = slots.slice(0, breakSlotIndex);

  // Part before break
  if (breakPointInSlot > slot.start) {
    newSlots.push({ start: new Date(slot.start), end: new Date(breakPointInSlot) });
  }

  // Part after break (if break ends before the slot ends)
  if (breakEnd < slot.end) {
    newSlots.push({ start: new Date(breakEnd), end: new Date(slot.end) });
  }

  // Remaining slots — adjust if break spills into them
  for (var j = breakSlotIndex + 1; j < slots.length; j++) {
    if (breakEnd <= slots[j].start) {
      newSlots.push(slots[j]);
    } else if (breakEnd < slots[j].end) {
      newSlots.push({ start: new Date(breakEnd), end: new Date(slots[j].end) });
    }
    // else: break completely covers this slot — skip it
  }

  return newSlots;
}

/**
 * Schedules tasks into free time slots and creates calendar events.
 */
function scheduleTasks(tasks, slots, targetDate) {
  var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) cal = CalendarApp.getDefaultCalendar();

  var scheduled = 0;
  var scheduledRows = [];

  // Clone slots into a mutable list of remaining free fragments.
  // Each fragment is shrunk in place as tasks consume its leading edge.
  var remainingSlots = slots.map(function(s) {
    return { start: new Date(s.start), end: new Date(s.end) };
  });

  for (var i = 0; i < tasks.length; i++) {
    var task = tasks[i];
    var durationMs = task.durationMin * 60 * 1000;

    // Scan ALL remaining fragments and place the task in the first one that fits.
    // This way a task that doesn't fit a small leftover doesn't block subsequent
    // smaller tasks from using that leftover.
    for (var s = 0; s < remainingSlots.length; s++) {
      var slot = remainingSlots[s];
      var available = slot.end - slot.start;
      if (available < durationMs) continue;

      var eventStart = new Date(slot.start);
      var eventEnd = new Date(slot.start.getTime() + durationMs);

      var description = '';
      if (task.project) description += 'Проект: ' + task.project + '\n';
      if (task.sphere) description += 'Сфера: ' + task.sphere + '\n';
      description += 'Приоритет: ' + task.priority + '\n';
      description += 'Тип: ' + getTaskTypeLabel(task.taskType);
      if (task.comment) description += '\nКомментарий: ' + task.comment;

      var event = cal.createEvent(
        CONFIG.EVENT_PREFIX + task.name,
        eventStart,
        eventEnd,
        { description: description }
      );

      try {
        // Blue for tasks with Сфера, Green for tasks with Проект, default otherwise
        if (task.sphere) {
          event.setColor(CalendarApp.EventColor.BLUE);
        } else if (task.project) {
          event.setColor(CalendarApp.EventColor.GREEN);
        } else {
          event.setColor(CONFIG.EVENT_COLOR);
        }
      } catch (e) {
        // Color setting might fail on some calendar types — ignore
      }

      // Shrink this fragment so its remaining tail can still host more tasks.
      slot.start = eventEnd;

      scheduled++;
      scheduledRows.push(task.rowIndex);
      break;
    }
    // If the task didn't fit anywhere, keep trying the remaining (possibly shorter) tasks.
  }

  return {
    scheduled: scheduled,
    scheduledRows: scheduledRows,
    unscheduled: tasks.length - scheduled
  };
}

/**
 * Returns task names that are already scheduled as auto-events on OTHER days
 * (not on the targetDate). Looks 7 days back and 7 days forward.
 */
function getAlreadyScheduledTaskNames(targetDate) {
  var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) cal = CalendarApp.getDefaultCalendar();

  var searchStart = new Date(targetDate);
  searchStart.setDate(searchStart.getDate() - 7);
  searchStart.setHours(0, 0, 0, 0);

  var searchEnd = new Date(targetDate);
  searchEnd.setDate(searchEnd.getDate() + 7);
  searchEnd.setHours(23, 59, 59, 999);

  var targetDayStart = new Date(targetDate);
  targetDayStart.setHours(0, 0, 0, 0);
  var targetDayEnd = new Date(targetDate);
  targetDayEnd.setHours(23, 59, 59, 999);

  var events = cal.getEvents(searchStart, searchEnd);
  var names = [];

  for (var i = 0; i < events.length; i++) {
    var title = events[i].getTitle();
    if (title.indexOf(CONFIG.EVENT_PREFIX) !== 0) continue;

    var evStart = events[i].getStartTime();
    // Skip events that are on the target date itself (those will be replaced)
    if (evStart >= targetDayStart && evStart <= targetDayEnd) continue;

    var taskName = title.substring(CONFIG.EVENT_PREFIX.length);
    if (names.indexOf(taskName) < 0) {
      names.push(taskName);
    }
  }

  return names;
}

/**
 * Analyzes all task comments in a single Gemini API call.
 * Returns an array of { matchesDate: bool, preferredStartHour: number|null }
 * aligned with the tasks array.
 */
function analyzeCommentsWithGemini(tasks, targetDate) {
  var defaultResult = [];
  for (var i = 0; i < tasks.length; i++) {
    defaultResult.push({ matchesDate: false, preferredStartHour: null });
  }

  // Collect tasks that actually have comments
  var commentEntries = [];
  for (var j = 0; j < tasks.length; j++) {
    if (tasks[j].comment) {
      commentEntries.push({ index: j, comment: tasks[j].comment });
    }
  }

  if (commentEntries.length === 0) return defaultResult;

  // If no API key, skip AI analysis
  if (!CONFIG.GEMINI_API_KEY) {
    Logger.log('GEMINI_API_KEY не указан в CONFIG. Комментарии не анализируются.');
    return defaultResult;
  }

  var today = new Date();
  var dayOfWeek = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'][today.getDay()];
  var targetDayOfWeek = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'][targetDate.getDay()];

  // Build the list of comments for the prompt
  var commentList = '';
  for (var c = 0; c < commentEntries.length; c++) {
    commentList += (c + 1) + '. "' + commentEntries[c].comment + '"\n';
  }

  var prompt = 'Ты — помощник по планированию рабочего дня. ' +
    'Сегодня ' + formatDate(today) + ' (' + dayOfWeek + '). ' +
    'Дата планирования: ' + formatDate(targetDate) + ' (' + targetDayOfWeek + '). ' +
    'Рабочие часы: с ' + CONFIG.WORK_START_HOUR + ':00 до ' + CONFIG.WORK_END_HOUR + ':00.\n\n' +
    'Проанализируй комментарии к задачам и определи для каждого:\n' +
    '1) matchesDate — относится ли указание в комментарии к дате планирования (' + formatDate(targetDate) + ')? true/false\n' +
    '2) preferredStartHour — предпочтительный час начала (целое число от ' + CONFIG.WORK_START_HOUR + ' до ' + (CONFIG.WORK_END_HOUR - 1) + ') или null если не указано.\n\n' +
    'Маппинг времени суток:\n' +
    '- "утро/с утра" → 8\n' +
    '- "до обеда" → 10\n' +
    '- "обед" → 12\n' +
    '- "после обеда/днём" → 13\n' +
    '- "вечером/конец дня" → 16\n' +
    '- Конкретное время ("в 14:00") → соответствующий час\n\n' +
    'Комментарии:\n' + commentList + '\n' +
    'Ответь ТОЛЬКО валидным JSON-массивом без markdown-обёртки, без пояснений. ' +
    'Формат: [{"matchesDate": true, "preferredStartHour": 8}, ...]\n' +
    'Если в комментарии нет указаний на дату/время, верни {"matchesDate": false, "preferredStartHour": null}.\n' +
    'Количество элементов в массиве должно быть ровно ' + commentEntries.length + '.';

  try {
    var response = callGemini(prompt);
    var parsed = JSON.parse(response);

    if (!Array.isArray(parsed) || parsed.length !== commentEntries.length) {
      Logger.log('Gemini вернул неверное количество результатов: ' + response);
      return defaultResult;
    }

    // Map results back to tasks array
    for (var r = 0; r < parsed.length; r++) {
      var entry = parsed[r];
      var taskIdx = commentEntries[r].index;
      defaultResult[taskIdx] = {
        matchesDate: !!entry.matchesDate,
        preferredStartHour: (typeof entry.preferredStartHour === 'number' &&
          entry.preferredStartHour >= CONFIG.WORK_START_HOUR &&
          entry.preferredStartHour < CONFIG.WORK_END_HOUR)
          ? entry.preferredStartHour : null
      };
    }
  } catch (e) {
    Logger.log('Ошибка при вызове Gemini: ' + e.message);
  }

  return defaultResult;
}

/**
 * Calls the Gemini API (free tier) and returns the text response.
 */
function callGemini(prompt) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + CONFIG.GEMINI_API_KEY;

  var payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 1024
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code !== 200) {
    throw new Error('Gemini API HTTP ' + code + ': ' + response.getContentText());
  }

  var json = JSON.parse(response.getContentText());
  var text = json.candidates[0].content.parts[0].text;

  // Strip markdown code fences if present
  text = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();

  return text;
}

/**
 * Schedules pinned tasks at their preferred times.
 * Returns { scheduled, scheduledRows, remainingSlots }.
 */
function schedulePinnedTasks(pinnedTasks, slots, targetDate) {
  var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) cal = CalendarApp.getDefaultCalendar();

  var scheduled = 0;
  var scheduledRows = [];

  for (var i = 0; i < pinnedTasks.length; i++) {
    var task = pinnedTasks[i];
    var durationMs = task.durationMin * 60 * 1000;
    var preferredStart = new Date(targetDate);
    preferredStart.setHours(task.schedule.preferredStartHour, 0, 0, 0);

    // Find a slot that contains the preferred start time
    var placed = false;
    for (var s = 0; s < slots.length; s++) {
      var slot = slots[s];
      // Adjust: if preferred start is before slot start, use slot start
      var actualStart = preferredStart < slot.start ? slot.start : preferredStart;
      if (actualStart >= slot.end) continue;

      var available = slot.end - actualStart;
      if (available >= durationMs) {
        var eventStart = new Date(actualStart);
        var eventEnd = new Date(actualStart.getTime() + durationMs);

        var description = '';
        if (task.project) description += 'Проект: ' + task.project + '\n';
        if (task.sphere) description += 'Сфера: ' + task.sphere + '\n';
        description += 'Приоритет: ' + task.priority + '\n';
        description += 'Тип: ' + getTaskTypeLabel(task.taskType) + '\n';
        if (task.comment) description += 'Комментарий: ' + task.comment;

        var event = cal.createEvent(
          CONFIG.EVENT_PREFIX + task.name,
          eventStart,
          eventEnd,
          { description: description }
        );

        try {
          if (task.sphere) {
            event.setColor(CalendarApp.EventColor.BLUE);
          } else if (task.project) {
            event.setColor(CalendarApp.EventColor.GREEN);
          } else {
            event.setColor(CONFIG.EVENT_COLOR);
          }
        } catch (e) {}

        // Split the used slot into before/after fragments
        var newSlots = [];
        for (var k = 0; k < slots.length; k++) {
          if (k === s) {
            if (slot.start < eventStart) {
              newSlots.push({ start: new Date(slot.start), end: new Date(eventStart) });
            }
            if (eventEnd < slot.end) {
              newSlots.push({ start: new Date(eventEnd), end: new Date(slot.end) });
            }
          } else {
            newSlots.push(slots[k]);
          }
        }
        slots = newSlots;

        scheduled++;
        scheduledRows.push(task.rowIndex);
        placed = true;
        break;
      }
    }
  }

  return {
    scheduled: scheduled,
    scheduledRows: scheduledRows,
    remainingSlots: slots
  };
}

/**
 * Removes all auto-created events (with prefix) for a given date.
 */
function clearAutoEvents(targetDate) {
  var cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!cal) cal = CalendarApp.getDefaultCalendar();

  var dayStart = new Date(targetDate);
  dayStart.setHours(0, 0, 0, 0);
  var dayEnd = new Date(targetDate);
  dayEnd.setHours(23, 59, 59, 999);

  var events = cal.getEvents(dayStart, dayEnd);
  var count = 0;
  var deletedNames = [];

  for (var i = 0; i < events.length; i++) {
    var title = events[i].getTitle();
    if (title.indexOf(CONFIG.EVENT_PREFIX) === 0) {
      deletedNames.push(title.substring(CONFIG.EVENT_PREFIX.length));
      events[i].deleteEvent();
      count++;
    }
  }

  // Uncheck "Запланировано" for deleted tasks
  if (deletedNames.length > 0) {
    unmarkPlannedByName(deletedNames);
  }

  SpreadsheetApp.getUi().alert('Удалено автозадач: ' + count + ' на ' + formatDate(targetDate));
}

/**
 * Sets or clears the "Запланировано" checkbox for given row numbers.
 */
function markPlanned(rowNumbers, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var layout = getSheetLayout(sheet);
  if (!layout || layout.colMap.planned < 0) return;

  var col = layout.colMap.planned;
  for (var i = 0; i < rowNumbers.length; i++) {
    sheet.getRange(rowNumbers[i], col + 1).setValue(value);
  }
}

/**
 * Unmarks "Запланировано" for tasks matching given names.
 */
function unmarkPlannedByName(taskNames) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var layout = getSheetLayout(sheet);
  if (!layout || layout.colMap.planned < 0 || layout.colMap.name < 0) return;

  var plannedCol = layout.colMap.planned;
  var nameCol = layout.colMap.name;
  var lastRow = sheet.getLastRow();
  if (lastRow < layout.dataStartRow) return;

  var names = sheet.getRange(layout.dataStartRow, nameCol + 1, lastRow - layout.dataStartRow + 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    var cellName = String(names[i][0]).trim();
    if (taskNames.indexOf(cellName) >= 0) {
      sheet.getRange(i + layout.dataStartRow, plannedCol + 1).setValue(false);
    }
  }
}

/**
 * Removes empty rows from the task sheet and shifts data up.
 */
function compactRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var layout = getSheetLayout(sheet);
  if (!layout || layout.colMap.name < 0) return;

  var nameCol = layout.colMap.name;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < layout.dataStartRow) return;

  var dataRowCount = lastRow - layout.dataStartRow + 1;
  var data = sheet.getRange(layout.dataStartRow, 1, dataRowCount, lastCol).getValues();

  // Separate non-empty and empty rows
  var nonEmpty = [];
  var emptyCount = 0;
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][nameCol]).trim();
    if (name) {
      nonEmpty.push(data[i]);
    } else {
      emptyCount++;
    }
  }

  if (emptyCount === 0) return; // Nothing to compact

  // Write non-empty rows back starting from dataStartRow
  if (nonEmpty.length > 0) {
    sheet.getRange(layout.dataStartRow, 1, nonEmpty.length, lastCol).setValues(nonEmpty);
  }

  // Clear remaining rows that are now empty
  if (emptyCount > 0) {
    var clearStart = layout.dataStartRow + nonEmpty.length;
    if (clearStart <= lastRow) {
      sheet.getRange(clearStart, 1, lastRow - clearStart + 1, lastCol).clearContent();
    }
  }
}

// ======================== HELPERS ============================

/**
 * Returns { headerRowIndex (1-based), headers, colMap } for the task sheet.
 * Searches through the entire sheet for the header row.
 * If the header row has been displaced by sorting, moves it back to row 1.
 * Returns null if sheet or headers not found.
 */
function getSheetLayout(sheet) {
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return null;

  var headerIdx = findHeaderRow(data);
  if (headerIdx < 0) return null;

  var headers = data[headerIdx];
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  // If header is not in row 1 — fix it: move header to row 1 and shift data
  if (headerIdx > 0) {
    Logger.log('Заголовки сместились в строку ' + (headerIdx + 1) + '. Восстанавливаю порядок...');

    // Collect all rows: header first, then all data rows (skipping the header's old position)
    var reordered = [headers];
    for (var r = 0; r < data.length; r++) {
      if (r === headerIdx) continue; // skip old header position
      reordered.push(data[r]);
    }

    // Write everything back
    sheet.getRange(1, 1, reordered.length, lastCol).setValues(reordered);

    // Clear any leftover rows below (if the data shrunk)
    if (reordered.length < lastRow) {
      sheet.getRange(reordered.length + 1, 1, lastRow - reordered.length, lastCol).clearContent();
    }

    Logger.log('Заголовки возвращены в строку 1.');
  }

  var headerNames = headers.map(function(h) { return String(h).trim().toLowerCase(); });

  return {
    headerRowIndex: 1,                     // always row 1 after fix
    dataStartRow: 2,                       // data always starts at row 2
    headers: headers,                      // original header values
    headerNames: headerNames,              // lowercased trimmed
    colMap: {
      number: findCol(headerNames, ['номер', '№', 'number', '#']),
      name: findCol(headerNames, ['наименование', 'название', 'name', 'задача']),
      priority: findCol(headerNames, ['приоритет', 'priority']),
      taskType: findCol(headerNames, ['характер задач', 'характер', 'тип', 'type']),
      time: findCol(headerNames, ['время', 'time', 'длительность', 'duration']),
      project: findCol(headerNames, ['проект', 'project']),
      sphere: findCol(headerNames, ['сфера', 'sphere', 'область']),
      planned: findCol(headerNames, ['запланировано', 'planned']),
      comment: findCol(headerNames, ['комментарий', 'комментарии', 'comment', 'comments', 'примечание'])
    }
  };
}

/**
 * Finds the header row by looking for a row that contains
 * at least 2 known column keywords.
 */
function findHeaderRow(data) {
  var keywords = ['наименование', 'название', 'приоритет', 'характер',
                  'время', 'проект', 'сфера', 'номер', '№', 'запланировано',
                  'name', 'priority', 'project', 'number', 'комментар'];

  // Search ALL rows — sorting could have moved the header anywhere
  for (var r = 0; r < data.length; r++) {
    var matches = 0;
    for (var c = 0; c < data[r].length; c++) {
      var cell = String(data[r][c]).trim().toLowerCase();
      if (!cell) continue;
      // Skip cells that are just numbers (data rows, not headers)
      if (/^\d+([.,]\d+)?$/.test(cell)) continue;
      for (var k = 0; k < keywords.length; k++) {
        if (cell.indexOf(keywords[k]) >= 0) {
          matches++;
          break;
        }
      }
    }
    if (matches >= 2) return r;
  }
  return -1;
}

/**
 * Finds a column index by trying multiple possible header names.
 * Uses partial matching (indexOf) so "Наименование задачи" matches "наименование".
 */
function findCol(headers, possibleNames) {
  // First pass: exact match
  for (var i = 0; i < possibleNames.length; i++) {
    var idx = headers.indexOf(possibleNames[i]);
    if (idx >= 0) return idx;
  }
  // Second pass: partial match (header contains the keyword)
  for (var j = 0; j < possibleNames.length; j++) {
    for (var k = 0; k < headers.length; k++) {
      if (headers[k].indexOf(possibleNames[j]) >= 0) return k;
    }
  }
  return -1;
}

/**
 * Normalizes task type string to internal key.
 */
function normalizeTaskType(raw) {
  var lower = raw.toLowerCase();
  if (lower.indexOf('привязан') >= 0 || lower.indexOf('времени') >= 0 || lower.indexOf('time') >= 0) {
    return 'timebound';
  }
  if (lower.indexOf('краткосроч') >= 0 || lower.indexOf('short') >= 0) {
    return 'short';
  }
  if (lower.indexOf('бюджет') >= 0 || lower.indexOf('budget') >= 0) {
    return 'budgeted';
  }
  return 'budgeted'; // default
}

/**
 * Returns a human-readable label for task type.
 */
function getTaskTypeLabel(type) {
  var labels = {
    'timebound': 'Привязанные ко времени',
    'short': 'Краткосрочные',
    'budgeted': 'Бюджетируемые'
  };
  return labels[type] || type;
}

/**
 * Parses a time value from the sheet into minutes.
 * Plain numbers are treated as HOURS (e.g. 1.5 = 90 minutes).
 * Supports: "1:30", "1ч 30мин", 1.5, Date objects (treated as HH:MM).
 */
function parseTime(val) {
  if (!val) return 0;

  // If it's a Date object (Google Sheets time values are Date objects)
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }

  // If it's a raw number (not stringified), treat as hours
  if (typeof val === 'number') {
    return Math.round(val * 60);
  }

  var str = String(val).trim();
  if (!str) return 0;

  // "1:30" format
  var colonMatch = str.match(/^(\d+):(\d+)$/);
  if (colonMatch) {
    return parseInt(colonMatch[1]) * 60 + parseInt(colonMatch[2]);
  }

  // "1ч 30мин" or "1ч30м" format
  var ruMatch = str.match(/(\d+)\s*ч/);
  var ruMinMatch = str.match(/(\d+)\s*м/);
  if (ruMatch || ruMinMatch) {
    var hours = ruMatch ? parseInt(ruMatch[1]) : 0;
    var mins = ruMinMatch ? parseInt(ruMinMatch[1]) : 0;
    return hours * 60 + mins;
  }

  // Plain number string — treat as hours (e.g. "1.5" = 90 min)
  var num = parseFloat(str);
  if (!isNaN(num)) return Math.round(num * 60);

  return 0;
}

/**
 * Formats a date as DD.MM.YYYY.
 */
function formatDate(d) {
  var day = ('0' + d.getDate()).slice(-2);
  var month = ('0' + (d.getMonth() + 1)).slice(-2);
  return day + '.' + month + '.' + d.getFullYear();
}

/**
 * Formats a date as HH:MM.
 */
function formatTime(d) {
  var hh = ('0' + d.getHours()).slice(-2);
  var mm = ('0' + d.getMinutes()).slice(-2);
  return hh + ':' + mm;
}
