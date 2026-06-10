/**
 * ═══════════════════════════════════════════════════════════════
 * MEDICAL SRS v3 → SHEETS SUPERPOWERS + GOOGLE CALENDAR SYNC
 * ═══════════════════════════════════════════════════════════════
 *
 * v3 ADDS:
 * ✅ Command Palette sidebar — fuzzy-search lessons / modules / commands,
 *    log a review by pressing 0–5, filter any module in one keystroke
 * ✅ Real module filtering (sets the sheet's filter for you)
 * ✅ Daily progress snapshot → History sheet (feeds the Dashboard
 *    "Progress over time" chart)
 * ✅ Single INTERVALS constant (must mirror column H in the sheet)
 * (v2 kept: graceful recovery, workload cap, smart reschedule,
 *  color-coded events, auto-status, batch ops)
 *
 * SETUP (one time):
 *   1. Run setupTriggers() → hourly sync + on-edit sync + daily snapshot.
 *   2. Palette keyboard shortcut ("Ctrl+K", the Sheets way):
 *      Extensions → Macros → Import macro → openCommandPalette,
 *      then Extensions → Macros → Manage macros → assign it number 1.
 *      Now  Ctrl+Alt+Shift+1  opens the palette. (Sheets reserves the real
 *      Ctrl+K; inside the palette, Ctrl+K refocuses the search box.)
 *
 * COLUMN MAP (0-indexed) — MUST match srs-system.xlsx exactly:
 *   A=# | B=Semester | C=Module | D=Subject | E=Lesson | F=Last Review
 *   G=Mastery | H=Interval | I=Next Review | J=Status | K=Priority
 *   L=Synced | M=Event ID | (N=Notes — user-optional, script ignores it)
 */

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const CONFIG = {
  SHEET_NAME: 'lesson-database',
  HEADER_ROWS: 2,  // Data starts at row 3

  // Column indices (0-based from data range)
  COL: {
    NUM:         0,   // A - row number
    SEMESTER:    1,   // B
    MODULE:      2,   // C
    SUBJECT:     3,   // D
    LESSON:      4,   // E
    LAST_REVIEW: 5,   // F — USER INPUT
    MASTERY:     6,   // G — USER INPUT
    INTERVAL:    7,   // H — formula
    NEXT_REVIEW: 8,   // I — formula
    STATUS:      9,   // J — formula
    PRIORITY:   10,   // K — formula
    SYNCED:     11,   // L — script-managed
    EVENT_ID:   12,   // M — script-managed (hidden column)
  },

  // Calendar
  CALENDAR_PREFIX: '📚 ',
  SYNC_FUTURE_DAYS: 90,

  // Workload management
  MAX_REVIEWS_PER_DAY: 15,   // Cap: don't schedule more than this per day
  SPREAD_OVERDUE_DAYS: 3,    // Spread overdue items across this many days

  // Mastery → color mapping for calendar events
  MASTERY_COLORS: {
    0: CalendarApp.EventColor.RED,       // Forgot
    1: CalendarApp.EventColor.ORANGE,    // Hard
    2: CalendarApp.EventColor.YELLOW,    // Medium
    3: CalendarApp.EventColor.CYAN,      // Easy
    4: CalendarApp.EventColor.GREEN,     // Confident
    5: CalendarApp.EventColor.GRAPE,     // Mastered
  },

  // Auto-sync interval
  AUTO_SYNC_HOURS: 1,

  // Daily progress snapshot
  HISTORY_SHEET: 'History',
  SNAPSHOT_HOUR: 22,   // 22:00 local time
};

// Mastery → days until next review.
// ⚠️ MUST stay identical to the column-H formula in srs-system.xlsx.
const INTERVALS = [1, 3, 7, 14, 30, 60];

// ═══════════════════════════════════════════════════════════════
// MAIN SYNC
// ═══════════════════════════════════════════════════════════════

function syncToCalendar() {
  const t0 = Date.now();
  const calendar = CalendarApp.getDefaultCalendar();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log('Sheet not found: ' + CONFIG.SHEET_NAME);
    return;
  }

  // Ensure hidden Event ID column exists
  ensureEventIdColumn_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROWS) return;

  const numRows = lastRow - CONFIG.HEADER_ROWS;
  const numCols = CONFIG.COL.EVENT_ID + 1;  // Through column L
  const range = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1, numRows, numCols);
  const data = range.getValues();

  const stats = { created: 0, updated: 0, deleted: 0, skipped: 0 };
  const batchUpdates = [];  // [{row, synced, eventId}]

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + CONFIG.HEADER_ROWS + 1;

    const lesson = row[CONFIG.COL.LESSON];
    const subject = row[CONFIG.COL.SUBJECT];
    const module = row[CONFIG.COL.MODULE];
    const semester = row[CONFIG.COL.SEMESTER];
    if (!lesson || !module) continue;

    const nextReview = parseDate_(row[CONFIG.COL.NEXT_REVIEW]);
    const mastery = parseInt(row[CONFIG.COL.MASTERY]) || 0;
    const lastReview = parseDate_(row[CONFIG.COL.LAST_REVIEW]);
    const eventId = String(row[CONFIG.COL.EVENT_ID] || '').trim();
    const synced = String(row[CONFIG.COL.SYNCED] || '').trim();

    const hasEvent = eventId.length > 5;
    const hasValidDate = nextReview !== null;
    const isMastered = mastery >= 5;

    // DECISION LOGIC
    if (!hasValidDate && hasEvent) {
      // Date removed or invalid → delete event
      if (deleteEvent_(calendar, eventId)) {
        batchUpdates.push({ row: rowNum, synced: '', eventId: '' });
        stats.deleted++;
      }
    } else if (hasValidDate && !isMastered) {
      const title = buildTitle_(semester, module, lesson);
      const desc = buildDescription_(semester, module, subject, lesson, mastery, lastReview, nextReview);
      const color = CONFIG.MASTERY_COLORS[Math.min(mastery, 5)];

      if (hasEvent) {
        // Update existing
        if (updateEvent_(calendar, eventId, title, desc, nextReview, color)) {
          batchUpdates.push({ row: rowNum, synced: '✅', eventId: eventId });
          stats.updated++;
        } else {
          // Event was deleted externally → recreate
          const newId = createEvent_(calendar, title, desc, nextReview, color);
          if (newId) {
            batchUpdates.push({ row: rowNum, synced: '✅', eventId: newId });
            stats.created++;
          }
        }
      } else {
        // Create new
        const newId = createEvent_(calendar, title, desc, nextReview, color);
        if (newId) {
          batchUpdates.push({ row: rowNum, synced: '✅', eventId: newId });
          stats.created++;
        }
      }
    } else if (isMastered && hasEvent) {
      // Mastered → remove from calendar (no need to review)
      deleteEvent_(calendar, eventId);
      batchUpdates.push({ row: rowNum, synced: '✅ done', eventId: '' });
      stats.deleted++;
    } else {
      stats.skipped++;
    }
  }

  // BATCH WRITE (much faster than cell-by-cell)
  applyBatchUpdates_(sheet, batchUpdates);

  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  const msg = `Sync done (${elapsed}s): +${stats.created} ~${stats.updated} -${stats.deleted} =${stats.skipped} skipped`;
  Logger.log(msg);

  try {
    if (stats.created + stats.updated + stats.deleted > 0) {
      SpreadsheetApp.getUi().alert('✅ ' + msg);
    } else {
      SpreadsheetApp.getUi().alert('✅ Already up to date.');
    }
  } catch (e) {
    // Silent if triggered automatically
  }
}

// ═══════════════════════════════════════════════════════════════
// CALENDAR OPERATIONS
// ═══════════════════════════════════════════════════════════════

function buildTitle_(semester, module, lesson) {
  // Format: "📚 s3 - Sémiologie I | Diarrhée aiguë"
  return CONFIG.CALENDAR_PREFIX + String(semester).toLowerCase() + ' - ' + module + ' | ' + lesson;
}

function buildDescription_(semester, module, subject, lesson, mastery, lastReview, nextReview) {
  const stars = '⭐'.repeat(Math.min(mastery, 5)) + '☆'.repeat(Math.max(0, 5 - mastery));
  const lr = lastReview ? Utilities.formatDate(lastReview, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'never';
  return [
    `📚 Spaced Repetition Review`,
    ``,
    `Semester: ${semester}`,
    `Module: ${module}`,
    subject ? `Subject: ${subject}` : null,
    `Lesson: ${lesson}`,
    `Mastery: ${stars} (${mastery}/5)`,
    `Last reviewed: ${lr}`,
    ``,
    `After studying, update:`,
    `  • Last Review → today's date (Ctrl+;)`,
    `  • Mastery → 0-5 based on recall`,
    ``,
    `🔗 Auto-managed by Medical SRS`,
  ].filter(line => line !== null).join('\n');
}

function createEvent_(calendar, title, description, date, color) {
  try {
    const event = calendar.createAllDayEvent(title, date, { description });
    event.setColor(color);
    event.removeAllReminders();
    event.addPopupReminder(480);  // 8 hours before (morning reminder)
    return event.getId();
  } catch (e) {
    Logger.log('Create failed: ' + e);
    return null;
  }
}

function updateEvent_(calendar, eventId, title, description, newDate, color) {
  try {
    const event = calendar.getEventById(eventId);
    if (!event) return false;

    const currentDate = event.getAllDayStartDate();
    currentDate.setHours(0, 0, 0, 0);
    newDate.setHours(0, 0, 0, 0);

    // Only modify if something changed
    if (currentDate.getTime() !== newDate.getTime() || event.getTitle() !== title) {
      event.setTitle(title);
      event.setAllDayDate(newDate);
      event.setDescription(description);
      event.setColor(color);
    }
    return true;
  } catch (e) {
    Logger.log('Update failed for ' + eventId + ': ' + e);
    return false;
  }
}

function deleteEvent_(calendar, eventId) {
  try {
    const event = calendar.getEventById(eventId);
    if (event) event.deleteEvent();
    return true;
  } catch (e) {
    Logger.log('Delete failed: ' + e);
    return true; // Treat as success if already gone
  }
}

// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════

function parseDate_(value) {
  if (!value) return null;
  try {
    const d = new Date(value);
    if (isNaN(d.getTime()) || d.getFullYear() < 2025 || d.getFullYear() > 2100) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  } catch (e) {
    return null;
  }
}

function ensureEventIdColumn_(sheet) {
  // Column L (12) = Event ID. Add header if missing.
  const headerRow = CONFIG.HEADER_ROWS;
  const cell = sheet.getRange(headerRow, CONFIG.COL.EVENT_ID + 1);
  if (cell.getValue() !== 'Event ID') {
    cell.setValue('Event ID');
    cell.setFontSize(8).setFontColor('#999999');
    // Hide the column so it doesn't clutter the view
    sheet.hideColumns(CONFIG.COL.EVENT_ID + 1);
  }
}

function applyBatchUpdates_(sheet, updates) {
  if (updates.length === 0) return;
  
  // Group updates for efficiency
  const syncCol = CONFIG.COL.SYNCED + 1;     // Column K
  const eventCol = CONFIG.COL.EVENT_ID + 1;   // Column L

  updates.forEach(u => {
    sheet.getRange(u.row, syncCol).setValue(u.synced);
    sheet.getRange(u.row, eventCol).setValue(u.eventId);
  });
}

// ═══════════════════════════════════════════════════════════════
// SMART RESCHEDULE: Spread overdue items across multiple days
// ═══════════════════════════════════════════════════════════════

/**
 * Run this when you've been away for days and have 30+ overdue items.
 * Instead of piling everything on today, it spreads reviews across
 * the next few days with a daily cap.
 */
function smartReschedule() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🔄 Smart Reschedule',
    'This will spread your overdue reviews across the next ' +
    CONFIG.SPREAD_OVERDUE_DAYS + ' days (max ' + CONFIG.MAX_REVIEWS_PER_DAY +
    '/day).\n\nOverdue items will be reassigned to upcoming days.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROWS) return;

  const numRows = lastRow - CONFIG.HEADER_ROWS;
  const range = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1, numRows, CONFIG.COL.EVENT_ID + 1);
  const data = range.getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Find all overdue items (sorted by oldest first)
  const overdue = [];
  for (let i = 0; i < data.length; i++) {
    const nextReview = parseDate_(data[i][CONFIG.COL.NEXT_REVIEW]);
    const mastery = parseInt(data[i][CONFIG.COL.MASTERY]) || 0;
    if (nextReview && nextReview < today && mastery < 5) {
      overdue.push({
        index: i,
        rowNum: i + CONFIG.HEADER_ROWS + 1,
        nextReview: nextReview,
        mastery: mastery,
        daysOverdue: Math.floor((today - nextReview) / 86400000),
      });
    }
  }

  if (overdue.length === 0) {
    ui.alert('✅ No overdue items! You\'re on track.');
    return;
  }

  // Sort: most overdue + lowest mastery first
  overdue.sort((a, b) => {
    if (a.mastery !== b.mastery) return a.mastery - b.mastery;
    return b.daysOverdue - a.daysOverdue;
  });

  // Distribute across days
  const perDay = Math.min(CONFIG.MAX_REVIEWS_PER_DAY,
    Math.ceil(overdue.length / CONFIG.SPREAD_OVERDUE_DAYS));

  let rescheduled = 0;
  for (let i = 0; i < overdue.length; i++) {
    const dayOffset = Math.floor(i / perDay);
    const newDate = new Date(today);
    newDate.setDate(newDate.getDate() + dayOffset);

    // Update the Last Review to trigger recalculation
    // We set Next Review date by adjusting the "Last Review" so the formula
    // Last Review + Interval = target date
    // Actually, we just need to clear synced so the event gets updated
    const item = overdue[i];
    const nextRevCell = sheet.getRange(item.rowNum, CONFIG.COL.NEXT_REVIEW + 1);

    // Directly set the next review date (override formula temporarily)
    // Better approach: set last_review = newDate - interval so formula calculates correctly
    const interval = INTERVALS[item.mastery] || 1;
    const fakeLastReview = new Date(newDate);
    fakeLastReview.setDate(fakeLastReview.getDate() - interval);

    sheet.getRange(item.rowNum, CONFIG.COL.LAST_REVIEW + 1).setValue(fakeLastReview);
    sheet.getRange(item.rowNum, CONFIG.COL.SYNCED + 1).setValue('');  // Force re-sync
    rescheduled++;
  }

  ui.alert(
    `✅ Rescheduled ${rescheduled} overdue items!\n\n` +
    `Spread across ${Math.min(CONFIG.SPREAD_OVERDUE_DAYS, Math.ceil(rescheduled / perDay))} days\n` +
    `(~${perDay} reviews/day)\n\n` +
    'Run Sync Now to update your calendar.'
  );
}

// ═══════════════════════════════════════════════════════════════
// DAILY DIGEST: Shows what to study today
// ═══════════════════════════════════════════════════════════════

function showDailyDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROWS) return;

  const data = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1, lastRow - CONFIG.HEADER_ROWS, CONFIG.COL.PRIORITY + 1).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const overdue = [];
  const dueToday = [];
  const dueTomorrow = [];

  for (const row of data) {
    const nextReview = parseDate_(row[CONFIG.COL.NEXT_REVIEW]);
    const mastery = parseInt(row[CONFIG.COL.MASTERY]) || 0;
    if (!nextReview || mastery >= 5) continue;

    const diff = Math.floor((nextReview - today) / 86400000);
    const item = `  • [M${mastery}] ${row[CONFIG.COL.SEMESTER]} ${row[CONFIG.COL.MODULE]} | ${row[CONFIG.COL.LESSON]}`;

    if (diff < 0) overdue.push(item);
    else if (diff === 0) dueToday.push(item);
    else if (diff === 1) dueTomorrow.push(item);
  }

  const msg = [
    `📅 ${Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMM dd')}`,
    '',
    `🔴 OVERDUE (${overdue.length}):`,
    overdue.length ? overdue.join('\n') : '  None! 🎉',
    '',
    `🟢 TODAY (${dueToday.length}):`,
    dueToday.length ? dueToday.join('\n') : '  None scheduled',
    '',
    `🔵 TOMORROW (${dueTomorrow.length}):`,
    dueTomorrow.length ? dueTomorrow.join('\n') : '  None scheduled',
    '',
    `⏱️ Est. time: ~${(overdue.length + dueToday.length) * 10} min`,
  ].join('\n');

  SpreadsheetApp.getUi().alert(msg);
}

// ═══════════════════════════════════════════════════════════════
// SETUP & TRIGGERS
// ═══════════════════════════════════════════════════════════════

function setupTriggers() {
  deleteTriggers_();

  ScriptApp.newTrigger('syncToCalendar')
    .timeBased()
    .everyHours(CONFIG.AUTO_SYNC_HOURS)
    .create();

  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  ScriptApp.newTrigger('logDailySnapshot')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.SNAPSHOT_HOUR)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Auto-Sync Enabled!\n\n' +
    `• Syncs every ${CONFIG.AUTO_SYNC_HOURS}h automatically\n` +
    '• Syncs when you edit Last Review or Mastery\n' +
    '• Calendar events are color-coded by mastery\n' +
    `• Daily progress snapshot at ${CONFIG.SNAPSHOT_HOUR}:00 → History sheet\n\n` +
    '💡 Palette shortcut: Extensions → Macros → Import macro →\n' +
    'openCommandPalette → assign number 1 → Ctrl+Alt+Shift+1.\n\n' +
    'Running first sync now...'
  );

  syncToCalendar();
  logDailySnapshot();
}

function deleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (['syncToCalendar', 'onEditTrigger', 'logDailySnapshot'].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function disableAutoSync() {
  deleteTriggers_();
  SpreadsheetApp.getUi().alert('✅ Auto-sync disabled.');
}

function onEditTrigger(e) {
  if (!e) return;
  const sheetName = e.range.getSheet().getName();
  if (sheetName !== CONFIG.SHEET_NAME) return;

  const col = e.range.getColumn();
  // Only sync on Last Review or Mastery edits (1-based columns)
  if (col === CONFIG.COL.LAST_REVIEW + 1 || col === CONFIG.COL.MASTERY + 1) {
    Utilities.sleep(1500);
    syncToCalendar();
  }
}

// ═══════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📚 SRS')
    .addItem('🎛 Command Palette', 'openCommandPalette')
    .addItem('📋 Daily Digest', 'showDailyDigest')
    .addItem('🔄 Sync Calendar Now', 'syncToCalendar')
    .addSeparator()
    .addItem('🧠 Smart Reschedule (overdue)', 'smartReschedule')
    .addItem('📸 Log Progress Snapshot', 'logDailySnapshot')
    .addItem('🧹 Clear Module Filter', 'clearModuleFilter')
    .addSeparator()
    .addItem('⚙️ Setup Auto-Sync', 'setupTriggers')
    .addItem('🛑 Disable Auto-Sync', 'disableAutoSync')
    .addSeparator()
    .addItem('🧽 Clear Sync Markers', 'clearSyncMarkers')
    .addItem('🗑️ Delete All SRS Events', 'deleteAllSRSEvents')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════
// MAINTENANCE
// ═══════════════════════════════════════════════════════════════

// Note: menu items can't call underscore-"private" functions, hence no trailing _.
function clearSyncMarkers() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear all sync markers?', 'Events in calendar will remain.\nNext sync recreates all.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow > CONFIG.HEADER_ROWS) {
    const numRows = lastRow - CONFIG.HEADER_ROWS;
    sheet.getRange(CONFIG.HEADER_ROWS + 1, CONFIG.COL.SYNCED + 1, numRows, 1).clearContent();
    sheet.getRange(CONFIG.HEADER_ROWS + 1, CONFIG.COL.EVENT_ID + 1, numRows, 1).clearContent();
  }
  ui.alert('✅ Cleared.');
}

function deleteAllSRSEvents() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Delete ALL SRS events from calendar?', 'Cannot be undone!', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const cal = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const future = new Date(now);
  future.setFullYear(future.getFullYear() + 1);

  let count = 0;
  cal.getEvents(new Date(2025, 0, 1), future).forEach(ev => {
    if (ev.getTitle().startsWith(CONFIG.CALENDAR_PREFIX)) {
      ev.deleteEvent();
      count++;
    }
  });

  ui.alert(`✅ Deleted ${count} SRS events.`);
}

// ═══════════════════════════════════════════════════════════════
// COMMAND PALETTE  (the "Ctrl+K" of this system)
// Open: 📚 SRS menu → Command Palette, or Ctrl+Alt+Shift+1 after the
// one-time macro import described in the header.
// ═══════════════════════════════════════════════════════════════

function openCommandPalette() {
  const html = HtmlService.createHtmlOutput(PALETTE_HTML_)
    .setTitle('📚 SRS Command Palette');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Everything the palette can search: commands, modules, all lessons. */
function getPaletteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lessons = [];
  const moduleSeen = {};
  const modules = [];
  const lastRow = sheet ? sheet.getLastRow() : 0;
  if (sheet && lastRow > CONFIG.HEADER_ROWS) {
    const data = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1,
      lastRow - CONFIG.HEADER_ROWS, CONFIG.COL.PRIORITY + 1).getValues();
    for (let i = 0; i < data.length; i++) {
      const sem = data[i][CONFIG.COL.SEMESTER];
      const mod = data[i][CONFIG.COL.MODULE];
      const les = data[i][CONFIG.COL.LESSON];
      if (!les || !mod) continue;
      const mastery = data[i][CONFIG.COL.MASTERY];
      lessons.push({
        r: i + CONFIG.HEADER_ROWS + 1,
        t: sem + ' · ' + mod + ' | ' + les,
        m: (mastery === '' || mastery === null) ? null : Number(mastery),
      });
      const key = sem + ' — ' + mod;
      if (!moduleSeen[key]) {
        moduleSeen[key] = true;
        modules.push({ label: key, sem: String(sem), mod: String(mod) });
      }
    }
  }
  const commands = [
    { id: 'sync',        label: '🔄 Sync calendar now' },
    { id: 'digest',      label: '📋 Daily digest' },
    { id: 'reschedule',  label: '🧠 Smart reschedule overdue' },
    { id: 'clearFilter', label: '🧹 Clear module filter' },
    { id: 'today',       label: '⚡ Go to Today tab' },
    { id: 'dashboard',   label: '📊 Go to Dashboard' },
    { id: 'moduleView',  label: '🔍 Go to Module View' },
    { id: 'database',    label: '🗂 Go to lesson-database' },
    { id: 'snapshot',    label: '📸 Log progress snapshot now' },
  ];
  return { commands: commands, modules: modules, lessons: lessons };
}

function paletteCommand(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const go = name => { const s = ss.getSheetByName(name); if (s) ss.setActiveSheet(s); return name; };
  switch (id) {
    case 'sync':        syncToCalendar();      return 'Calendar synced ✓';
    case 'digest':      showDailyDigest();     return 'Digest opened';
    case 'reschedule':  smartReschedule();     return 'Done';
    case 'clearFilter': clearModuleFilter();   return 'Filters cleared ✓';
    case 'today':       return go('Today');
    case 'dashboard':   return go('Dashboard');
    case 'moduleView':  return go('Module View');
    case 'database':    return go(CONFIG.SHEET_NAME);
    case 'snapshot':    logDailySnapshot();    return 'Snapshot logged ✓';
  }
  return 'Unknown command: ' + id;
}

/** Filter lesson-database to one module (sets the real sheet filter). */
function filterModuleFromPalette(sem, mod) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = Math.max(sheet.getLastRow(), CONFIG.HEADER_ROWS + 1);
  let filter = sheet.getFilter();
  if (!filter) {
    filter = sheet.getRange(CONFIG.HEADER_ROWS, 1,
      lastRow - CONFIG.HEADER_ROWS + 1, CONFIG.COL.PRIORITY + 1).createFilter();
  }
  filter.setColumnFilterCriteria(CONFIG.COL.SEMESTER + 1,
    SpreadsheetApp.newFilterCriteria().whenTextEqualTo(sem).build());
  filter.setColumnFilterCriteria(CONFIG.COL.MODULE + 1,
    SpreadsheetApp.newFilterCriteria().whenTextEqualTo(mod).build());
  // Mirror the choice in Module View's dropdown, then show the filtered list.
  const mv = ss.getSheetByName('Module View');
  if (mv) mv.getRange('C3').setValue(sem + ' — ' + mod);
  ss.setActiveSheet(sheet);
  sheet.setActiveSelection('A' + (CONFIG.HEADER_ROWS + 1));
  return 'Filtered: ' + sem + ' — ' + mod;
}

function clearModuleFilter() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const f = sheet.getFilter();
  if (f) {
    f.setColumnFilterCriteria(CONFIG.COL.SEMESTER + 1, null);
    f.setColumnFilterCriteria(CONFIG.COL.MODULE + 1, null);
    f.setColumnFilterCriteria(CONFIG.COL.SUBJECT + 1, null);
  }
}

/** One-keystroke review logging from the palette: F=today, G=mastery. */
function reviewLessonFromPalette(row, mastery) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  row = parseInt(row, 10); mastery = parseInt(mastery, 10);
  if (isNaN(row) || row <= CONFIG.HEADER_ROWS || row > sheet.getLastRow()) return '⚠️ Invalid row';
  if (isNaN(mastery) || mastery < 0 || mastery > 5) return '⚠️ Mastery must be 0–5';
  const today = new Date(); today.setHours(0, 0, 0, 0);
  sheet.getRange(row, CONFIG.COL.LAST_REVIEW + 1).setValue(today);
  sheet.getRange(row, CONFIG.COL.MASTERY + 1).setValue(mastery);
  const les = sheet.getRange(row, CONFIG.COL.LESSON + 1).getValue();
  // Calendar catches up on the hourly sync (or run Sync now from the palette).
  return '✓ ' + les + ' → M' + mastery +
    (mastery >= 5 ? ' · mastered 🎉' : ' · back in ' + INTERVALS[mastery] + 'd');
}

// ═══════════════════════════════════════════════════════════════
// DAILY PROGRESS SNAPSHOT → History sheet (feeds Dashboard chart)
// ═══════════════════════════════════════════════════════════════

function logDailySnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!h) {
    h = ss.insertSheet(CONFIG.HISTORY_SHEET);
    h.appendRow(['Date', 'Total', 'Started', 'Mastered', 'Overdue', 'Reviewed that day']);
  }
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let total = 0, started = 0, mastered = 0, overdue = 0, reviewedToday = 0;
  if (lastRow > CONFIG.HEADER_ROWS) {
    const data = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1,
      lastRow - CONFIG.HEADER_ROWS, CONFIG.COL.PRIORITY + 1).getValues();
    for (const row of data) {
      if (!row[CONFIG.COL.LESSON]) continue;
      total++;
      const lr = parseDate_(row[CONFIG.COL.LAST_REVIEW]);
      const m = parseInt(row[CONFIG.COL.MASTERY], 10);
      const mast = isNaN(m) ? null : m;
      if (lr) started++;
      if (mast !== null && mast >= 5) mastered++;
      const nr = parseDate_(row[CONFIG.COL.NEXT_REVIEW]);
      if (lr && nr && nr < today && (mast === null || mast < 5)) overdue++;
      if (lr && lr.getTime() === today.getTime()) reviewedToday++;
    }
  }
  // Upsert: overwrite today's row if it already exists (re-runs are safe).
  const hLast = h.getLastRow();
  let target = hLast + 1;
  if (hLast >= 2) {
    const lastDate = parseDate_(h.getRange(hLast, 1).getValue());
    if (lastDate && lastDate.getTime() === today.getTime()) target = hLast;
  }
  h.getRange(target, 1, 1, 6).setValues([[today, total, started, mastered, overdue, reviewedToday]]);
  h.getRange(target, 1).setNumberFormat('yyyy-mm-dd');
  Logger.log('Snapshot: ' + [total, started, mastered, overdue, reviewedToday].join('/'));
}

// ═══════════════════════════════════════════════════════════════
// PALETTE UI (HTML kept inline so setup stays "paste ONE file")
// ═══════════════════════════════════════════════════════════════

const PALETTE_HTML_ = `<!DOCTYPE html><html><head><base target="_top"><style>
*{box-sizing:border-box;font-family:'Segoe UI',Roboto,Arial,sans-serif}
body{margin:0;padding:10px;background:#1e1e2e;color:#eee}
#q{width:100%;padding:10px;font-size:14px;border:1px solid #444;border-radius:8px;background:#2a2a3c;color:#fff;outline:none}
#q:focus{border-color:#7aa2f7}
.hint{color:#888;font-size:11px;margin:6px 2px}
#list{margin-top:8px;max-height:72vh;overflow-y:auto}
.item{padding:7px 9px;border-radius:6px;cursor:pointer;font-size:13px;line-height:1.35}
.item.sel{background:#33415e}
.meta{color:#9aa;font-size:11px}
.badge{display:inline-block;min-width:16px;text-align:center;border-radius:4px;background:#444;font-size:11px;padding:1px 4px;margin-left:6px}
.mast{display:none;margin-top:5px}
.item.sel .mast{display:block}
.mast button{margin:1px;padding:4px 10px;border:none;border-radius:5px;cursor:pointer;font-weight:bold;color:#fff}
#toast{position:fixed;bottom:8px;left:10px;right:10px;background:#2e7d32;color:#fff;padding:8px;border-radius:6px;display:none;font-size:13px;z-index:9}
</style></head><body>
<input id="q" placeholder="Type a lesson, module or command…" autocomplete="off">
<div class="hint">↑↓ move · Enter run · lesson: Enter then <b>0–5</b> logs today's review · Esc clear · Ctrl+K focus</div>
<div id="list"></div><div id="toast"></div>
<script>
let DATA={commands:[],modules:[],lessons:[]},RES=[],SEL=0,ARMED=null;
const q=document.getElementById('q'),list=document.getElementById('list'),toast=document.getElementById('toast');
const COLORS=['#d9534f','#e8833a','#c9a91e','#9ac34a','#5cb85c','#8e6bbf'];
function norm(s){return s.toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')}
function search(){
  ARMED=null;
  const t=norm(q.value.trim());
  const hits=[];
  const push=o=>{if(hits.length<60)hits.push(o)};
  if(!t){
    DATA.commands.forEach(c=>push({k:'c',o:c}));
    DATA.modules.slice(0,12).forEach(m=>push({k:'m',o:m}));
  }else{
    const words=t.split(/\\s+/);
    const match=s=>{const n=norm(s);return words.every(w=>n.includes(w))};
    DATA.commands.forEach(c=>{if(match(c.label))push({k:'c',o:c})});
    DATA.modules.forEach(m=>{if(match(m.label))push({k:'m',o:m})});
    for(const l of DATA.lessons){if(hits.length>=60)break;if(match(l.t))push({k:'l',o:l});}
  }
  RES=hits;SEL=0;render();
}
function render(){
  list.innerHTML='';
  RES.forEach((h,i)=>{
    const d=document.createElement('div');
    d.className='item'+(i===SEL?' sel':'');
    if(h.k==='c')d.innerHTML=h.o.label;
    else if(h.k==='m')d.innerHTML='📂 '+h.o.label+' <span class="meta">filter this module</span>';
    else d.innerHTML='📖 '+h.o.t+(h.o.m!==null?' <span class="badge">M'+h.o.m+'</span>':' <span class="badge">new</span>')+
      '<div class="mast">'+[0,1,2,3,4,5].map(n=>'<button style="background:'+COLORS[n]+'" onclick="review('+h.o.r+','+n+');event.stopPropagation()">'+n+'</button>').join('')+'</div>';
    d.onclick=()=>{SEL=i;render();run();};
    list.appendChild(d);
  });
}
function run(){
  const h=RES[SEL];if(!h)return;
  if(h.k==='c')gs('paletteCommand',h.o.id);
  else if(h.k==='m')gs('filterModuleFromPalette',h.o.sem,h.o.mod);
  else{ARMED=h.o.r;render();say('Press 0–5 → mastery for: '+h.o.t,'#555');}
}
function review(row,m){ARMED=null;gs('reviewLessonFromPalette',row,m);}
function gs(fn){
  const args=Array.prototype.slice.call(arguments,1);
  say('⏳ working…','#555');
  const runner=google.script.run
    .withSuccessHandler(msg=>{say(msg||'Done ✓','#2e7d32');if(fn==='reviewLessonFromPalette')refresh();})
    .withFailureHandler(e=>say('⚠️ '+e.message,'#b23b3b'));
  runner[fn].apply(runner,args);
}
function say(t,bg){toast.textContent=t;toast.style.background=bg;toast.style.display='block';
  clearTimeout(say.t);say.t=setTimeout(()=>toast.style.display='none',4000);}
function refresh(){google.script.run.withSuccessHandler(d=>{DATA=d;search();}).getPaletteData();}
q.addEventListener('input',search);
document.addEventListener('keydown',e=>{
  if(ARMED!==null&&/^[0-5]$/.test(e.key)){e.preventDefault();review(ARMED,+e.key);return;}
  if(e.key==='ArrowDown'){SEL=Math.min(SEL+1,RES.length-1);render();e.preventDefault();}
  else if(e.key==='ArrowUp'){SEL=Math.max(SEL-1,0);render();e.preventDefault();}
  else if(e.key==='Enter'){run();e.preventDefault();}
  else if(e.key==='Escape'){ARMED=null;q.value='';search();q.focus();}
  else if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){q.focus();q.select();e.preventDefault();}
});
refresh();q.focus();
</script></body></html>`;
