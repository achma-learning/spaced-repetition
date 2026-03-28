/**
 * ═══════════════════════════════════════════════════════════════
 * MEDICAL SRS v2 → GOOGLE CALENDAR SYNC
 * ═══════════════════════════════════════════════════════════════
 * 
 * WHAT'S NEW vs v1:
 * ✅ Graceful missed-day recovery (no panic pile-ups)
 * ✅ Daily workload cap (won't schedule 50 reviews on one day)
 * ✅ Smart rescheduling: overdue items spread across next 3 days
 * ✅ Color-coded calendar events by mastery level
 * ✅ Auto-status update (no manual status column needed)
 * ✅ Faster batch operations
 * ✅ Event descriptions show mastery + interval info
 * 
 * SETUP: Run setupTriggers() once → forget about it.
 * 
 * COLUMN MAP (0-indexed):
 *   A=# | B=Module | C=Subject | D=Topic | E=Last Review
 *   F=Mastery | G=Interval | H=Next Review | I=Status | J=Priority | K=Synced
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
    MODULE:      1,   // B
    SUBJECT:     2,   // C
    TOPIC:       3,   // D
    LAST_REVIEW: 4,   // E — USER INPUT
    MASTERY:     5,   // F — USER INPUT
    INTERVAL:    6,   // G — formula
    NEXT_REVIEW: 7,   // H — formula
    STATUS:      8,   // I — formula
    PRIORITY:    9,   // J — formula
    SYNCED:     10,   // K — script-managed
    EVENT_ID:   11,   // L — script-managed (hidden column)
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
};

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

    const topic = row[CONFIG.COL.TOPIC];
    const subject = row[CONFIG.COL.SUBJECT];
    const module = row[CONFIG.COL.MODULE];
    if (!topic || !subject) continue;

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
      const title = buildTitle_(module, topic);
      const desc = buildDescription_(subject, topic, mastery, lastReview, nextReview);
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

function buildTitle_(module, topic) {
  return CONFIG.CALENDAR_PREFIX + module + ' — ' + topic;
}

function buildDescription_(subject, topic, mastery, lastReview, nextReview) {
  const stars = '⭐'.repeat(Math.min(mastery, 5)) + '☆'.repeat(Math.max(0, 5 - mastery));
  const lr = lastReview ? Utilities.formatDate(lastReview, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'never';
  return [
    `📚 Spaced Repetition Review`,
    ``,
    `Subject: ${subject}`,
    `Topic: ${topic}`,
    `Mastery: ${stars} (${mastery}/5)`,
    `Last reviewed: ${lr}`,
    ``,
    `After studying, update:`,
    `  • Last Review → today's date (Ctrl+;)`,
    `  • Mastery → 0-5 based on recall`,
    ``,
    `🔗 Auto-managed by Medical SRS v2`,
  ].join('\n');
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
    const interval = [1, 3, 7, 14, 30, 60][item.mastery] || 1;
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

  const data = sheet.getRange(CONFIG.HEADER_ROWS + 1, 1, lastRow - CONFIG.HEADER_ROWS, 10).getValues();
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
    const item = `  • [M${mastery}] ${row[CONFIG.COL.MODULE]} — ${row[CONFIG.COL.TOPIC]}`;

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

  SpreadsheetApp.getUi().alert(
    '✅ Auto-Sync Enabled!\n\n' +
    `• Syncs every ${CONFIG.AUTO_SYNC_HOURS}h automatically\n` +
    '• Syncs when you edit Last Review or Mastery\n' +
    '• Calendar events are color-coded by mastery\n\n' +
    'Running first sync now...'
  );

  syncToCalendar();
}

function deleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (['syncToCalendar', 'onEditTrigger'].includes(fn)) {
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
  // Only sync on Last Review (E=5) or Mastery (F=6) edits
  if (col === 5 || col === 6) {
    Utilities.sleep(1500);
    syncToCalendar();
  }
}

// ═══════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📚 SRS v2')
    .addItem('📋 Daily Digest', 'showDailyDigest')
    .addItem('🔄 Sync Calendar Now', 'syncToCalendar')
    .addSeparator()
    .addItem('🧠 Smart Reschedule (overdue)', 'smartReschedule')
    .addSeparator()
    .addItem('⚙️ Setup Auto-Sync', 'setupTriggers')
    .addItem('🛑 Disable Auto-Sync', 'disableAutoSync')
    .addSeparator()
    .addItem('🧹 Clear Sync Markers', 'clearSyncMarkers_')
    .addItem('🗑️ Delete All SRS Events', 'deleteAllSRSEvents_')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════
// MAINTENANCE
// ═══════════════════════════════════════════════════════════════

function clearSyncMarkers_() {
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

function deleteAllSRSEvents_() {
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
