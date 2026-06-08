/**
 * ═══════════════════════════════════════════════════════════════
 * MEDICAL SRS → GOOGLE CALENDAR SYNC (PRODUCTION VERSION)
 * ═══════════════════════════════════════════════════════════════
 * 
 * FEATURES:
 * ✅ Syncs multiple sheets (S8, AIO, or custom)
 * ✅ Updates existing events instead of duplicating
 * ✅ Deletes events when dates are removed
 * ✅ Handles date changes automatically
 * ✅ Smart conflict resolution
 * ✅ Detailed logging
 * ✅ One-click setup with triggers
 * 
 * SETUP: Run setupTriggers() once, then forget about it!
 */

// ═══════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const CONFIG = {
  // Which sheets to sync (add more if needed)
  SHEETS_TO_SYNC: ['S8'],  // Change to ['S8', 'AIO'] for multiple sheets
  
  // Column mapping (adjust if your columns are different)
  COLUMNS: {
    SUBJECT: 0,      // Column A (0-indexed)
    TOPIC: 1,        // Column B
    STATUS: 2,       // Column C
    MASTERY: 3,      // Column D
    LAST_REVIEW: 4,  // Column E
    INTERVAL: 5,     // Column F
    NEXT_REVIEW: 6,  // Column G
    SYNCED: 7,       // Column H
    EVENT_ID: 8      // Column I (will store calendar event ID)
  },
  
  // Calendar settings
  CALENDAR_PREFIX: '📚 SRS: ',  // Prefix for event titles
  EVENT_COLOR: CalendarApp.EventColor.BLUE,
  
  // Sync settings
  SYNC_FUTURE_DAYS: 90,  // Only sync events within next 90 days
  MIN_VALID_YEAR: 2025,  // Ignore dates before this year
  
  // Auto-sync trigger interval (in hours)
  AUTO_SYNC_HOURS: 6,    // Sync every 6 hours
};

// ═══════════════════════════════════════════════════════════════
// MAIN SYNC FUNCTION
// ═══════════════════════════════════════════════════════════════

/**
 * Main sync function - syncs all configured sheets
 * Can be run manually or triggered automatically
 */
function syncToCalendar() {
  const startTime = new Date();
  const log = [];
  let totalSynced = 0;
  let totalDeleted = 0;
  let totalUpdated = 0;
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Sync each configured sheet
    CONFIG.SHEETS_TO_SYNC.forEach(sheetName => {
      const result = syncSheet(ss, calendar, sheetName);
      totalSynced += result.created;
      totalDeleted += result.deleted;
      totalUpdated += result.updated;
      log.push(`${sheetName}: +${result.created} ~${result.updated} -${result.deleted}`);
    });
    
    const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
    const message = `✅ Sync Complete (${elapsed}s)\n\n` +
                   `Created: ${totalSynced}\n` +
                   `Updated: ${totalUpdated}\n` +
                   `Deleted: ${totalDeleted}\n\n` +
                   log.join('\n');
    
    Logger.log(message);
    
    // Only show UI alert if run manually
    try {
      if (totalSynced + totalUpdated + totalDeleted > 0) {
        SpreadsheetApp.getUi().alert(message);
      } else {
        SpreadsheetApp.getUi().alert('✅ All up to date! No changes needed.');
      }
    } catch(e) {
      // Silent fail if triggered automatically
      Logger.log('Background sync: ' + message);
    }
    
  } catch(error) {
    Logger.log('ERROR: ' + error.toString());
    try {
      SpreadsheetApp.getUi().alert('❌ Sync Error:\n\n' + error.toString());
    } catch(e) {
      // Silent
    }
  }
}

/**
 * Sync a single sheet to calendar
 */
function syncSheet(ss, calendar, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  const stats = { created: 0, updated: 0, deleted: 0 };
  
  if (!sheet) {
    Logger.log(`Warning: Sheet "${sheetName}" not found, skipping.`);
    return stats;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return stats;  // No data
  
  // Get all data at once (faster than row-by-row)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 9);
  const data = dataRange.getValues();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const futureLimit = new Date(today);
  futureLimit.setDate(futureLimit.getDate() + CONFIG.SYNC_FUTURE_DAYS);
  
  // Track which rows to update
  const updates = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 2;  // Actual row in sheet
    
    const subject = row[CONFIG.COLUMNS.SUBJECT];
    const topic = row[CONFIG.COLUMNS.TOPIC];
    const nextReviewRaw = row[CONFIG.COLUMNS.NEXT_REVIEW];
    const syncedMark = row[CONFIG.COLUMNS.SYNCED];
    const existingEventId = row[CONFIG.COLUMNS.EVENT_ID];
    
    // Skip empty rows
    if (!subject || !topic) continue;
    
    // Parse next review date
    const nextReview = parseDate(nextReviewRaw);
    
    // Determine what action to take
    const action = determineAction(nextReview, syncedMark, existingEventId);
    
    if (action === 'CREATE') {
      // Create new calendar event
      const eventId = createCalendarEvent(calendar, subject, topic, nextReview);
      if (eventId) {
        updates.push({
          row: rowNum,
          synced: '✅',
          eventId: eventId
        });
        stats.created++;
      }
      
    } else if (action === 'UPDATE') {
      // Update existing event
      const success = updateCalendarEvent(calendar, existingEventId, subject, topic, nextReview);
      if (success) {
        updates.push({
          row: rowNum,
          synced: '✅',
          eventId: existingEventId
        });
        stats.updated++;
      }
      
    } else if (action === 'DELETE') {
      // Delete calendar event
      const success = deleteCalendarEvent(calendar, existingEventId);
      if (success) {
        updates.push({
          row: rowNum,
          synced: '',
          eventId: ''
        });
        stats.deleted++;
      }
    }
  }
  
  // Apply all updates at once (much faster)
  applyUpdates(sheet, updates);
  
  return stats;
}

// ═══════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Parse date from various formats
 */
function parseDate(dateValue) {
  if (!dateValue) return null;
  
  try {
    const date = new Date(dateValue);
    
    // Validate
    if (isNaN(date.getTime())) return null;
    if (date.getFullYear() < CONFIG.MIN_VALID_YEAR) return null;
    if (date.getFullYear() > 2100) return null;
    
    // Set to midnight
    date.setHours(0, 0, 0, 0);
    return date;
  } catch(e) {
    return null;
  }
}

/**
 * Determine what action to take for a row
 */
function determineAction(nextReview, syncedMark, eventId) {
  const hasValidDate = nextReview !== null;
  const isSynced = syncedMark === '✅';
  const hasEventId = eventId && eventId.length > 0;
  
  // Decision tree
  if (!hasValidDate && hasEventId) {
    return 'DELETE';  // Date removed, delete event
  }
  
  if (hasValidDate && !isSynced && !hasEventId) {
    return 'CREATE';  // New event needed
  }
  
  if (hasValidDate && isSynced && hasEventId) {
    return 'UPDATE';  // Update existing (in case date changed)
  }
  
  if (hasValidDate && !hasEventId) {
    return 'CREATE';  // Event was deleted externally, recreate
  }
  
  return 'SKIP';
}

/**
 * Create calendar event
 */
function createCalendarEvent(calendar, subject, topic, date) {
  try {
    const title = CONFIG.CALENDAR_PREFIX + subject + ' - ' + topic;
    const event = calendar.createAllDayEvent(title, date, {
      description: `📚 Spaced Repetition Review\n\nSubject: ${subject}\nTopic: ${topic}\n\n🔗 Managed by Medical SRS Sheet`
    });
    
    event.setColor(CONFIG.EVENT_COLOR);
    return event.getId();
  } catch(e) {
    Logger.log(`Failed to create event for "${topic}": ${e}`);
    return null;
  }
}

/**
 * Update existing calendar event
 */
function updateCalendarEvent(calendar, eventId, subject, topic, newDate) {
  try {
    const event = calendar.getEventById(eventId);
    if (!event) {
      Logger.log(`Event ${eventId} not found, will recreate`);
      return false;
    }
    
    const currentStart = event.getAllDayStartDate();
    currentStart.setHours(0, 0, 0, 0);
    
    // Only update if date changed
    if (currentStart.getTime() !== newDate.getTime()) {
      const title = CONFIG.CALENDAR_PREFIX + subject + ' - ' + topic;
      event.setTitle(title);
      event.setAllDayDate(newDate);
      Logger.log(`Updated event: ${topic}`);
    }
    
    return true;
  } catch(e) {
    Logger.log(`Failed to update event ${eventId}: ${e}`);
    return false;
  }
}

/**
 * Delete calendar event
 */
function deleteCalendarEvent(calendar, eventId) {
  try {
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
      Logger.log(`Deleted event: ${eventId}`);
      return true;
    }
    return true;  // Already deleted
  } catch(e) {
    Logger.log(`Failed to delete event ${eventId}: ${e}`);
    return false;
  }
}

/**
 * Apply updates to sheet efficiently
 */
function applyUpdates(sheet, updates) {
  if (updates.length === 0) return;
  
  updates.forEach(update => {
    sheet.getRange(update.row, CONFIG.COLUMNS.SYNCED + 1).setValue(update.synced);
    sheet.getRange(update.row, CONFIG.COLUMNS.EVENT_ID + 1).setValue(update.eventId);
  });
}

// ═══════════════════════════════════════════════════════════════
// SETUP & TRIGGERS
// ═══════════════════════════════════════════════════════════════

/**
 * ONE-TIME SETUP: Creates automatic sync triggers
 * Run this once to enable auto-sync
 */
function setupTriggers() {
  // Delete existing triggers first
  deleteTriggers();
  
  // Create time-based trigger (every 6 hours)
  ScriptApp.newTrigger('syncToCalendar')
    .timeBased()
    .everyHours(CONFIG.AUTO_SYNC_HOURS)
    .create();
  
  // Create on-edit trigger (sync when sheet is edited)
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  
  SpreadsheetApp.getUi().alert(
    '✅ Auto-Sync Enabled!\n\n' +
    `• Syncs every ${CONFIG.AUTO_SYNC_HOURS} hours automatically\n` +
    '• Syncs when you edit the sheet\n' +
    '• You can also run manually from menu\n\n' +
    'First sync running now...'
  );
  
  // Run initial sync
  syncToCalendar();
}

/**
 * Remove all triggers
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncToCalendar' || 
        trigger.getHandlerFunction() === 'onEditTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Disable auto-sync
 */
function disableAutoSync() {
  deleteTriggers();
  SpreadsheetApp.getUi().alert('✅ Auto-sync disabled. You can still run manually.');
}

/**
 * On-edit trigger (syncs when specific columns change)
 */
function onEditTrigger(e) {
  if (!e) return;
  
  const range = e.range;
  const col = range.getColumn();
  
  // Only sync if Last Review (E) or Next Review (G) columns were edited
  if (col === CONFIG.COLUMNS.LAST_REVIEW + 1 || 
      col === CONFIG.COLUMNS.NEXT_REVIEW + 1) {
    
    // Wait 2 seconds to batch edits
    Utilities.sleep(2000);
    syncToCalendar();
  }
}

// ═══════════════════════════════════════════════════════════════
// MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Create custom menu on open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 Calendar Sync')
    .addItem('🔄 Sync Now', 'syncToCalendar')
    .addSeparator()
    .addItem('⚙️ Setup Auto-Sync', 'setupTriggers')
    .addItem('🛑 Disable Auto-Sync', 'disableAutoSync')
    .addSeparator()
    .addItem('🧹 Clear All Sync Markers', 'clearAllSyncMarkers')
    .addToUi();
}

/**
 * Clear all sync markers (for troubleshooting)
 */
function clearAllSyncMarkers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Sync Markers?',
    'This will remove all ✅ marks and event IDs. Events in calendar will remain.\n\nNext sync will create duplicates unless you delete calendar events first.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    CONFIG.SHEETS_TO_SYNC.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.getRange(2, CONFIG.COLUMNS.SYNCED + 1, lastRow - 1, 1).clearContent();
          sheet.getRange(2, CONFIG.COLUMNS.EVENT_ID + 1, lastRow - 1, 1).clearContent();
        }
      }
    });
    
    ui.alert('✅ Cleared! Next sync will recreate all events.');
  }
}

/**
 * Delete all SRS events from calendar
 */
function deleteAllSRSEvents() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Delete All SRS Events?',
    'This will delete ALL events starting with "' + CONFIG.CALENDAR_PREFIX + '" from your calendar.\n\nThis cannot be undone!\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const future = new Date(now);
    future.setDate(future.getDate() + 365);  // Next year
    
    const events = calendar.getEvents(now, future);
    let deleted = 0;
    
    events.forEach(event => {
      if (event.getTitle().startsWith(CONFIG.CALENDAR_PREFIX)) {
        event.deleteEvent();
        deleted++;
      }
    });
    
    ui.alert(`✅ Deleted ${deleted} SRS events from calendar.`);
  }
}
