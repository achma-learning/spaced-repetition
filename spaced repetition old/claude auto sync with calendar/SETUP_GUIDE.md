# 📅 MEDICAL SRS → GOOGLE CALENDAR SYNC
## Complete Setup Guide

---

## 🚀 QUICK SETUP (5 Minutes)

### Step 1: Add Event ID Column to Your Sheet

1. Open your spreadsheet (the one you uploaded)
2. Go to the **S8** sheet
3. Click on column **I** (after the "Synced" column H)
4. Right-click → Insert 1 column to the left
5. In cell **I2**, enter header: **Event ID**
6. Leave all cells in column I empty (script will fill them)

**Your columns should now be:**
- A: Subject
- B: Topic  
- C: Status
- D: Mastery
- E: Last Review
- F: Interval
- G: Next Review
- H: Synced ✅
- I: Event ID (NEW!)

---

### Step 2: Install the Script

1. In your Google Sheet, click **Extensions** → **Apps Script**
2. Delete any existing code in the editor
3. Copy the ENTIRE script from `CalendarSync.js` 
4. Paste it into the Apps Script editor
5. Click **Save** (disk icon) and name it: "SRS Calendar Sync"
6. Close the Apps Script tab

---

### Step 3: Enable Auto-Sync

1. **Refresh your spreadsheet** (reload the page)
2. You'll see a new menu: **📅 Calendar Sync**
3. Click **📅 Calendar Sync** → **⚙️ Setup Auto-Sync**
4. **Grant permissions when asked:**
   - Click "Review Permissions"
   - Select your Google account
   - Click "Advanced" → "Go to [project name] (unsafe)"
   - Click "Allow"
5. The script will run and sync all lessons to your calendar!

---

### Step 4: Verify It Works

1. Open **Google Calendar**
2. Look for events starting with **"📚 SRS:"**
3. You should see all your lessons scheduled on their Next Review dates!
4. Go back to your sheet → Column H should have ✅ marks
5. Column I should have event IDs

**Done! 🎉**

---

## 🎯 HOW IT WORKS

### Automatic Sync
The script now syncs:
- ✅ **Every 6 hours** automatically
- ✅ **When you edit** the Last Review or Next Review columns
- ✅ **Manual sync** anytime from menu

### What Gets Synced
- **Creates** calendar events for new lessons
- **Updates** events when dates change
- **Deletes** events when lessons are removed or dates cleared
- **No duplicates** - tracks events with ID column

### Smart Features
- **Handles date changes**: Update "Next Review" → Event moves automatically
- **Conflict resolution**: Won't create duplicates
- **Batch processing**: Fast even with 100+ lessons
- **Error recovery**: Continues if one event fails

---

## 📋 DAILY WORKFLOW

### Studying Lessons

**IN YOUR SHEET:**
1. Study the lesson
2. Update **Last Review** (col E) → Press `Ctrl+;` to enter today
3. Update **Mastery** level (col D) → Increase by 1
4. **Next Review** (col G) auto-calculates!

**WHAT HAPPENS:**
- Script detects the change
- Updates the calendar event to new date
- ✅ mark stays in place
- Event ID preserved

**No manual sync needed!** ✨

---

## ⚙️ CONFIGURATION OPTIONS

Edit these in the script if needed:

```javascript
const CONFIG = {
  // Sync multiple sheets
  SHEETS_TO_SYNC: ['S8'],  // Change to ['S8', 'AIO'] for both
  
  // Event prefix
  CALENDAR_PREFIX: '📚 SRS: ',  // Change if you want different prefix
  
  // Sync frequency
  AUTO_SYNC_HOURS: 6,  // Change to sync more/less often
  
  // Only sync events in next N days
  SYNC_FUTURE_DAYS: 90,  // Increase if you want farther future
};
```

---

## 🛠️ TROUBLESHOOTING

### "Sheet not found" error
- Check that `SHEETS_TO_SYNC` matches your sheet name exactly
- Sheet names are case-sensitive!

### Duplicate events in calendar
1. Click **📅 Calendar Sync** → **🧹 Clear All Sync Markers**
2. Manually delete duplicate events from Google Calendar
3. Click **📅 Calendar Sync** → **🔄 Sync Now**

### Events not updating
1. Check that column I (Event ID) has values
2. If empty, the script will create new events
3. Run **🔄 Sync Now** manually

### Permission errors
- Re-run **⚙️ Setup Auto-Sync** to re-authorize
- Make sure you're signed in to correct Google account

### Events showing wrong dates
- Check that **Next Review** (col G) has valid dates
- Script ignores dates before 2025 or after 2100

---

## 🎨 CUSTOMIZATION IDEAS

### Different Calendar
Change this line in the script:
```javascript
const calendar = CalendarApp.getDefaultCalendar();
```
To:
```javascript
const calendar = CalendarApp.getCalendarById('your-calendar-id@group.calendar.google.com');
```

### Different Event Colors
Change in CONFIG:
```javascript
EVENT_COLOR: CalendarApp.EventColor.RED,  // or GREEN, YELLOW, etc.
```

### Custom Event Description
Edit the `createCalendarEvent` function:
```javascript
description: `Your custom description here\n\nSubject: ${subject}`
```

---

## 📊 ADVANCED FEATURES

### Sync Multiple Sheets
```javascript
SHEETS_TO_SYNC: ['S8', 'AIO', 'Hebdomadaire'],
```

### Different Column Layout
If your columns are different, adjust in CONFIG:
```javascript
COLUMNS: {
  SUBJECT: 0,      // First column (A)
  TOPIC: 1,        // Second column (B)
  // ... etc
}
```

### Disable Auto-Sync Temporarily
Click **📅 Calendar Sync** → **🛑 Disable Auto-Sync**

(Manual sync still works from menu)

---

## 🔒 PRIVACY & SECURITY

### What This Script Does
- ✅ Reads your spreadsheet data
- ✅ Creates/updates/deletes events in YOUR calendar
- ✅ Runs on YOUR Google account
- ✅ All data stays in YOUR Google Drive & Calendar

### What It Doesn't Do
- ❌ Doesn't send data anywhere external
- ❌ Doesn't access other calendars
- ❌ Doesn't share with anyone
- ❌ No third-party servers

**100% runs within Google's infrastructure!**

---

## 💡 TIPS & BEST PRACTICES

### Pro Tips
1. **Keep Event ID column (I) visible** - helps troubleshoot
2. **Don't manually edit column I** - let script manage it
3. **Use the menu for manual syncs** - faster than triggers
4. **Check logs** if something fails:
   - Apps Script editor → Executions tab

### Performance
- Script handles **200+ lessons** easily
- Syncs in **2-5 seconds** typically
- Batch updates are efficient

### Backup
- Column I stores event IDs - **don't delete it!**
- If you accidentally clear it, events will duplicate
- Use **Clear All Sync Markers** feature if needed

---

## 📞 MENU OPTIONS EXPLAINED

### 🔄 Sync Now
- Manual sync
- Use when you want instant update
- Shows detailed results

### ⚙️ Setup Auto-Sync  
- One-time setup
- Creates automatic triggers
- Run this first!

### 🛑 Disable Auto-Sync
- Stops automatic syncing
- Manual sync still works
- Good for maintenance

### 🧹 Clear All Sync Markers
- Removes all ✅ and event IDs
- Use if you want fresh start
- **Warning:** Next sync creates duplicates if calendar events still exist!

---

## ✅ SUCCESS CHECKLIST

After setup, verify:

- [ ] Column I (Event ID) exists in S8 sheet
- [ ] Script installed in Apps Script editor
- [ ] Ran "Setup Auto-Sync" and granted permissions
- [ ] See ✅ marks in column H
- [ ] See event IDs in column I
- [ ] Events appear in Google Calendar
- [ ] Events have "📚 SRS:" prefix
- [ ] Menu "📅 Calendar Sync" appears in sheet

**If all checked → You're done! 🎉**

---

## 🆘 NEED HELP?

### Check Logs
1. Extensions → Apps Script
2. Left sidebar → Executions
3. Click on latest execution
4. See what happened

### Common Issues
| Problem | Solution |
|---------|----------|
| No menu appears | Refresh page, check script saved |
| No events created | Check dates are valid (year > 2025) |
| Duplicates | Use "Clear Sync Markers" feature |
| Wrong calendar | Change calendar ID in script |

---

## 🎓 EXAMPLES

### Example 1: First Time Setup
```
Sheet has 50 lessons with Next Review dates
↓
Run "Setup Auto-Sync"
↓
Script creates 50 calendar events
↓
Column H: all ✅
Column I: all have event IDs
↓
Done! Events appear in calendar
```

### Example 2: Daily Use
```
You study "Trauma du bassin"
↓
Update Last Review: 10/02/26
↓
Update Mastery: 0 → 1
↓
Next Review auto-calculates: 11/02/26
↓
(wait 2 seconds)
↓
Script auto-runs
↓
Calendar event moves to 11/02/26
✓ Still marked synced
```

### Example 3: Deleting a Lesson
```
You remove a lesson (clear row)
↓
Script detects missing Next Review
↓
Deletes calendar event
↓
Clears ✅ and event ID
```

---

**Enjoy your automated medical SRS system! 🚀📚**
