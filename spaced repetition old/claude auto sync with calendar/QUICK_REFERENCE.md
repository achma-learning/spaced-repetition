# 📱 QUICK REFERENCE CARD
## Medical SRS + Google Calendar Sync

---

## 🎯 DAILY WORKFLOW

### Study a Lesson
```
1. Study the material
2. Press Ctrl+; in "Last Review" (col E) → enters today's date
3. Increase "Mastery" (col D) by 1
4. "Next Review" auto-calculates!
5. Wait 2 seconds → Calendar auto-updates ✨
```

**That's it! No manual sync needed.**

---

## 📅 MENU SHORTCUTS

| Action | Menu Path |
|--------|-----------|
| **Sync Now** | 📅 Calendar Sync → 🔄 Sync Now |
| **Setup Auto-Sync** | 📅 Calendar Sync → ⚙️ Setup Auto-Sync |
| **Stop Auto-Sync** | 📅 Calendar Sync → 🛑 Disable Auto-Sync |
| **Clear Markers** | 📅 Calendar Sync → 🧹 Clear All Sync Markers |

---

## 🔧 KEYBOARD SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `Ctrl+;` | Insert today's date |
| `Ctrl+Shift+;` | Insert current time |
| `F5` | Refresh page |

---

## 📊 COLUMN GUIDE

| Col | Name | You Edit? | Auto? |
|-----|------|-----------|-------|
| A | Subject | ✅ Yes | |
| B | Topic | ✅ Yes | |
| C | Status | | ✅ Formula |
| D | Mastery | ✅ Yes | |
| E | Last Review | ✅ Yes | |
| F | Interval | | ✅ Formula |
| G | Next Review | | ✅ Formula |
| H | Synced | | ✅ Script |
| I | Event ID | ❌ No | ✅ Script |

**Only edit: A, B, D, E**

---

## ⚡ AUTO-SYNC BEHAVIOR

### Triggers Sync When:
- ✅ Every 6 hours automatically
- ✅ You edit "Last Review" (col E)
- ✅ You edit "Next Review" (col G)
- ✅ You click "Sync Now" in menu

### What Happens:
- **New lesson** with date → Creates calendar event
- **Date changed** → Updates calendar event  
- **Date removed** → Deletes calendar event
- **Already synced** → Skips (no duplicate)

---

## 🎨 CALENDAR EVENTS

### Event Format
```
📚 SRS: [Subject] - [Topic]

Example:
📚 SRS: Tramatologie - Fc coude + Avant bras
```

### Event Details
- **All-day event** on Next Review date
- **Blue color** (customizable)
- **Description** includes subject & topic
- **Managed by script** (don't edit in calendar)

---

## ✅ STATUS INDICATORS

| Symbol | Meaning |
|--------|---------|
| ⚡ STUDY NOW | Due today or overdue |
| ✅ Wait | Not due yet |
| ✅ (col H) | Synced to calendar |
| (blank col H) | Not yet synced |
| Event ID (col I) | Linked to calendar event |

---

## 🚨 TROUBLESHOOTING

### No calendar events?
→ Check Next Review dates are valid (year > 2025)
→ Run "Sync Now" from menu

### Duplicate events?
→ Click "Clear All Sync Markers"
→ Delete duplicates from Google Calendar
→ Run "Sync Now"

### Event won't update?
→ Check Event ID (col I) is filled
→ Try manual "Sync Now"

### Script error?
→ Extensions → Apps Script → Executions
→ Check error message
→ Re-run "Setup Auto-Sync"

---

## 📈 MASTERY LEVELS

| Level | Interval | When to Use |
|-------|----------|-------------|
| 0 | 1 day | First time learning |
| 1 | 3 days | Recalled well |
| 2 | 7 days | Getting solid |
| 3 | 14 days | Pretty confident |
| 4 | 30 days | Very confident |
| 5 | 60 days | Nearly mastered |
| 6 | 90 days | Mastered |
| 7 | 120 days | Long-term retention |

**Tip:** Increase by 1 each time you recall successfully!

---

## 💡 PRO TIPS

### Efficiency
- Use filter to show only "⚡ STUDY NOW"
- Sort by Next Review to prioritize
- Batch study similar topics together

### Maintenance
- Check column I has event IDs
- Don't manually edit column I
- Keep ✅ marks visible

### Calendar
- Use Google Calendar app on phone
- Get notifications for reviews
- Color-code by subject (manual)

---

## 🔄 SYNC STATUS

### How to Check
1. Look at column H - should have ✅
2. Look at column I - should have event ID
3. Open Google Calendar - see events
4. Check menu → last sync time

### Force Refresh
```
Menu → 📅 Calendar Sync → 🔄 Sync Now
```

---

## 📞 QUICK FIXES

| Problem | 2-Second Fix |
|---------|--------------|
| Not syncing | Menu → Sync Now |
| Duplicates | Clear Markers + Sync |
| Wrong date | Edit col G + Sync |
| Missing event | Check col I has ID |
| Menu missing | Refresh page (F5) |

---

## 🎯 BEST PRACTICES

### Daily
1. Filter "⚡ STUDY NOW" 
2. Study each lesson
3. Update Last Review (Ctrl+;)
4. Increase Mastery +1
5. Done!

### Weekly
- Review upcoming in calendar
- Plan study time
- Adjust if needed

### Monthly
- Check sync status
- Verify no duplicates
- Review progress

---

## 📊 FORMULAS (READ-ONLY)

### Status (Col C)
```
=IF(G3 <= TODAY(), "⚡ STUDY NOW", "✅ Wait")
```

### Interval (Col F)
```
=CHOOSE(D3+1, 1, 3, 7, 14, 30, 60, 90, 120)
```

### Next Review (Col G)
```
=E3 + F3
```

**Don't edit these - they auto-calculate!**

---

## ⏱️ TIME ESTIMATES

| Task | Time |
|------|------|
| First setup | 5 minutes |
| Daily review | 2-10 minutes |
| Update after study | 5 seconds |
| Manual sync | 2 seconds |
| Check calendar | 1 second |

---

## 🎓 STUDY STRATEGY

### Optimal Review
1. **Before bed** - better retention
2. **Spaced out** - don't cram
3. **Active recall** - test yourself
4. **Consistent** - every day

### Using Calendar
- Set notifications 1 day before
- Block study time in calendar
- Use calendar view for planning
- Sync across all devices

---

**Keep this card handy! 📌**

Print or bookmark for quick reference.
