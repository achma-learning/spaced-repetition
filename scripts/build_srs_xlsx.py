# -*- coding: utf-8 -*-
"""
build_srs_xlsx.py — (re)generate ../srs-system.xlsx from the FMPM curriculum.

WHAT IT DOES
  • Reads data/curriculum_FMPM_S1-S10.txt (the single source of truth).
  • Rebuilds srs-system.xlsx with 5 sheets:
      Dashboard (macro/bird's-eye + charts) · lesson-database (the engine) ·
      Module View (micro/per-module) · Weekly · How-To
  • PRESERVES your study progress: before overwriting, it reads the existing
    srs-system.xlsx and re-applies Last Review + Mastery by matching
    (Semester | Module | Lesson). So it is safe to re-run after editing the
    curriculum or the visuals.

USAGE   (run from the repo root)
    python3 scripts/build_srs_xlsx.py

⚠️  If you change the column order here, update CONFIG.COL in srs-appscript.js
    too — the Apps Script reads cells by position. And keep the Interval rule
    (column H formula below) identical to the [1,3,7,14,30,60] array in the JS.
"""
import re, os
import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
SRC = 'data/curriculum_FMPM_S1-S10.txt'
OUT = 'srs-system.xlsx'
LAST = 2000           # formula range upper bound (data ~1310 rows; headroom to 2000)
H = 2                 # header rows; data starts row 3

# ───────────────────────── parse curriculum ─────────────────────────
def parse(path):
    toks = []
    for raw in open(path, encoding='utf-8'):
        line = raw.rstrip('\n')
        if not line.strip():
            toks.append((-1, '')); continue
        if re.match(r'^S\d+$', line.strip()) and not line.startswith((' ', '-')):
            toks.append((0, line.strip())); continue
        mm = re.match(r'^(\s*)-\s+(.*)$', line)
        if mm:
            toks.append((1 + len(mm.group(1)) // 2, mm.group(2).strip()))
        else:
            toks.append((99, line.strip()))
    out, sem, mod, subj, n = [], None, None, None, len(toks)
    for i, (lvl, txt) in enumerate(toks):
        if lvl == 0:
            sem, mod, subj = txt, None, None
        elif lvl == 1:
            mod, subj = txt, None
        elif lvl == 2:
            j = i + 1
            while j < n and toks[j][0] == -1:
                j += 1
            if j < n and toks[j][0] == 3:
                subj = txt
            else:
                out.append([sem, mod, '', txt]); subj = None
        elif lvl == 3:
            out.append([sem, mod, subj or '', txt])
    return out

rows = parse(SRC)
modpairs = []
for s, m, _, _ in rows:
    if (s, m) not in modpairs:
        modpairs.append((s, m))
NMOD = len(modpairs)
print(f'Parsed {len(rows)} lessons across {NMOD} module/semester pairs')

# ── read existing progress so re-running never wipes study history ──
progress = {}
EXAM_DATE = datetime.date(2026, 5, 12)   # default; K1 of an existing workbook wins
if os.path.exists(OUT):
    try:
        ex = openpyxl.load_workbook(OUT, data_only=False)['lesson-database']
        for r in range(3, ex.max_row + 1):
            sem, mod, les = ex.cell(r, 2).value, ex.cell(r, 3).value, ex.cell(r, 5).value
            lr, ma = ex.cell(r, 6).value, ex.cell(r, 7).value
            if les and ((lr not in (None, '')) or (ma not in (None, ''))):
                progress[(sem, mod, les)] = (lr, ma)
        print(f'Found existing progress for {len(progress)} lessons (will preserve)')
        prev_exam = ex.cell(1, 11).value          # K1 = exam date (user-editable)
        if isinstance(prev_exam, (datetime.date, datetime.datetime)):
            EXAM_DATE = prev_exam
            print(f'Preserving exam date from K1: {EXAM_DATE}')
    except Exception as e:
        print('Could not read existing progress:', e)

# ───────────────────────── styles ─────────────────────────
WHITE = Font(color='FFFFFF', bold=True)
BOLD = Font(bold=True)
TITLE_FILL = PatternFill('solid', fgColor='1F3864')
HEAD_FILL = PatternFill('solid', fgColor='2E5496')
SUB_FILL = PatternFill('solid', fgColor='D6DCE4')
INPUT_FILL = PatternFill('solid', fgColor='DDEBF7')   # light blue = "edit here"
KPI_FILL = PatternFill('solid', fgColor='EAF1FB')
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CEN = Alignment(horizontal='center', vertical='center')

# one soft fill per module (cycled) so each module block reads as one unit;
# deliberately excludes DDEBF7 (the "edit here" blue on F/G)
MODULE_COLORS = ['FDE9D9', 'E2EFDA', 'D9E1F2', 'FFF2CC', 'E4DFEC',
                 'FCE4D6', 'DAEEF3', 'F2DCDB', 'EBF1DE', 'F2F2F2']
MODULE_FILLS = [PatternFill('solid', fgColor=c) for c in MODULE_COLORS]

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = HEAD_FILL; cell.font = WHITE
        cell.alignment = CEN; cell.border = BORDER

# ───────────────────────── workbook ─────────────────────────
wb = Workbook()
db = wb.active; db.title = 'Dashboard'
ldb = wb.create_sheet('lesson-database')
mv = wb.create_sheet('Module View')
wk = wb.create_sheet('Weekly')
ht = wb.create_sheet('How-To')
db.sheet_properties.tabColor = '70AD47'
ldb.sheet_properties.tabColor = '2E75B6'
mv.sheet_properties.tabColor = 'ED7D31'

# ═══════════════ lesson-database (engine) ═══════════════
HEADERS = ['#', 'Semester', 'Module', 'Subject', 'Lesson', 'Last Review',
           'Mastery', 'Interval', 'Next Review', 'Status', 'Priority',
           'Synced', 'Event ID']
# row 1 = info strip: title · today · days left until exam · exam date (K1, editable)
for c in range(1, len(HEADERS) + 1):
    cell = ldb.cell(1, c)
    cell.fill = TITLE_FILL; cell.font = WHITE
    cell.alignment = Alignment(horizontal='left', vertical='center')
ldb.merge_cells('A1:D1')
ldb['A1'] = '🧠 MEDICAL SRS — type only F & G · ▼ filters on row 2'
ldb['E1'] = '=TODAY()'
ldb['E1'].number_format = '"📅 Today:  "dddd", "mmmm dd", "yyyy'
ldb.merge_cells('F1:H1')
ldb['F1'] = '=MAX(0,$K$1-TODAY())'
ldb['F1'].number_format = '"⏳ Days left: "0" days available"'
ldb.merge_cells('I1:J1')
ldb['I1'] = '🎯 Exam date:'
ldb['I1'].alignment = Alignment(horizontal='right', vertical='center')
ldb.merge_cells('K1:L1')
ek = ldb['K1']
ek.value = EXAM_DATE
ek.number_format = 'dd/mm/yyyy'
ek.fill = INPUT_FILL
ek.font = Font(bold=True, color='1F3864')
ek.alignment = CEN; ek.border = BORDER
ldb.row_dimensions[1].height = 26
for c, hdr in enumerate(HEADERS, 1):
    ldb.cell(2, c, hdr)
style_header(ldb, 2, len(HEADERS))
for c in (6, 7):
    ldb.cell(2, c).fill = INPUT_FILL; ldb.cell(2, c).font = Font(bold=True, color='1F3864')

date_fmt = 'yyyy-mm-dd'
last_data = H + len(rows)
pair2idx = {p: i for i, p in enumerate(modpairs)}
for i, (sem, mod, subj, lesson) in enumerate(rows):
    r = H + 1 + i
    ldb.cell(r, 1, '=ROW()-2').alignment = CEN
    ldb.cell(r, 2, sem); ldb.cell(r, 3, mod); ldb.cell(r, 4, subj); ldb.cell(r, 5, lesson)
    mf = MODULE_FILLS[pair2idx[(sem, mod)] % len(MODULE_FILLS)]
    for c in range(1, 6):                     # A–E share the module's color
        ldb.cell(r, c).fill = mf
    ldb.cell(r, 6).number_format = date_fmt
    ldb.cell(r, 6).fill = INPUT_FILL
    ldb.cell(r, 7).fill = INPUT_FILL; ldb.cell(r, 7).alignment = CEN
    ldb.cell(r, 8, f'=IF($G{r}="",1,IF($G{r}=0,1,IF($G{r}=1,3,IF($G{r}=2,7,IF($G{r}=3,14,IF($G{r}=4,30,60))))))').alignment = CEN
    ldb.cell(r, 9, f'=IF($F{r}="","",$F{r}+$H{r})'); ldb.cell(r, 9).number_format = date_fmt
    ldb.cell(r, 10, f'=IF($F{r}="","⚪ NEW",IF($G{r}>=5,"✅ DONE",IF($I{r}<TODAY(),"🔴 OVERDUE",IF($I{r}=TODAY(),"🟢 TODAY",IF($I{r}=TODAY()+1,"🔵 TOMORROW",IF($I{r}<=TODAY()+7,"📅 THIS WEEK","⏳ SCHEDULED"))))))')
    ldb.cell(r, 11, f'=IF($F{r}="",999,IF($G{r}>=5,9999,IF($I{r}="",999,$I{r}-TODAY())))').alignment = CEN

# re-apply preserved progress (match on Semester|Module|Lesson)
applied = 0
for i, (sem, mod, subj, lesson) in enumerate(rows):
    hit = progress.get((sem, mod, lesson))
    if hit:
        lr, ma = hit; rr = H + 1 + i
        if lr not in (None, ''):
            ldb.cell(rr, 6, lr); ldb.cell(rr, 6).number_format = date_fmt
        if ma not in (None, ''):
            ldb.cell(rr, 7, int(ma))
        applied += 1
print(f'Re-applied progress to {applied} lessons')

for col, w in {'A':5,'B':9,'C':26,'D':28,'E':52,'F':13,'G':9,'H':9,'I':13,'J':14,'K':9,'L':9,'M':22}.items():
    ldb.column_dimensions[col].width = w
ldb.column_dimensions['M'].hidden = True
ldb.freeze_panes = 'F3'
ldb.auto_filter.ref = f'A2:K{last_data}'
ldb.conditional_formatting.add(f'G3:G{last_data}',
    ColorScaleRule(start_type='num', start_value=0, start_color='F8696B',
                   mid_type='num', mid_value=2, mid_color='FFEB84',
                   end_type='num', end_value=5, end_color='63BE7B'))
for kw, color in {'OVERDUE':'FFC7CE','TODAY':'C6EFCE','TOMORROW':'BDD7EE',
                  'THIS WEEK':'FFF2CC','DONE':'E2EFDA','SCHEDULED':'F2F2F2'}.items():
    ldb.conditional_formatting.add(f'J3:J{last_data}',
        FormulaRule(formula=[f'ISNUMBER(SEARCH("{kw}",$J3))'],
                    fill=PatternFill('solid', fgColor=color)))
dv = DataValidation(type='list', formula1='"0,1,2,3,4,5"', allow_blank=True, showErrorMessage=True)
dv.prompt = 'How well did you recall it? 0 = forgot … 5 = mastered'; dv.promptTitle = 'Mastery'
ldb.add_data_validation(dv); dv.add(f'G3:G{last_data}')

# ═══════════════ Dashboard (MACRO) ═══════════════
def L(col, r1, r2):
    return f"'lesson-database'!${col}${r1}:${col}${r2}"

db.sheet_view.showGridLines = False
db.merge_cells('A1:R1')
t = db['A1']; t.value = "📊 MEDICAL SRS — DASHBOARD  ·  bird's-eye view of every semester & module"
t.fill = TITLE_FILL; t.font = Font(color='FFFFFF', bold=True, size=13)
t.alignment = Alignment(horizontal='left', vertical='center'); db.row_dimensions[1].height = 30

kpis = [
    ('📚 Total lessons',  f"=COUNTA({L('E',3,LAST)})"),
    ('✏️ Started',        f"=COUNTA({L('F',3,LAST)})"),
    ('✅ Mastered',       f"=COUNTIF({L('G',3,LAST)},5)"),
    ('⏳ In progress',    '=B4-B5'),
    ('🔴 Overdue',        f'=COUNTIF({L("J",3,LAST)},"*OVERDUE*")'),
    ('🟢 Due today',      f'=COUNTIF({L("J",3,LAST)},"*TODAY*")'),
    ('📅 Due this week',  f'=COUNTIF({L("J",3,LAST)},"*THIS WEEK*")'),
    ('📈 % Complete',     '=IF(B3=0,0,B5/B3)'),
]
db['A2'] = 'AT A GLANCE'; db['A2'].font = BOLD
for i, (lab, frm) in enumerate(kpis):
    r = 3 + i
    db.cell(r, 1, lab).fill = KPI_FILL
    vc = db.cell(r, 2, frm); vc.font = Font(bold=True, size=12); vc.alignment = CEN; vc.fill = KPI_FILL
    db.cell(r, 1).border = BORDER; vc.border = BORDER
db['B10'].number_format = '0%'

db['A12'] = 'BY SEMESTER'; db['A12'].font = BOLD
for c, hh in enumerate(['Semester', 'Total', 'Started', 'Mastered', 'Overdue', '% Done'], 1):
    db.cell(13, c, hh)
style_header(db, 13, 6)
for i in range(10):
    r = 14 + i; s = f'S{i+1}'
    db.cell(r, 1, s).alignment = CEN
    db.cell(r, 2, f"=COUNTIF({L('B',3,LAST)},$A{r})")
    db.cell(r, 3, f'=COUNTIFS({L("B",3,LAST)},$A{r},{L("F",3,LAST)},"<>")')
    db.cell(r, 4, f"=COUNTIFS({L('B',3,LAST)},$A{r},{L('G',3,LAST)},5)")
    db.cell(r, 5, f'=COUNTIFS({L("B",3,LAST)},$A{r},{L("J",3,LAST)},"*OVERDUE*")')
    db.cell(r, 6, f'=IF($B{r}=0,0,$D{r}/$B{r})'); db.cell(r, 6).number_format = '0%'
    for c in range(1, 7):
        db.cell(r, c).border = BORDER
sem_last = 23
db.conditional_formatting.add(f'F14:F{sem_last}',
    DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='63BE7B'))

db['H12'] = 'MASTERY MIX'; db['H12'].font = BOLD
db.cell(13, 8, 'Mastery'); db.cell(13, 9, 'Lessons'); style_header(db, 13, 9)
for k in range(6):
    r = 14 + k
    db.cell(r, 8, k).alignment = CEN
    db.cell(r, 9, f"=COUNTIF({L('G',3,LAST)},$H{r})")
    db.cell(r, 8).border = BORDER; db.cell(r, 9).border = BORDER

for col, w in {'A':20,'B':9,'C':9,'D':10,'E':9,'F':9,'G':3,'H':9,'I':9}.items():
    db.column_dimensions[col].width = w

c1 = BarChart(); c1.type = 'col'; c1.title = 'Lessons per semester'; c1.legend = None
c1.add_data(Reference(db, min_col=2, min_row=13, max_row=sem_last), titles_from_data=True)
c1.set_categories(Reference(db, min_col=1, min_row=14, max_row=sem_last))
c1.height, c1.width = 7.5, 15; db.add_chart(c1, 'K3')
c2 = BarChart(); c2.type = 'col'; c2.title = 'Progress per semester'; c2.grouping = 'stacked'; c2.overlap = 100
c2.add_data(Reference(db, min_col=3, max_col=4, min_row=13, max_row=sem_last), titles_from_data=True)
c2.set_categories(Reference(db, min_col=1, min_row=14, max_row=sem_last))
c2.height, c2.width = 7.5, 15; db.add_chart(c2, 'K19')
c3 = BarChart(); c3.type = 'col'; c3.title = 'Mastery distribution (0=forgot → 5=mastered)'; c3.legend = None
c3.add_data(Reference(db, min_col=9, min_row=13, max_row=19), titles_from_data=True)
c3.set_categories(Reference(db, min_col=8, min_row=14, max_row=19))
c3.height, c3.width = 7.5, 15; db.add_chart(c3, 'A26')

MHR = 45
db.cell(MHR - 1, 1, 'BY MODULE  (filter the lesson-database tab, or use the Module View tab for a detailed view)').font = BOLD
for c, hh in enumerate(['Module (pick this in Module View)', 'Sem', 'Module name', 'Total', 'Started', 'Mastered', 'Overdue', '% Done'], 1):
    db.cell(MHR, c, hh)
style_header(db, MHR, 8)
for i, (s, m) in enumerate(modpairs):
    r = MHR + 1 + i
    db.cell(r, 1, f'{s} — {m}')
    db.cell(r, 2, s).alignment = CEN
    db.cell(r, 3, m)
    db.cell(r, 4, f"=COUNTIFS({L('B',3,LAST)},$B{r},{L('C',3,LAST)},$C{r})")
    db.cell(r, 5, f'=COUNTIFS({L("B",3,LAST)},$B{r},{L("C",3,LAST)},$C{r},{L("F",3,LAST)},"<>")')
    db.cell(r, 6, f"=COUNTIFS({L('B',3,LAST)},$B{r},{L('C',3,LAST)},$C{r},{L('G',3,LAST)},5)")
    db.cell(r, 7, f'=COUNTIFS({L("B",3,LAST)},$B{r},{L("C",3,LAST)},$C{r},{L("J",3,LAST)},"*OVERDUE*")')
    db.cell(r, 8, f'=IF($D{r}=0,0,$F{r}/$D{r})'); db.cell(r, 8).number_format = '0%'
    for c in range(1, 9):
        db.cell(r, c).border = BORDER
mod_last = MHR + NMOD
db.conditional_formatting.add(f'H{MHR+1}:H{mod_last}',
    DataBarRule(start_type='num', start_value=0, end_type='num', end_value=1, color='5B9BD5'))
c4 = BarChart(); c4.type = 'bar'; c4.title = 'Lessons per module'; c4.legend = None
c4.add_data(Reference(db, min_col=4, min_row=MHR, max_row=mod_last), titles_from_data=True)
c4.set_categories(Reference(db, min_col=1, min_row=MHR + 1, max_row=mod_last))
c4.height, c4.width = 28, 16; db.add_chart(c4, 'J45')

# ═══════════════ Module View (MICRO) ═══════════════
mv.sheet_view.showGridLines = False
mv.merge_cells('A1:H1')
t = mv['A1']; t.value = '🔍 MODULE VIEW — pick one module to see its lessons, stats and mastery chart'
t.fill = TITLE_FILL; t.font = Font(color='FFFFFF', bold=True, size=12)
t.alignment = Alignment(horizontal='left', vertical='center'); mv.row_dimensions[1].height = 26
mv['A3'] = '▶ Module:'; mv['A3'].font = Font(bold=True, size=12)
mv.merge_cells('C3:F3')
pick = mv['C3']; pick.value = f'{modpairs[0][0]} — {modpairs[0][1]}'
pick.fill = PatternFill('solid', fgColor='FCE4D6'); pick.font = Font(bold=True, size=12)
pick.alignment = CEN; pick.border = BORDER
# Data-validation lists may not reference another sheet directly (Excel rejects
# it and the dropdown shows nothing) — route through a workbook-level named range.
wb.defined_names.add(DefinedName('ModuleList', attr_text=f'Dashboard!$A${MHR+1}:$A${mod_last}'))
dvm = DataValidation(type='list', formula1='=ModuleList', allow_blank=False, showErrorMessage=True)
dvm.prompt = 'Choose the module to inspect'; dvm.promptTitle = 'Module'
mv.add_data_validation(dvm); dvm.add('C3')
mv['N1'] = f'=IFERROR(INDEX(Dashboard!$B${MHR+1}:$B${mod_last},MATCH($C$3,Dashboard!$A${MHR+1}:$A${mod_last},0)),"")'
mv['N2'] = f'=IFERROR(INDEX(Dashboard!$C${MHR+1}:$C${mod_last},MATCH($C$3,Dashboard!$A${MHR+1}:$A${mod_last},0)),"")'
mv.column_dimensions['N'].hidden = True

def cifs(extra=''):
    return f"=COUNTIFS({L('B',3,LAST)},$N$1,{L('C',3,LAST)},$N$2{extra})"

stats = [('Total', cifs()), ('Started', cifs(f',{L("F",3,LAST)},"<>"')),
         ('Mastered', cifs(f',{L("G",3,LAST)},5')), ('Overdue', cifs(f',{L("J",3,LAST)},"*OVERDUE*"')),
         ('Due today', cifs(f',{L("J",3,LAST)},"*TODAY*"')), ('This week', cifs(f',{L("J",3,LAST)},"*THIS WEEK*"'))]
mv['A5'] = 'THIS MODULE'; mv['A5'].font = BOLD
for i, (lab, frm) in enumerate(stats):
    c = 1 + (i % 3) * 2; r = 6 + i // 3
    mv.cell(r, c, lab).fill = SUB_FILL; mv.cell(r, c, lab).border = BORDER
    vc = mv.cell(r, c + 1, frm); vc.font = BOLD; vc.alignment = CEN; vc.border = BORDER
mv['H5'] = 'Mastery'; mv['I5'] = 'Lessons'; style_header(mv, 5, 9)
for k in range(6):
    r = 6 + k
    mv.cell(r, 8, k).alignment = CEN; mv.cell(r, 8).border = BORDER
    mv.cell(r, 9, cifs(f',{L("G",3,LAST)},$H{r}')); mv.cell(r, 9).border = BORDER
cm = BarChart(); cm.type = 'col'; cm.title = 'Mastery in this module'; cm.legend = None
cm.add_data(Reference(mv, min_col=9, min_row=5, max_row=11), titles_from_data=True)
cm.set_categories(Reference(mv, min_col=8, min_row=6, max_row=11))
cm.height, cm.width = 6.5, 12; mv.add_chart(cm, 'K5')
mv['A13'] = 'LESSONS IN THIS MODULE'; mv['A13'].font = BOLD
for c, hh in enumerate(['Lesson', 'Last Review', 'Mastery', 'Interval', 'Next Review', 'Status', 'Priority'], 1):
    mv.cell(14, c, hh)
style_header(mv, 14, 7)
block = f"'lesson-database'!$E$3:$K${LAST}"
mv['A15'] = (f"=IFERROR(FILTER({block},({L('B',3,LAST)}=$N$1)*({L('C',3,LAST)}=$N$2),\"— no lessons —\"),"
             f"\"FILTER needs Excel 365 or Google Sheets — use the lesson-database tab's filter arrows instead.\")")
for col, w in {'A':50,'B':13,'C':9,'D':9,'E':13,'F':14,'G':9,'H':3,'I':9}.items():
    mv.column_dimensions[col].width = w
mv.freeze_panes = 'A15'

# ═══════════════ Weekly ═══════════════
wk.sheet_view.showGridLines = False
wk.merge_cells('A1:E1')
t = wk['A1']; t.value = '🗓️ WEEKLY — your next 7 days of reviews'
t.fill = TITLE_FILL; t.font = WHITE; t.alignment = Alignment(horizontal='left', vertical='center')
wk.row_dimensions[1].height = 24
for c, hh in enumerate(['Date', 'Day', 'Reviews', 'Load', 'Est. min'], 1):
    wk.cell(3, c, hh)
style_header(wk, 3, 5)
for i in range(7):
    r = 4 + i
    wk.cell(r, 1, '=TODAY()' if i == 0 else f'=A{r-1}+1'); wk.cell(r, 1).number_format = date_fmt
    wk.cell(r, 2, f'=TEXT(A{r},"ddd")').alignment = CEN
    if i == 0:
        wk.cell(r, 3, f'=COUNTIF({L("I",3,LAST)},A{r})+COUNTIF({L("J",3,LAST)},"*OVERDUE*")')
    else:
        wk.cell(r, 3, f'=COUNTIF({L("I",3,LAST)},A{r})')
    wk.cell(r, 3).alignment = CEN
    wk.cell(r, 4, f'=IF(C{r}=0,"🟢 Free",IF(C{r}<=5,"🟢 Light",IF(C{r}<=10,"🟡 Medium",IF(C{r}<=15,"🟠 Heavy","🔴 Overload"))))')
    wk.cell(r, 5, f'=C{r}*10').alignment = CEN
    for c in range(1, 6):
        wk.cell(r, c).border = BORDER
for col, w in {'A':13,'B':7,'C':10,'D':13,'E':10}.items():
    wk.column_dimensions[col].width = w
cw = BarChart(); cw.type = 'col'; cw.title = 'Reviews per day'; cw.legend = None
cw.add_data(Reference(wk, min_col=3, min_row=3, max_row=10), titles_from_data=True)
cw.set_categories(Reference(wk, min_col=2, min_row=4, max_row=10))
cw.height, cw.width = 7, 13; wk.add_chart(cw, 'G3')

# ═══════════════ How-To ═══════════════
ht.sheet_view.showGridLines = False
ht.column_dimensions['A'].width = 110
guide = [
    ('🧠 MEDICAL SRS — HOW TO USE', True),
    ('', False),
    ('THE ONE RULE: you only ever type in two columns on the lesson-database tab —', True),
    ('   • F  Last Review  → the day you studied it (press Ctrl+;  for today)', False),
    ('   • G  Mastery      → how well you recalled it, 0 (forgot) to 5 (mastered)', False),
    ('Everything else (interval, next review, status, priority) fills in by itself.', False),
    ('Top bar (row 1): today + days left until the exam — set your exam date in the blue K1 cell.', False),
    ('Rows are tinted per module (same color = same module) for easier scanning.', False),
    ('', False),
    ('HOW THE SCHEDULE WORKS', True),
    ('   Mastery 0→review in 1 day · 1→3d · 2→7d · 3→14d · 4→30d · 5→done (drops off).', False),
    ('   Forgot it? set 0 and it comes back tomorrow. Solid? a higher score pushes it out.', False),
    ('', False),
    ('FILTERING (quick filter)', True),
    ('   lesson-database tab → click the ▼ arrow on row 2 of Semester, Module or Subject.', False),
    ('   This is the fast, works-everywhere way to see one semester or one module.', False),
    ('', False),
    ('THE TABS', True),
    ("   • Dashboard   — bird's-eye view: totals, % done per semester & per module, charts.", False),
    ('   • Module View — pick one module from the dropdown to see its lessons + mastery chart.', False),
    ('   • Weekly      — how many reviews land on each of the next 7 days.', False),
    ('', False),
    ('ONLINE vs OFFLINE', True),
    ('   • Online  (Google Sheets): also auto-syncs reviews to Google Calendar via the 📚 SRS menu.', False),
    ('   • Offline (Microsoft Excel): everything works except calendar sync (Apps Script is Google-only).', False),
    ('     Tip: the Module View lesson list uses FILTER — needs Excel 365 or Google Sheets.', False),
]
for i, (txt, bold) in enumerate(guide):
    cell = ht.cell(i + 1, 1, txt)
    cell.font = Font(bold=bold, size=13 if i == 0 else 11)
    if i == 0:
        cell.fill = TITLE_FILL; cell.font = Font(color='FFFFFF', bold=True, size=13)

try:
    wb.calculation.fullCalcOnLoad = True
except Exception as e:
    print('calc flag skip:', e)
wb.active = 0
wb.save(OUT)
print(f'Saved {OUT} · {os.path.getsize(OUT)} bytes · sheets {wb.sheetnames} · data rows 3..{last_data}')
