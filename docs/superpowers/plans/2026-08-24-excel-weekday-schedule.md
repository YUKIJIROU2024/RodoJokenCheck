# Excel版 ①労働時間 曜日別所定労働時間 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a user set start time / end time / break minutes per weekday (月〜日) in `generate_check_sheet.py`'s generated Excel チェックシート, mirroring the Web版 feature already merged in `index.html` — expressed as Excel formulas (no VBA/macros).

**Architecture:** `generate_check_sheet.py` is a single Python script (no build step, no test framework) that programmatically builds an `.xlsx` via `openpyxl`, using a `SheetBuilder` helper (`b`) whose `inp()`/`cb()`/`judgment_row()` methods assign each new named value the next physical spreadsheet row and let later formulas reference it by name via `b.ref(name)` (same-sheet) or `xref(b, name)` (cross-sheet, from the チェック結果 summary sheet). Because Excel can't dynamically show/hide rows without macros, this plan always displays a 7-row per-weekday input table (matching the existing always-visible `wd{i}`＝休日 checkboxes pattern) alongside a `useWeekday` toggle cell that controls which values the FORMULAS consume — the toggle governs computation, not visibility. A single per-day cell `dayMin{i}` (i=0..6, 0=月..6=日) is the one place that decides "common schedule or this weekday's own schedule," mirroring the Web版's `getDaySchedule(i)` accessor. Every aggregate (`weeklyH`, `maxDailyH`, ③'s per-day break check) is built from `dayMin{i}`, never from `startTime`/`endTime`/`breakTime` or `dayStart{i}`/`dayEnd{i}`/`dayBreak{i}` directly.

**Tech Stack:** Python 3 + `openpyxl` (already a dependency of this script — no new dependencies). Verification uses `libreoffice`/`soffice --headless --convert-to csv` (already installed in this environment) to force-recalculate the generated `.xlsx` and read the computed values back as CSV — there is no automated test framework in this repo, so this is the equivalent of driving a real browser for the Web版's Playwright-based verification. `openpyxl` is also used directly (loading the generated file, mutating input cells, saving a scratch copy) to script different input scenarios before each recalculation pass.

**Spec:** `docs/superpowers/specs/2026-08-24-excel-weekday-schedule-design.md`

## Global Constraints

- Weekday index convention: 0=月, 1=火, 2=水, 3=木, 4=金, 5=土, 6=日 (matches the existing `wd0`..`wd6` input cells and the existing `wd_labels = ['月','火','水','木','金','土','日']` local variable already declared in `build_check_sheet` right before the `wd0`..`wd6` loop).
- New cell names: `useWeekday` (`はい`/`いいえ`, default `いいえ`). `dayStart{i}`/`dayEnd{i}`/`dayBreak{i}` for i=0..6 (time/time/number inputs, default 9/24, 18/24, 60 — matching the existing common `startTime`/`endTime`/`breakTime` defaults). `dayMin{i}` for i=0..6 (computed, minutes).
- Every aggregate formula (`weeklyH`, `maxDailyH`, ③3-1's per-day pass check) must read a day's effective minutes only via `dayMin{i}` — never by referencing `startTime`/`endTime`/`breakTime`/`dayStart{i}`/`dayEnd{i}`/`dayBreak{i}` directly.
- `dayMin{i}`'s own formula is the only place that branches on `useWeekday`: `=IF(useWeekday="はい", <that day's own start/end/break calculation>, dailyMin)` (`dailyMin` is the existing common-mode cell, left untouched).
- 1-1（8時間超過）must flag ✗ if ANY work day (`wd{i}="いいえ"`) has `dayMin{i}/60 > 8`, and the detail text must name every offending weekday via `TEXTJOIN`, not just one.
- 1-2（週40/44時間）must sum `dayMin{i}` across work days only (`(wd{i}="いいえ")*dayMin{i}`), not one common value × day count.
- 1-3（年間総労働時間）and ⑧最低賃金（日給制）must share the same average-daily-hours basis: `dailyH = IF(weeklyWorkDays>0, weeklyH/weeklyWorkDays, 0)`. `dailyH`'s cell must be the ONLY place this average is computed — `annualWorkH` and ⑧'s `hourlyRateBasic` formulas must NOT be edited; they already reference `dailyH` by name and inherit the new meaning automatically.
- ③3-1（休憩時間の長さ）must check each work day's own required break (>480min→60min, >360min→45min, else 0) against that day's own effective break minutes (`IF(useWeekday="はい",dayBreak{i},breakTime)`), flag ✗ if ANY work day fails, and name every offending weekday via `TEXTJOIN`.
- TEXTJOIN is acceptable (Excel 2019/365+, no compatibility fallback needed — confirmed with the user).
- Common-mode regression (all judgments identical to the pre-feature script's output) when `useWeekday="いいえ"` must hold algebraically, not just "usually" — every task below states why.
- Excel-version support (`generate_check_sheet.py`) only. `index.html` is out of scope for this plan (already done, merged).

## Verification Toolkit (used by every task below)

There is no test framework in this repo. Verification means: (1) run `python3 generate_check_sheet.py` to regenerate the `.xlsx` from the modified script, (2) optionally mutate specific input cells with `openpyxl` to set up a scenario, (3) force LibreOffice to recalculate and export computed values as CSV, (4) read the CSV.

**Command to force-recalculate and export a sheet as CSV** (the exported sheet is whichever one is FIRST in the workbook's sheet order — `wb.sheetnames[0]` — not necessarily whichever is "active"):

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir <outdir> <path-to-xlsx>
```

**To inspect the チェック結果 (results) summary sheet** (already first in the workbook by construction — `wb.create_sheet('チェック結果', 0)` in `main()`): run the command above directly on the generated file. The output CSV's rows are `No.,チェック項目,判定,詳細` — this is what most task verifications below read.

**To inspect the チェックシート (detail) sheet directly** (e.g. to read a `dayMin{i}` cell's raw computed value, not surfaced in チェック結果): make a scratch copy with `チェックシート` moved to position 0 first:

```python
import openpyxl
wb = openpyxl.load_workbook('rodo-joken-check.xlsx')
wb.move_sheet('チェックシート', offset=-wb.sheetnames.index('チェックシート'))
wb.save('/path/to/scratch.xlsx')
```
then run the `soffice` command above on `scratch.xlsx`.

**To set up a test scenario** (e.g. weekday mode, 土曜3時間), mutate input cells directly with `openpyxl` before saving the scratch copy — input cells are plain `.value` assignments at the row `SheetBuilder` assigned them (readable from the generated file by searching column A for the label text, or by counting rows in the known layout each task documents). Every task below gives the exact scenario as a self-contained Python snippet.

All scratch files belong in the repo-appropriate scratch directory for this session — do not commit generated `.xlsx`/`.csv` scratch files (the repo's `.gitignore` already excludes `*.xlsx`).

---

### Task 1: Data layer — weekday toggle + per-day inputs + `dayMin{i}` (additive only)

**Files:**
- Modify: `generate_check_sheet.py` (inside `build_check_sheet`, right after the existing `dailyH` cb call, currently ending `narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間"')`, and right before the existing `wd_labels = ['月', '火', '水', '木', '金', '土', '日']` line)

**Interfaces:**
- Produces: cell names `useWeekday`, `dayStart{i}`/`dayEnd{i}`/`dayBreak{i}`/`dayMin{i}` for i=0..6, all reachable via `b.ref(name)` for the rest of `build_check_sheet`.
- Consumes: existing `dailyMin` cell (already defined earlier in the function, used as the common-mode fallback inside `dayMin{i}`'s formula).

This task is purely additive: the existing `dailyH`, `wd0`..`wd6`, `weeklyH`, and all judgment formulas (1-1, 1-2, 3-1, ⑧) are left completely untouched. Nothing downstream reads the new cells yet, so this task cannot change any existing computed value — its own success bar is "the generated file still produces byte-identical チェック結果 output to before, plus the new cells compute sensible values when inspected directly and when `useWeekday` is manually flipped."

- [ ] **Step 1: Insert the new cells**

In `generate_check_sheet.py`, find (the `dailyH` block followed by the `wd_labels` line):

```python
    b.reserve('dailyH')
    b.cb('dailyH', '1日の所定労働時間（時間）', f'={b.ref("dailyMin")}/60', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間"')

    wd_labels = ['月', '火', '水', '木', '金', '土', '日']
```

Replace with:

```python
    b.reserve('dailyH')
    b.cb('dailyH', '1日の所定労働時間（時間）', f'={b.ref("dailyMin")}/60', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間"')

    b.inp('useWeekday', '曜日ごとに設定する', 'いいえ', choices=['はい', 'いいえ'],
          note='「はい」で下の曜日別入力を使用します（未入力の間も安全に共通値へフォールバック）')

    wd_sched_labels = ['月', '火', '水', '木', '金', '土', '日']
    for idx, day in enumerate(wd_sched_labels):
        b.inp(f'dayStart{idx}', f'{day}曜日 始業時刻', 9/24, fmt='h:mm')
        b.inp(f'dayEnd{idx}', f'{day}曜日 終業時刻', 18/24, fmt='h:mm')
        b.inp(f'dayBreak{idx}', f'{day}曜日 休憩時間（分）', 60)
        b.reserve(f'dayMin{idx}')
        b.cb(f'dayMin{idx}', f'{day}曜日 有効な所定労働時間（分）',
             f'=IF({b.ref("useWeekday")}="はい",'
             f'MAX(0,(IF({b.ref(f"dayEnd{idx}")}<{b.ref(f"dayStart{idx}")},{b.ref(f"dayEnd{idx}")}+1,{b.ref(f"dayEnd{idx}")})-{b.ref(f"dayStart{idx}")})*24*60-{b.ref(f"dayBreak{idx}")}),'
             f'{b.ref("dailyMin")})', fmt='0',
             narration=f'=IF({b.ref("useWeekday")}="はい","曜日別: "&TEXT({b.ref(f"dayMin{idx}")}/60,"0.00")&"時間","共通値を使用")')

    wd_labels = ['月', '火', '水', '木', '金', '土', '日']
```

(`wd_sched_labels` is a separate local variable from the pre-existing `wd_labels` — at this point in the function, `wd_labels` has not been declared yet, so the new loop cannot reuse it. `wd_labels` itself remains unchanged, declared right after, exactly as before.)

- [ ] **Step 2: Verify — regenerate and confirm zero output change**

```bash
cd /mnt/c/Users/y_uej/CC-RodoJokenCheck
python3 generate_check_sheet.py
```

Expected: prints `生成完了: ...（項目数: 22）` (same judgment count as before this task — this task adds no new judgment rows).

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify rodo-joken-check.xlsx
cat /tmp/excel-verify/rodo-joken-check.csv
```

Expected: every row's 判定/詳細 for 1-1, 1-2, 1-3, 3-1, 8-1 identical to a run of the SAME command taken before this task's edit (capture that baseline first, before making the change, and diff against it).

- [ ] **Step 3: Verify — new cells compute correctly in isolation**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')
wb.move_sheet('チェックシート', offset=-wb.sheetnames.index('チェックシート'))
wb.save('/tmp/excel-verify/task1_scratch.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task1_scratch.xlsx
grep -A2 "曜日ごとに設定する" /tmp/excel-verify/task1_scratch.csv
grep "曜日 有効な所定労働時間" /tmp/excel-verify/task1_scratch.csv
```

Expected: `useWeekday` shows `いいえ`; all 7 `dayMin{i}` rows show `480`（分, matching the common `dailyMin`=480 since default day inputs are 09:00-18:00/60分 same as common defaults）.

Now flip `useWeekday` to `はい` and give 土 (day index 5) a distinct schedule:

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '土曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '土曜日 終業時刻'), column=2).value = 12/24
ws.cell(row=find_row(ws, '土曜日 休憩時間（分）'), column=2).value = 0

wb.move_sheet('チェックシート', offset=-wb.sheetnames.index('チェックシート'))
wb.save('/tmp/excel-verify/task1_scratch2.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task1_scratch2.xlsx
grep "曜日 有効な所定労働時間" /tmp/excel-verify/task1_scratch2.csv
```

Expected: 土曜日 有効な所定労働時間（分）＝`180`（09:00〜12:00・休憩0分＝3時間）, all other 6 days still `480`（未変更の共通値相当のまま）.

- [ ] **Step 4: Commit**

```bash
git add generate_check_sheet.py
git commit -m "$(cat <<'EOF'
①労働時間: 曜日別スケジュールの入力セルとdayMin{i}を追加（Excel版）

useWeekday切替セルと曜日ごとのdayStart/dayEnd/dayBreak/dayMin{i}
（0=月〜6=日）を追加。dayMin{i}はWeb版のgetDaySchedule(i)に相当する
単一の判断ポイントで、切替が「いいえ」の間は共通のdailyMinへ安全に
フォールバックする。既存のdailyH・wd{i}・週次/年次集計・各判定は
まだ変更しておらず、既存の判定結果に影響はない（後続タスクで順次移行）。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 2: `maxDailyH` + wire ①1-1 onto per-day data

**Files:**
- Modify: `generate_check_sheet.py` (insert after the existing `weeklyH` cb call, before `holidayType`; replace the `j1_1` judgment_row)

**Interfaces:**
- Consumes: `dayMin{i}` (Task 1), existing `wd{i}` cells, existing `wd_labels` local variable (already declared earlier in `build_check_sheet`, at its original position right before the `wd0`..`wd6` loop).
- Produces: cell name `maxDailyH`.

- [ ] **Step 1: Add `maxDailyH` after `weeklyH`**

In `generate_check_sheet.py`, find:

```python
    b.reserve('weeklyH')
    b.cb('weeklyH', '週の所定労働時間（時間）', f'={b.ref("dailyH")}*{b.ref("weeklyWorkDays")}', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間/日 × "&{b.ref("weeklyWorkDays")}&"日 = "&TEXT({b.ref("weeklyH")},"0.00")&"時間/週"')

    b.inp('holidayType', '祝日の扱い', '休日扱い', choices=['休日扱い', '出勤日'])
```

Replace with:

```python
    b.reserve('weeklyH')
    b.cb('weeklyH', '週の所定労働時間（時間）', f'={b.ref("dailyH")}*{b.ref("weeklyWorkDays")}', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間/日 × "&{b.ref("weeklyWorkDays")}&"日 = "&TEXT({b.ref("weeklyH")},"0.00")&"時間/週"')

    max_parts = ','.join(
        f'({b.ref(f"wd{d}")}="いいえ")*{b.ref(f"dayMin{d}")}' for d in range(7)
    )
    b.reserve('maxDailyH')
    b.cb('maxDailyH', '就業日の中の最大1日所定労働時間（時間）',
         f'=MAX({max_parts})/60', fmt='0.00',
         narration=f'="就業日の最大: "&TEXT({b.ref("maxDailyH")},"0.00")&"時間"')

    b.inp('holidayType', '祝日の扱い', '休日扱い', choices=['休日扱い', '出勤日'])
```

(`weeklyH`'s own block is untouched here — Task 3 edits its formula separately. This insertion only adds `maxDailyH` after it.)

- [ ] **Step 2: Rewire j1_1 onto `maxDailyH`**

In `generate_check_sheet.py`, find:

```python
    b.judgment_row('j1_1', '1-1', '1日の所定労働時間',
        f'=IF({b.ref("dailyH")}<=8,"○","✗")',
        detail=f'="1日労働時間: "&TEXT({xref(b,"dailyH")},"0.00")&"時間（上限8時間）"')
```

Replace with:

```python
    over_parts = ','.join(
        f'IF(AND({xref(b,f"wd{d}")}="いいえ",{xref(b,f"dayMin{d}")}/60>8),"{wd_labels[d]}","")'
        for d in range(7)
    )
    b.judgment_row('j1_1', '1-1', '1日の所定労働時間',
        f'=IF({b.ref("maxDailyH")}<=8,"○","✗")',
        detail=f'=IF({xref(b,"maxDailyH")}<=8,'
               f'"最大1日労働時間: "&TEXT({xref(b,"maxDailyH")},"0.00")&"時間（上限8時間）",'
               f'"8時間超過曜日: "&TEXTJOIN("・",TRUE,{over_parts})&"曜日")')
```

- [ ] **Step 3: Verify — regression (common mode unchanged)**

```bash
cd /mnt/c/Users/y_uej/CC-RodoJokenCheck
python3 generate_check_sheet.py
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify rodo-joken-check.xlsx
grep "^1-1," /tmp/excel-verify/rodo-joken-check.csv
```

Expected: `1-1,1日の所定労働時間,○,最大1日労働時間: 8.00時間（上限8時間）` — same judgment (○) and same 8.00時間 value as before this task; only the detail wording changed from "1日労働時間" to "最大1日労働時間" (semantically identical in common mode).

- [ ] **Step 4: Verify — weekday mode, one day over 8 hours**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '土曜日（休日）'), column=2).value = 'いいえ'  # 土を就業日にする
ws.cell(row=find_row(ws, '土曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '土曜日 終業時刻'), column=2).value = 12/24
ws.cell(row=find_row(ws, '土曜日 休憩時間（分）'), column=2).value = 0
ws.cell(row=find_row(ws, '月曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '月曜日 終業時刻'), column=2).value = 19/24  # 9時間勤務
ws.cell(row=find_row(ws, '月曜日 休憩時間（分）'), column=2).value = 60

wb.save('/tmp/excel-verify/task2_scratch.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task2_scratch.xlsx
grep "^1-1," /tmp/excel-verify/task2_scratch.csv
```

Expected: `1-1,1日の所定労働時間,✗,8時間超過曜日: 月曜日` — 月 named specifically, 土(3時間)は超過に含まれない。

- [ ] **Step 5: Commit**

```bash
git add generate_check_sheet.py
git commit -m "$(cat <<'EOF'
①1-1: 曜日ごとの最大労働時間で8時間超過を判定（Excel版）

maxDailyHを追加し、就業日のdayMin{i}の最大値と8時間を比較するよう
1-1を変更。詳細欄はTEXTJOINで超過している曜日をすべて列挙する。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 3: Redefine `weeklyH` to sum per-day minutes (wires ①1-2 automatically)

**Files:**
- Modify: `generate_check_sheet.py` (the `weeklyH` cb block)

**Interfaces:**
- Consumes: `dayMin{i}` (Task 1), existing `wd{i}` cells.
- Produces: `weeklyH` (same cell name, redefined formula — j1_2's judgment/detail formulas are untouched and pick up the new meaning automatically since they reference `weeklyH` by name).

- [ ] **Step 1: Redefine `weeklyH`**

In `generate_check_sheet.py`, find:

```python
    b.reserve('weeklyH')
    b.cb('weeklyH', '週の所定労働時間（時間）', f'={b.ref("dailyH")}*{b.ref("weeklyWorkDays")}', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間/日 × "&{b.ref("weeklyWorkDays")}&"日 = "&TEXT({b.ref("weeklyH")},"0.00")&"時間/週"')
```

Replace with:

```python
    weekly_sum_parts = '+'.join(
        f'({b.ref(f"wd{d}")}="いいえ")*{b.ref(f"dayMin{d}")}' for d in range(7)
    )
    b.reserve('weeklyH')
    b.cb('weeklyH', '週の所定労働時間（時間）', f'=({weekly_sum_parts})/60', fmt='0.00',
         narration=f'="就業日の所定労働時間合計 ÷ 60 = "&TEXT({b.ref("weeklyH")},"0.00")&"時間/週"')
```

- [ ] **Step 2: Verify — regression (common mode unchanged)**

```bash
cd /mnt/c/Users/y_uej/CC-RodoJokenCheck
python3 generate_check_sheet.py
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify rodo-joken-check.xlsx
grep "^1-2," /tmp/excel-verify/rodo-joken-check.csv
```

Expected: `1-2,週の所定労働時間,○,週労働時間: 40.00時間（上限40時間）` — identical to before this task (5 work days × 8h = 40h either way, per the algebraic equivalence when `useWeekday="いいえ"`).

- [ ] **Step 3: Verify — weekday mode, partial short day**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '土曜日（休日）'), column=2).value = 'いいえ'
ws.cell(row=find_row(ws, '土曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '土曜日 終業時刻'), column=2).value = 12/24
ws.cell(row=find_row(ws, '土曜日 休憩時間（分）'), column=2).value = 0

wb.save('/tmp/excel-verify/task3_scratch.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task3_scratch.xlsx
grep "^1-2," /tmp/excel-verify/task3_scratch.csv
```

Expected: `1-2,週の所定労働時間,✗,週労働時間: 43.00時間（上限40時間）` — 月〜金5日×8h(=40h) ＋ 土3h = 43h（4-1で土が休日から就業日に変わったことで週休1日となり、40時間超過で✗になる）。

- [ ] **Step 4: Commit**

```bash
git add generate_check_sheet.py
git commit -m "$(cat <<'EOF'
①1-2: 週の所定労働時間を曜日ごとの合計に変更（Excel版）

weeklyHの数式を「共通値×就業日数」から、就業日のdayMin{i}合計に変更。
1-2の判定式・詳細欄は既存のまま（weeklyH参照のみ）で新しい挙動を継承する。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 4: Move + redefine `dailyH` as the weekly average (wires ①1-3 and ⑧ automatically)

**Files:**
- Modify: `generate_check_sheet.py` (delete `dailyH`'s original block, re-insert a redefined version after `maxDailyH`)

**Interfaces:**
- Consumes: `weeklyH` (Task 3), existing `weeklyWorkDays` cell.
- Produces: `dailyH` (same cell name, moved position and redefined formula — `annualWorkH`（1-3で使用）と⑧の`hourlyRateBasic`（日給制）は共に`dailyH`を参照済みのため無改修で新しい平均値を継承する).

`dailyH`'s new formula depends on `weeklyH`, which (after Task 3) depends on `dayMin{i}`/`wd{i}`, all of which are defined LATER in the file than `dailyH`'s ORIGINAL position (right after `dailyMin`). `SheetBuilder` rows are assigned in code-execution order, so `dailyH`'s definition must physically move to after `weeklyH`/`maxDailyH` for `b.ref("weeklyH")` to resolve — this task deletes it from its old position and re-creates it in the new one.

- [ ] **Step 1: Delete `dailyH` from its original position**

In `generate_check_sheet.py`, find:

```python
    b.reserve('dailyH')
    b.cb('dailyH', '1日の所定労働時間（時間）', f'={b.ref("dailyMin")}/60', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間"')

    b.inp('useWeekday', '曜日ごとに設定する', 'いいえ', choices=['はい', 'いいえ'],
```

Replace with:

```python
    b.inp('useWeekday', '曜日ごとに設定する', 'いいえ', choices=['はい', 'いいえ'],
```

- [ ] **Step 2: Re-insert `dailyH` after `maxDailyH`, redefined as the weekly average**

In `generate_check_sheet.py`, find:

```python
    b.reserve('maxDailyH')
    b.cb('maxDailyH', '就業日の中の最大1日所定労働時間（時間）',
         f'=MAX({max_parts})/60', fmt='0.00',
         narration=f'="就業日の最大: "&TEXT({b.ref("maxDailyH")},"0.00")&"時間"')

    b.inp('holidayType', '祝日の扱い', '休日扱い', choices=['休日扱い', '出勤日'])
```

Replace with:

```python
    b.reserve('maxDailyH')
    b.cb('maxDailyH', '就業日の中の最大1日所定労働時間（時間）',
         f'=MAX({max_parts})/60', fmt='0.00',
         narration=f'="就業日の最大: "&TEXT({b.ref("maxDailyH")},"0.00")&"時間"')

    b.reserve('dailyH')
    b.cb('dailyH', '1日の所定労働時間（平均・時間）',
         f'=IF({b.ref("weeklyWorkDays")}>0,{b.ref("weeklyH")}/{b.ref("weeklyWorkDays")},0)', fmt='0.00',
         narration=f'=IF({b.ref("weeklyWorkDays")}>0,'
                    f'TEXT({b.ref("weeklyH")},"0.00")&"時間/週 ÷ "&{b.ref("weeklyWorkDays")}&"日 = "&TEXT({b.ref("dailyH")},"0.00")&"時間/日（平均）",'
                    f'"就業日なし")')

    b.inp('holidayType', '祝日の扱い', '休日扱い', choices=['休日扱い', '出勤日'])
```

- [ ] **Step 3: Verify — regression (common mode unchanged)**

```bash
cd /mnt/c/Users/y_uej/CC-RodoJokenCheck
python3 generate_check_sheet.py
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify rodo-joken-check.xlsx
grep "^1-3,\|^8-1," /tmp/excel-verify/rodo-joken-check.csv
```

Expected: `1-3` shows the same 年間労働 value as before this task (`244日 / 1952.0時間` — unchanged since common-mode average equals the single common value). `8-1` shows `△`（最低賃金・日給/月給が未入力のデフォルト状態のまま — this task doesn't change that, only its internal basis）.

- [ ] **Step 4: Verify — weekday mode, 1-3 and ⑧ share the same average**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '土曜日（休日）'), column=2).value = 'いいえ'
ws.cell(row=find_row(ws, '土曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '土曜日 終業時刻'), column=2).value = 12/24
ws.cell(row=find_row(ws, '土曜日 休憩時間（分）'), column=2).value = 0
ws.cell(row=find_row(ws, '賃金の種類'), column=2).value = '日給制'
ws.cell(row=find_row(ws, '日給（円）'), column=2).value = 9000

wb.save('/tmp/excel-verify/task4_scratch.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task4_scratch.xlsx
grep "^1-3,\|^8-1," /tmp/excel-verify/task4_scratch.csv
```

Expected: 1-3の詳細に含まれる「年間労働」の元になる平均時間（週43.00時間÷6日＝7.17時間/日）と、8-1の詳細に含まれる時間換算額（9000円÷7.17時間）が、どちらも同じ7.17時間ベースであること — 手計算 9000/7.1667≒1255.81 と一致することを確認する。

- [ ] **Step 5: Commit**

```bash
git add generate_check_sheet.py
git commit -m "$(cat <<'EOF'
①1-3・⑧: 1日の所定労働時間を週平均ベースに変更（Excel版）

dailyHセルをweeklyH/weeklyWorkDaysの後ろへ移動し、数式を「共通1日値」
から「週合計÷週就業日数」の平均値に変更。annualWorkH（1-3）と
hourlyRateBasic（⑧日給制）はdailyHを参照済みのため無改修で
平均ベースを継承する（Web版Task4と同じ設計）。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 5: Wire ③3-1 onto per-day break requirements

**Files:**
- Modify: `generate_check_sheet.py` (the ③ section: remove `breakRequiredMin`, rewrite `j3_1`)

**Interfaces:**
- Consumes: `dayMin{i}` (Task 1), existing `wd{i}`/`useWeekday`/`dayBreak{i}`/`breakTime` cells, existing `wd_labels` local variable.
- Produces: same judgment names `j3_1`/`j3_2` — `j3_2`'s formula is untouched.

- [ ] **Step 1: Replace `breakRequiredMin`/`j3_1`**

In `generate_check_sheet.py`, find the full ③ section:

```python
    b.section_hdr('③ 休憩')
    b.inp('breakException', '一斉付与 適用除外（労使協定）', 'なし', choices=['なし', 'あり'])
    b.reserve('breakRequiredMin')
    b.cb('breakRequiredMin', '必要休憩時間（分）',
         f'=IF({b.ref("dailyMin")}>480,60,IF({b.ref("dailyMin")}>360,45,0))', fmt='0',
         narration=f'=IF({b.ref("breakRequiredMin")}=0,"6時間以下 → 休憩付与義務なし",'
                    f'IF({b.ref("dailyMin")}>480,"8時間超 → 60分以上必要","6時間超8時間以下 → 45分以上必要"))')

    b.judgment_row('j3_1', '3-1', '休憩時間の長さ',
        f'=IF({b.ref("breakRequiredMin")}=0,"○",IF({b.ref("breakTime")}>={b.ref("breakRequiredMin")},"○","✗"))',
        detail=f'="設定休憩: "&{xref(b,"breakTime")}&"分 / 必要休憩: "&{xref(b,"breakRequiredMin")}&"分"')
    b.judgment_row('j3_2', '3-2', '休憩の一斉付与', '="○"',
        detail=f'="適用除外: "&{xref(b,"breakException")}')
```

Replace with:

```python
    b.section_hdr('③ 休憩')
    b.inp('breakException', '一斉付与 適用除外（労使協定）', 'なし', choices=['なし', 'あり'])

    day_pass_parts = [
        f'OR({b.ref(f"wd{d}")}="はい",'
        f'IF({b.ref("useWeekday")}="はい",{b.ref(f"dayBreak{d}")},{b.ref("breakTime")})'
        f'>=IF({b.ref(f"dayMin{d}")}>480,60,IF({b.ref(f"dayMin{d}")}>360,45,0)))'
        for d in range(7)
    ]
    fail_parts = ','.join(
        f'IF(AND({xref(b,f"wd{d}")}<>"はい",'
        f'IF({xref(b,"useWeekday")}="はい",{xref(b,f"dayBreak{d}")},{xref(b,"breakTime")})'
        f'<IF({xref(b,f"dayMin{d}")}>480,60,IF({xref(b,f"dayMin{d}")}>360,45,0))),'
        f'"{wd_labels[d]}","")'
        for d in range(7)
    )
    b.judgment_row('j3_1', '3-1', '休憩時間の長さ',
        f'=IF(AND({",".join(day_pass_parts)}),"○","✗")',
        detail=f'=IF({xref(b,"j3_1")}="○","各曜日の必要休憩を満たしています",'
               f'"休憩不足曜日: "&TEXTJOIN("・",TRUE,{fail_parts})&"曜日")')
    b.judgment_row('j3_2', '3-2', '休憩の一斉付与', '="○"',
        detail=f'="適用除外: "&{xref(b,"breakException")}')
```

(`breakRequiredMin` is removed entirely — its only consumer was `j3_1`, now rewritten to check each work day individually. `j3_2` is copied verbatim, unchanged.)

- [ ] **Step 2: Verify — regression (common mode unchanged)**

```bash
cd /mnt/c/Users/y_uej/CC-RodoJokenCheck
python3 generate_check_sheet.py
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify rodo-joken-check.xlsx
grep "^3-1," /tmp/excel-verify/rodo-joken-check.csv
```

Expected: `3-1,休憩時間の長さ,○,各曜日の必要休憩を満たしています` — 判定は○のまま（共通モードでは全曜日が同じdailyMin/breakTimeを参照するため、8時間労働・60分休憩で要件を満たす）。

- [ ] **Step 3: Verify — weekday mode, one day's break is insufficient**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '月曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '月曜日 終業時刻'), column=2).value = 19/24  # 9時間拘束
ws.cell(row=find_row(ws, '月曜日 休憩時間（分）'), column=2).value = 30  # 60分未満で不足

wb.save('/tmp/excel-verify/task5_scratch.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task5_scratch.xlsx
grep "^3-1," /tmp/excel-verify/task5_scratch.csv
```

Expected: `3-1,休憩時間の長さ,✗,休憩不足曜日: 月曜日` — 月（8.5時間労働・休憩30分＜必要60分）だけが不足として列挙される。

- [ ] **Step 4: Commit**

```bash
git add generate_check_sheet.py
git commit -m "$(cat <<'EOF'
③3-1: 休憩時間の長さを曜日ごとの判定に変更（Excel版）

各就業日のdayMin{i}から必要休憩を求め、その日の有効な休憩時間
（useWeekdayに応じてdayBreak{i}またはbreakTime）と比較。1日でも
不足があれば✗とし、詳細欄はTEXTJOINで不足曜日を列挙する。
共通モードのみで使われていたbreakRequiredMinは不要になったため削除。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 6: Full regression + edge-case verification pass

**Files:** none (verification only; fix inline if something's off, then commit any fix)

This task walks every item in the spec's テスト観点 list end-to-end against the finished feature, using the LibreOffice + openpyxl verification toolkit established above.

- [ ] **Step 1: Common mode parity**

```bash
cd /mnt/c/Users/y_uej/CC-RodoJokenCheck
python3 generate_check_sheet.py
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify rodo-joken-check.xlsx
grep "^1-1,\|^1-2,\|^1-3,\|^3-1,\|^8-1," /tmp/excel-verify/rodo-joken-check.csv
```

Expected: identical to the baseline captured before Task 1 started (1-1 ○ 8.00時間, 1-2 ○ 40.00時間, 1-3 244日/1952.0時間, 3-1 ○, 8-1 △（未入力）).

- [ ] **Step 2: All-days-same weekday mode ≡ common mode**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
# 全曜日デフォルト値（09:00-18:00/60分）のまま — 何も変更しない

wb.save('/tmp/excel-verify/task6_step2.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task6_step2.xlsx
grep "^1-1,\|^1-2,\|^1-3,\|^3-1," /tmp/excel-verify/task6_step2.csv
```

Expected: Step 1と数値が完全一致（1-1 ○ 8.00時間・1-2 ○ 40.00時間・1-3 1952.0時間・3-1 ○）。

- [ ] **Step 3: Partial short day + over-8h day combined (曜日別モードの本領)**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '土曜日（休日）'), column=2).value = 'いいえ'
ws.cell(row=find_row(ws, '土曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '土曜日 終業時刻'), column=2).value = 12/24
ws.cell(row=find_row(ws, '土曜日 休憩時間（分）'), column=2).value = 0
ws.cell(row=find_row(ws, '月曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '月曜日 終業時刻'), column=2).value = 19/24
ws.cell(row=find_row(ws, '月曜日 休憩時間（分）'), column=2).value = 30

wb.save('/tmp/excel-verify/task6_step3.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task6_step3.xlsx
grep "^1-1,\|^1-2,\|^3-1," /tmp/excel-verify/task6_step3.csv
```

Expected: `1-1,...,✗,8時間超過曜日: 月曜日`（土は3時間で超過なし）、`1-2,...,✗,週労働時間: 44.50時間（上限40時間）`（月9.5h+火水木金各8h+土3h＝44.5h）、`3-1,...,✗,休憩不足曜日: 月曜日`（火〜土は要件を満たす）。

- [ ] **Step 4: Holiday day excluded from all three aggregates**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
ws.cell(row=find_row(ws, '曜日ごとに設定する'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '水曜日（休日）'), column=2).value = 'はい'  # 水を休日にする
ws.cell(row=find_row(ws, '水曜日 始業時刻'), column=2).value = 9/24
ws.cell(row=find_row(ws, '水曜日 終業時刻'), column=2).value = 22/24  # 13時間・大きな値だが休日なので集計に影響しないはず
ws.cell(row=find_row(ws, '水曜日 休憩時間（分）'), column=2).value = 0

wb.save('/tmp/excel-verify/task6_step4.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task6_step4.xlsx
grep "^1-1,\|^1-2,\|^3-1," /tmp/excel-verify/task6_step4.csv
```

Expected: 水曜日を休日にしたことで残る就業日は月火木金の4日（各8h）＝32.00時間。`1-1,...,○`（水の13時間は休日扱いのため最大値の集計に含まれず、月〜金・土の他の値がいずれも8時間以下なら○）、`1-2,...,○,週労働時間: 32.00時間（上限40時間）`、`3-1,...,○`（水の休憩0分は休日のため`OR(wd="はい",...)`で自動的にパス扱い）。

- [ ] **Step 5: All-days-holiday edge case**

```python
import openpyxl

wb = openpyxl.load_workbook('/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx')

def find_row(ws, label):
    for row in ws.iter_rows():
        if row[0].value == label:
            return row[0].row
    raise ValueError(label)

ws = wb['チェックシート']
for day in ['月', '火', '水', '木', '金', '土', '日']:
    ws.cell(row=find_row(ws, f'{day}曜日（休日）'), column=2).value = 'はい'
ws.cell(row=find_row(ws, '賃金の種類'), column=2).value = '日給制'
ws.cell(row=find_row(ws, '日給（円）'), column=2).value = 9000

wb.save('/tmp/excel-verify/task6_step5.xlsx')
```

```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,0,1,,0,false,true,false,false,false --outdir /tmp/excel-verify /tmp/excel-verify/task6_step5.xlsx
grep "^1-1,\|^1-2,\|^3-1,\|^8-1," /tmp/excel-verify/task6_step5.csv
```

Expected: `1-1,...,○`（就業日なし→maxDailyH=0）、`1-2,...,○,週労働時間: 0.00時間`、`3-1,...,○,各曜日の必要休憩を満たしています`、`8-1,...,△`（既存の`dailyH>0`ガードにより最低賃金比較不能扱い）。エラー値（`#DIV/0!`等）がどのセルにも出ていないことを確認する。1-1/1-2/3-1が○なのに8-1だけ△になる不整合はWeb版でも既知・容認済みの挙動（悪化ではない）。

- [ ] **Step 6: Fix anything found, or confirm clean**

If any of Steps 1-5 surfaced a discrepancy, fix it in `generate_check_sheet.py` now and re-run the specific step that failed until it passes. If everything passed, no code change is needed for this task.

- [ ] **Step 7: Commit (only if Step 6 made a fix)**

If Step 6 found nothing to fix, skip this step — there is nothing to commit.

Otherwise, commit with a message whose body states which of Steps 1-5 failed and what the actual code change was:

```bash
git add generate_check_sheet.py
git commit -m "$(cat <<'EOF'
①③⑧: 曜日別スケジュールの回帰確認で見つかった不具合を修正（Excel版）

<Step 6で実際に直した内容を1-2文で、失敗したステップ番号と症状・原因を明記して書く>

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```
