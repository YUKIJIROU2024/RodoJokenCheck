# ①労働時間 曜日別所定労働時間 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a user set start time / end time / break minutes per weekday (月〜日) in `index.html`'s ①労働時間 tab, instead of one common set for all days, and have judgments 1-1, 1-2, 1-3, 3-1, and ⑧ reflect it correctly.

**Architecture:** `index.html` is a single static file with no build step, no bundler, and no test framework — all HTML/CSS/JS lives inline in one `<script>` block, and calculations read straight from `document.getElementById(...)`. This plan keeps that pattern (no new files, no new dependencies). It introduces one new DOM-reading accessor, `getDaySchedule(i)`, as the single place that decides "common schedule or this weekday's own schedule" — every other function that needs a day's hours goes through it (or through `getDailyMinFor(i)` / `getWeeklyMinutes()` / `getAvgDailyMin()`, which are built on it). No other function is allowed to read `startTime`/`endTime`/`breakTime`/`dayStart{i}`/`dayEnd{i}`/`dayBreak{i}` directly once this plan is done.

**Tech Stack:** Vanilla HTML/CSS/JS (no framework, no npm). Verification uses the Playwright browser tools (`mcp__plugin_playwright_playwright__browser_navigate`, `browser_snapshot`, `browser_fill_form`, `browser_type`, `browser_click`, `browser_evaluate`) already available in this environment, driving `index.html` directly via a `file://` URL — there is no dev server.

**Spec:** `docs/superpowers/specs/2026-08-23-weekday-schedule-design.md`

## Global Constraints

- Weekday index convention: 0=月, 1=火, 2=水, 3=木, 4=金, 5=土, 6=日 (matches existing `wd0`..`wd6` checkboxes in ④休日, and `dnames = ['月','火','水','木','金','土','日']` already used in `calcS1`).
- New checkbox id: `useWeekdaySchedule`. New per-day input ids: `dayStart{i}`, `dayEnd{i}`, `dayBreak{i}` for i=0..6.
- Every calculation function must obtain a day's hours only via `getDaySchedule(i)` / `getDailyMinFor(i)` / `getWeeklyMinutes()` / `getAvgDailyMin()` — never by reading `startTime`/`endTime`/`breakTime` or `dayStart{i}` etc. directly.
- 1-3 and ⑧'s annual-hours calculation basis: average daily hours = (sum of each work day's minutes) ÷ (number of work days), then × annual work days (365 − ④-2's annual holiday count). This replaces the old single `dm` (daily minutes) basis.
- `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html` is the URL to open in Playwright for all verification steps below.
- Excel version (`generate_check_sheet.py`) is out of scope for this plan.

---

### Task 1: Data-layer helpers + weekday schedule UI (additive only)

**Files:**
- Modify: `index.html` (CSS block ~line 65, HTML ①panel ~line 232, JS ~line 828, JS ~line 1485, JS init ~line 1780)

**Interfaces:**
- Produces: `getWorkDayIndices()` → `number[]` (indices 0-6 of days NOT marked holiday in ④). `getDaySchedule(i)` → `{start:'HH:MM', end:'HH:MM', break:number}`. `getDailyMinFor(i)` → `number` (minutes). `getWeeklyMinutes()` → `number`. `getAvgDailyMin()` → `number`. `toggleWeekdaySchedule()` — onchange handler for the new checkbox. `applyWeekdayScheduleVisibility()` — shows/hides the day table based on the checkbox, no side effects on input values.
- Consumes: existing `document.getElementById('wd'+i)` checkboxes (④), existing `timeToMin()`.

This task is purely additive: the old `getDailyMin()` function and all three of its existing call sites (`calcS1`, `calcS3`, `calcS8`) are left untouched. Nothing in the page's behavior changes yet — you're only adding new, unused-so-far functions and UI. This makes the task safe to verify as "zero visible change" plus "the new checkbox/table behaves correctly in isolation."

- [ ] **Step 1: Add CSS for the weekday schedule table**

In `index.html`, right after the `.weekday-label.is-holiday { ... }` rule (currently line 65), insert:

```css
    .weekday-schedule-table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 13px; }
    .weekday-schedule-table th, .weekday-schedule-table td { border: 1px solid #cbd5e0; padding: 6px 8px; text-align: left; }
    .weekday-schedule-table th { background: #eef2f7; font-weight: bold; }
    .weekday-schedule-table input[type="time"], .weekday-schedule-table input[type="number"] { width: 100%; box-sizing: border-box; padding: 4px 6px; border: 1px solid #cbd5e0; border-radius: 4px; }
    .weekday-schedule-table tr.is-holiday-row { background: #fef3c7; }
    .weekday-schedule-table tr.is-holiday-row input { background: #f3f4f6; color: #9ca3af; }
```

- [ ] **Step 2: Add the checkbox + day table markup to the ①panel**

In `index.html`, find (currently lines 232-235):

```html
    <div class="info-box">
      週の就業日数はタブ「④ 休日」で設定した休日から自動計算されます。変形労働時間制を採用する場合は「⑤ 変形労働時間制」タブも設定してください。
    </div>
```

Replace with:

```html
    <div class="info-box">
      週の就業日数はタブ「④ 休日」で設定した休日から自動計算されます。変形労働時間制を採用する場合は「⑤ 変形労働時間制」タブも設定してください。
    </div>
    <div class="form-row" style="margin-top:10px">
      <label style="display:flex;align-items:center;gap:6px;font-size:13px;font-weight:bold;color:#4a5568;cursor:pointer">
        <input type="checkbox" id="useWeekdaySchedule" onchange="toggleWeekdaySchedule(); recalcAll(); saveAll()">
        曜日ごとに所定労働時間を設定する
      </label>
    </div>
    <div id="weekdayScheduleTable" style="display:none; margin:6px 0 16px">
      <table class="weekday-schedule-table">
        <thead>
          <tr><th style="width:60px">曜日</th><th>始業</th><th>終業</th><th>休憩（分）</th></tr>
        </thead>
        <tbody>
          <tr id="dayRow0"><td>月</td><td><input type="time" id="dayStart0" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd0" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak0" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
          <tr id="dayRow1"><td>火</td><td><input type="time" id="dayStart1" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd1" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak1" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
          <tr id="dayRow2"><td>水</td><td><input type="time" id="dayStart2" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd2" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak2" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
          <tr id="dayRow3"><td>木</td><td><input type="time" id="dayStart3" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd3" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak3" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
          <tr id="dayRow4"><td>金</td><td><input type="time" id="dayStart4" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd4" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak4" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
          <tr id="dayRow5"><td>土</td><td><input type="time" id="dayStart5" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd5" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak5" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
          <tr id="dayRow6"><td>日</td><td><input type="time" id="dayStart6" oninput="recalcAll(); saveAll()"></td><td><input type="time" id="dayEnd6" oninput="recalcAll(); saveAll()"></td><td><input type="number" id="dayBreak6" min="0" max="480" oninput="recalcAll(); saveAll()"></td></tr>
        </tbody>
      </table>
    </div>
```

- [ ] **Step 3: Add the data-layer helper functions**

In `index.html`, find (currently lines 819-828):

```js
// ────────────────────────────────────────────────
// Core calculations
// ────────────────────────────────────────────────
function getDailyMin() {
  const start = timeToMin(document.getElementById('startTime').value);
  let end = timeToMin(document.getElementById('endTime').value);
  const brk = parseInt(document.getElementById('breakTime').value) || 0;
  if (end < start) end += 24 * 60; // overnight
  return Math.max(0, end - start - brk);
}
```

Replace with (keeps `getDailyMin()` byte-for-byte, only adds new functions after it):

```js
// ────────────────────────────────────────────────
// Core calculations
// ────────────────────────────────────────────────
function getDailyMin() {
  const start = timeToMin(document.getElementById('startTime').value);
  let end = timeToMin(document.getElementById('endTime').value);
  const brk = parseInt(document.getElementById('breakTime').value) || 0;
  if (end < start) end += 24 * 60; // overnight
  return Math.max(0, end - start - brk);
}
// NOTE: getDailyMin() above is the pre-weekday-schedule accessor. It is
// still used by calcS1/calcS3/calcS8 until Tasks 2-4 migrate them onto
// getDailyMinFor(i) below; it will be deleted once nothing calls it.

function getWorkDayIndices() {
  const r = [];
  for (let i = 0; i < 7; i++) if (!document.getElementById('wd' + i).checked) r.push(i);
  return r;
}
function getDaySchedule(i) {
  if (document.getElementById('useWeekdaySchedule').checked) {
    return {
      start: document.getElementById('dayStart' + i).value,
      end:   document.getElementById('dayEnd' + i).value,
      break: parseInt(document.getElementById('dayBreak' + i).value) || 0
    };
  }
  return {
    start: document.getElementById('startTime').value,
    end:   document.getElementById('endTime').value,
    break: parseInt(document.getElementById('breakTime').value) || 0
  };
}
function getDailyMinFor(i) {
  const sched = getDaySchedule(i);
  const start = timeToMin(sched.start);
  let end = timeToMin(sched.end);
  if (end < start) end += 24 * 60; // overnight
  return Math.max(0, end - start - sched.break);
}
function getWeeklyMinutes() {
  return getWorkDayIndices().reduce((sum, i) => sum + getDailyMinFor(i), 0);
}
function getAvgDailyMin() {
  const days = getWorkDayIndices();
  return days.length > 0 ? getWeeklyMinutes() / days.length : 0;
}
```

- [ ] **Step 4: Add the toggle/visibility handlers and wire them into the weekday-holiday styling**

In `index.html`, find (currently lines 1485-1488):

```js
function updateWeekdayStyle() {
  for (let i=0;i<7;i++)
    document.getElementById('wdLabel'+i).classList.toggle('is-holiday', document.getElementById('wd'+i).checked);
}
```

Replace with:

```js
function updateWeekdayStyle() {
  for (let i=0;i<7;i++) {
    const isHoliday = document.getElementById('wd'+i).checked;
    document.getElementById('wdLabel'+i).classList.toggle('is-holiday', isHoliday);
    const row = document.getElementById('dayRow'+i);
    if (row) {
      row.classList.toggle('is-holiday-row', isHoliday);
      document.getElementById('dayStart'+i).disabled = isHoliday;
      document.getElementById('dayEnd'+i).disabled = isHoliday;
      document.getElementById('dayBreak'+i).disabled = isHoliday;
    }
  }
}
function applyWeekdayScheduleVisibility() {
  document.getElementById('weekdayScheduleTable').style.display =
    document.getElementById('useWeekdaySchedule').checked ? 'block' : 'none';
}
function toggleWeekdaySchedule() {
  if (document.getElementById('useWeekdaySchedule').checked) {
    const common = {
      start: document.getElementById('startTime').value,
      end:   document.getElementById('endTime').value,
      brk:   document.getElementById('breakTime').value
    };
    for (let i = 0; i < 7; i++) {
      const s = document.getElementById('dayStart' + i);
      if (!s.value) {
        s.value = common.start;
        document.getElementById('dayEnd' + i).value = common.end;
        document.getElementById('dayBreak' + i).value = common.brk;
      }
    }
  }
  applyWeekdayScheduleVisibility();
}
```

- [ ] **Step 5: Sync visibility on load/restore**

In `index.html`, find in `loadAll()` (currently line 1762):

```js
  showWageFields();
}
```
(this is the one inside `function loadAll() {`, immediately before `resetAll()`)

Replace with:

```js
  showWageFields();
  applyWeekdayScheduleVisibility();
}
```

In `index.html`, find in `restoreCheck()` (currently line 1675):

```js
  showFlexDetails(); showPartTimeFields(); showWageFields(); updateWeekdayStyle();
```

Replace with:

```js
  showFlexDetails(); showPartTimeFields(); showWageFields(); updateWeekdayStyle(); applyWeekdayScheduleVisibility();
```

In `index.html`, find in the `DOMContentLoaded` handler (currently lines 1777-1780):

```js
  showFlexDetails();
  showPartTimeFields();
  showWageFields();
  updateWeekdayStyle();
```

Replace with:

```js
  showFlexDetails();
  showPartTimeFields();
  showWageFields();
  updateWeekdayStyle();
  applyWeekdayScheduleVisibility();
```

- [ ] **Step 6: Verify — regression (existing behavior unchanged)**

Using the Playwright browser tools:
1. `browser_navigate` to `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html`.
2. `browser_snapshot` and confirm judgment 1-1 shows `○` and its detail text still reads exactly as before (始業 09:00 終業 18:00 休憩60分 → 8時間).
3. Confirm the new checkbox "曜日ごとに所定労働時間を設定する" is present, unchecked, and the day table is not visible (no 月/火/… rows shown).

Expected: identical judgments to before this task (1-1 ○, 1-2 ✗ as in the original screenshot with 09:00-17:00/60分/土曜出勤), because `calcS1`/`calcS3`/`calcS8` still call the untouched `getDailyMin()`.

- [ ] **Step 7: Verify — new UI in isolation**

1. `browser_click` the "曜日ごとに所定労働時間を設定する" checkbox.
2. `browser_snapshot` — the day table should now be visible with all 7 rows, and 月〜金・土 (whichever are work days per ④'s default: 土・日 are checked as holiday by default per current index.html, so 土 and 日 rows should render disabled/greyed) pre-filled with 09:00 / 17:00 / 60 (the common values), non-holiday rows enabled, holiday rows (土, 日 by default) disabled and styled with the `is-holiday-row` background.
3. `browser_type` a different value into `dayStart5` (even though disabled by default this row won't accept input — instead uncheck `wd5` in ④休日 tab first via `browser_click`, return to ①, then edit `dayStart5`/`dayEnd5`/`dayBreak5`).
4. `browser_click` the checkbox off, then on again. Confirm the value you just typed is still there (not reset to the common value) — this checks the "only prefill empty fields" rule.

Expected: table shows/hides correctly, prefills only once, holiday rows are disabled and visually distinct, and manual edits survive toggling.

- [ ] **Step 8: Commit**

```bash
git add index.html
git commit -m "$(cat <<'EOF'
①労働時間: 曜日別スケジュールの入力UIとデータ層を追加

始業・終業・休憩を曜日ごとに設定できるチェックボックス＋テーブルと、
getDaySchedule/getDailyMinFor/getWeeklyMinutes/getAvgDailyMinを追加。
既存のgetDailyMin()とcalcS1/calcS3/calcS8はまだ変更しておらず、
既存の判定結果に影響はない（後続タスクで順次移行）。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 2: Wire ①1-1 / 1-2 / 1-3 onto the weekday schedule

**Files:**
- Modify: `index.html` (`calcS1`, currently lines 877-926)

**Interfaces:**
- Consumes: `getWorkDayIndices()`, `getDaySchedule(i)`, `getDailyMinFor(i)` from Task 1.
- Produces: same return shape as before, `{ '1-1': 'ok'|'ng', '1-2': 'ok'|'ng' }`.

- [ ] **Step 1: Replace `calcS1`**

In `index.html`, find the full function (currently lines 877-926):

```js
function calcS1() {
  const dm = getDailyMin();
  const wdays = getWeeklyWorkDays();
  const wMin = dm * wdays;
  const wLim = weeklyLimit();
  const brk = parseInt(document.getElementById('breakTime').value) || 0;
  const st = document.getElementById('startTime').value;
  const et = document.getElementById('endTime').value;
  const year = new Date().getFullYear();
  const annOff1 = getAnnualOffDays(year);   // ④-2の年間休日日数と同じ値
  const awd = 365 - annOff1;
  const pubHol = countPubHolOnWorkDays(year);
  const addHol = getAdditionalHolidays();
  const dnames = ['月','火','水','木','金','土','日'];
  const offIdx = [];
  for (let i = 0; i < 7; i++) if (document.getElementById('wd'+i).checked) offIdx.push(i);
  const workStr = dnames.filter((_,i)=>!offIdx.includes(i)).join('・') || 'なし';
  const offStr  = offIdx.length ? offIdx.map(i=>dnames[i]).join('・') : 'なし';
  const holLabel = document.getElementById('holidayType').value === 'off' ? '休日扱い' : '出勤日';

  const j11 = dm/60 <= 8 ? 'ok' : 'ng';
  setJ('j1-1', j11);
  setP('proc1-1',
    '始業: ' + st + '  終業: ' + et + '  休憩: ' + brk + '分\n' +
    '計算式: (' + et + ' − ' + st + ') − ' + brk + '分 = ' + dm + '分 = ' + minToStr(dm) + '\n' +
    '法定基準: 8時間以内（労基法 第32条2項）\n' +
    '判  定: ' + (j11==='ok' ? '○ 法定基準内' : '✗ 8時間超過 → 時間外労働となるため36協定が必要'));

  const j12 = wMin/60 <= wLim ? 'ok' : 'ng';
  setJ('j1-2', j12);
  setP('proc1-2',
    '就業日: ' + workStr + '（週' + wdays + '日）  休日: ' + offStr + '（週' + (7-wdays) + '日）\n' +
    '計算式: ' + minToH(dm) + '時間/日 × 週' + wdays + '日 = ' + minToH(wMin) + '時間/週\n' +
    '法定基準: ' + wLim + '時間以内' + (isSpecial() ? '（特例措置適用中: 44時間）' : '') + '（労基法 第32条1項' + (isSpecial() ? '、第40条' : '') + '）\n' +
    '判  定: ' + (j12==='ok' ? '○ 法定基準内' : '✗ ' + wLim + '時間超過'));

  const offDaysBreakdown =
    '週休: ' + getWeeklyOffDays() + '日 × 52週 = ' + (getWeeklyOffDays()*52) + '日' +
    (pubHol > 0 ? '\n祝日（就業日に当たる分）: ' + pubHol + '日  ← 祝日は' + holLabel : '') +
    (addHol > 0 ? '\n追加休日: ' + addHol + '日' : '');
  setJ('j1-3', 'info');
  setP('proc1-3',
    '④ 休日の設定から推計\n' +
    offDaysBreakdown + '\n' +
    '= 年間休日合計（④-2の値）: ' + annOff1 + '日\n' +
    '年間就業日数: 365 − ' + annOff1 + ' = ' + awd + '日\n' +
    '年間総労働時間（推定）: ' + awd + '日 × ' + minToH(dm) + '時間 = ' + (dm * awd / 60).toFixed(1) + '時間');

  return { '1-1': j11, '1-2': j12 };
}
```

Replace with:

```js
function calcS1() {
  const dnames = ['月','火','水','木','金','土','日'];
  const workDays = getWorkDayIndices();
  const wLim = weeklyLimit();
  const year = new Date().getFullYear();
  const annOff1 = getAnnualOffDays(year);   // ④-2の年間休日日数と同じ値
  const awd = 365 - annOff1;
  const pubHol = countPubHolOnWorkDays(year);
  const addHol = getAdditionalHolidays();
  const offIdx = [];
  for (let i = 0; i < 7; i++) if (document.getElementById('wd'+i).checked) offIdx.push(i);
  const workStr = dnames.filter((_,i)=>!offIdx.includes(i)).join('・') || 'なし';
  const offStr  = offIdx.length ? offIdx.map(i=>dnames[i]).join('・') : 'なし';
  const holLabel = document.getElementById('holidayType').value === 'off' ? '休日扱い' : '出勤日';

  const perDay = workDays.map(i => {
    const sched = getDaySchedule(i);
    const min = getDailyMinFor(i);
    return { i, name: dnames[i], sched, min };
  });
  const wMin = perDay.reduce((s,d) => s + d.min, 0);
  const avgMin = perDay.length ? wMin / perDay.length : 0;
  const overDays = perDay.filter(d => d.min/60 > 8);

  const j11 = overDays.length ? 'ng' : 'ok';
  setJ('j1-1', j11);
  const dayLines1 = perDay.map(d =>
    d.name + ': 始業 ' + d.sched.start + '  終業 ' + d.sched.end + '  休憩 ' + d.sched.break + '分 → ' + minToStr(d.min) +
    (d.min/60 > 8 ? '（8時間超）' : '')
  ).join('\n');
  setP('proc1-1',
    (dayLines1 || '（就業日がありません）') + '\n\n' +
    '法定基準: 8時間以内（労基法 第32条2項）\n' +
    '判  定: ' + (j11==='ok' ? '○ 法定基準内' : '✗ ' + overDays.map(d=>d.name).join('・') + '曜日が8時間超過 → 時間外労働となるため36協定が必要'));

  const j12 = wMin/60 <= wLim ? 'ok' : 'ng';
  setJ('j1-2', j12);
  const dayLines2 = perDay.map(d => minToH(d.min) + 'h(' + d.name + ')').join(' + ') || '0';
  setP('proc1-2',
    '就業日: ' + workStr + '（週' + workDays.length + '日）  休日: ' + offStr + '（週' + (7-workDays.length) + '日）\n' +
    '計算式: ' + dayLines2 + ' = ' + minToH(wMin) + '時間/週\n' +
    '法定基準: ' + wLim + '時間以内' + (isSpecial() ? '（特例措置適用中: 44時間）' : '') + '（労基法 第32条1項' + (isSpecial() ? '、第40条' : '') + '）\n' +
    '判  定: ' + (j12==='ok' ? '○ 法定基準内' : '✗ ' + wLim + '時間超過'));

  const offDaysBreakdown =
    '週休: ' + getWeeklyOffDays() + '日 × 52週 = ' + (getWeeklyOffDays()*52) + '日' +
    (pubHol > 0 ? '\n祝日（就業日に当たる分）: ' + pubHol + '日  ← 祝日は' + holLabel : '') +
    (addHol > 0 ? '\n追加休日: ' + addHol + '日' : '');
  setJ('j1-3', 'info');
  setP('proc1-3',
    '④ 休日の設定から推計\n' +
    offDaysBreakdown + '\n' +
    '= 年間休日合計（④-2の値）: ' + annOff1 + '日\n' +
    '年間就業日数: 365 − ' + annOff1 + ' = ' + awd + '日\n' +
    '平均1日所定労働時間: ' + minToH(wMin) + '時間/週 ÷ 週' + workDays.length + '日 = ' + minToH(avgMin) + '時間/日\n' +
    '年間総労働時間（推定）: ' + awd + '日 × ' + minToH(avgMin) + '時間 = ' + (avgMin * awd / 60).toFixed(1) + '時間');

  return { '1-1': j11, '1-2': j12 };
}
```

- [ ] **Step 2: Verify — regression (common mode unchanged)**

1. `browser_navigate` to `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html` (leave 曜日ごとに設定する unchecked).
2. `browser_snapshot`. Confirm 1-1, 1-2, 1-3 show the same judgments and same numeric values as before Task 2 (09:00-17:00/60分, 土曜出勤 → 1-1 ○ 7時間, 1-2 ✗ 42時間超過, matching the screenshot in the original request).

- [ ] **Step 3: Verify — weekday schedule affects 1-1/1-2/1-3**

1. `browser_click` "曜日ごとに設定する". In ④休日, ensure only 日 is checked as holiday (uncheck 土 if needed) so 月〜土 are all work days.
2. Set 土 (`dayStart5`/`dayEnd5`/`dayBreak5`) to 09:00 / 12:00 / 0 via `browser_fill_form` or `browser_type`, leave 月〜金 at 09:00/17:00/60.
3. `browser_snapshot`. Expected: 1-1 stays ○ (no day exceeds 8h — 月〜金は7h、土は3h), 1-2's detail text lists `7.00h(月) + 7.00h(火) + ... + 3.00h(土) = 38.00時間/週` and judges ○ (38 ≤ 40), 1-3 shows 平均1日所定労働時間 = 38.00 ÷ 6 = 6.33時間/日.
4. Change 月 (`dayEnd0`) to 19:00 (10h day, minus 60min break = 9h). `browser_snapshot`. Expected: 1-1 turns ✗ and its detail names 月 as the over-8h day.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "$(cat <<'EOF'
①労働時間: 1-1/1-2/1-3を曜日別スケジュール対応に変更

各就業日の所定労働時間をgetDaySchedule/getDailyMinForから取得し、
1-1は1日でも8時間超があれば✗、1-2は曜日ごとの合計で週判定、
1-3の年間総労働時間は週合計÷週就業日数の平均1日時間を基礎にした。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 3: Wire ③3-1 onto the weekday schedule

**Files:**
- Modify: `index.html` (`calcS3`, currently lines 976-997)

**Interfaces:**
- Consumes: `getWorkDayIndices()`, `getDaySchedule(i)`, `getDailyMinFor(i)`.
- Produces: same return shape as before, `{ '3-1': 'ok'|'ng', '3-2': 'ok' }`.

- [ ] **Step 1: Replace `calcS3`**

In `index.html`, find the full function (currently lines 976-997):

```js
function calcS3() {
  const dm  = getDailyMin();
  const brk = parseInt(document.getElementById('breakTime').value) || 0;
  const exc = document.getElementById('breakException').value;
  let req = 0;
  if (dm > 480) req = 60;
  else if (dm > 360) req = 45;

  const j31 = brk >= req ? 'ok' : 'ng';
  setJ('j3-1', j31);
  setP('proc3-1', req === 0
    ? '1日の労働時間: ' + minToStr(dm) + '（6時間以下）\n→ 法律上の休憩付与義務なし\n判定: ○'
    : '1日の労働時間: ' + minToStr(dm) + '\n→ ' + (dm>480 ? '8時間超 → 60分以上' : '6時間超8時間以下 → 45分以上') + 'の休憩が必要（労基法 第34条1項）\n設定休憩: ' + brk + '分  必要休憩: ' + req + '分\n判  定: ' + (j31==='ok' ? '○ 基準内' : '✗ 休憩時間不足（' + req + '分以上必要）'));

  setJ('j3-2', 'ok');
  setP('proc3-2',
    '一斉付与の適用除外（労使協定）: ' + (exc==='yes' ? 'あり' : 'なし') + '\n' +
    (exc==='yes'
      ? '労使協定あり → 一斉付与の適用除外として問題なし\n判定: ○'
      : '一斉付与が原則（労基法 第34条2項）\n例外業種（運輸・商業・飲食等）は適用除外\n判定: ○（対象業種を個別確認推奨）'));
  return { '3-1': j31, '3-2': 'ok' };
}
```

Replace with:

```js
function calcS3() {
  const dnames = ['月','火','水','木','金','土','日'];
  const exc = document.getElementById('breakException').value;
  const workDays = getWorkDayIndices();

  const perDay = workDays.map(i => {
    const sched = getDaySchedule(i);
    const min = getDailyMinFor(i);
    let req = 0;
    if (min > 480) req = 60;
    else if (min > 360) req = 45;
    return { i, name: dnames[i], min, brk: sched.break, req, ok: sched.break >= req };
  });
  const shortDays = perDay.filter(d => !d.ok);

  const j31 = shortDays.length ? 'ng' : 'ok';
  setJ('j3-1', j31);
  const dayLines = perDay.map(d =>
    d.name + ': 労働 ' + minToStr(d.min) + ' → 必要休憩 ' + (d.req === 0 ? '不要（6時間以下）' : d.req + '分以上') +
    '  設定休憩 ' + d.brk + '分' + (d.ok ? '' : '（不足）')
  ).join('\n');
  setP('proc3-1',
    (dayLines || '（就業日がありません）') + '\n\n' +
    '法定基準: 6時間超→45分以上、8時間超→60分以上（労基法 第34条1項）\n' +
    '判  定: ' + (j31==='ok' ? '○ 基準内' : '✗ ' + shortDays.map(d=>d.name).join('・') + '曜日の休憩時間不足'));

  setJ('j3-2', 'ok');
  setP('proc3-2',
    '一斉付与の適用除外（労使協定）: ' + (exc==='yes' ? 'あり' : 'なし') + '\n' +
    (exc==='yes'
      ? '労使協定あり → 一斉付与の適用除外として問題なし\n判定: ○'
      : '一斉付与が原則（労基法 第34条2項）\n例外業種（運輸・商業・飲食等）は適用除外\n判定: ○（対象業種を個別確認推奨）'));
  return { '3-1': j31, '3-2': 'ok' };
}
```

- [ ] **Step 2: Verify — regression (common mode unchanged)**

1. `browser_navigate` to `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html` (曜日ごとに設定する unchecked).
2. `browser_snapshot`. Confirm 3-1 shows the same judgment/text as before (09:00-17:00, 休憩60分, 7時間 → 休憩不要なし判定 ○, since 7h ≤ 6h is false... verify against the actual current output rather than assuming — read the snapshot and confirm it is unchanged from a pre-Task-3 run of the same inputs).

- [ ] **Step 3: Verify — per-day break requirement**

1. `browser_click` "曜日ごとに設定する". Ensure 月〜土 are work days (④で日のみ休日).
2. Set 月 to 09:00/18:00/60分 (8h work → needs 60min break; 60 is OK). Set 火 to 09:00/18:30/30分 (8.5h work → needs 60min break; 30 is short).
3. `browser_snapshot`. Expected: 3-1 judges ✗ and names 火 as the day with insufficient break; the detail lines show 月 as satisfied and 火 as `（不足）`.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "$(cat <<'EOF'
①休憩: 3-1を曜日別スケジュール対応に変更

各就業日の労働時間から必要休憩を求め、その日の設定休憩と比較。
1日でも不足があれば✗とし、詳細欄に曜日別の内訳を表示する。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 4: Wire ⑧最低賃金 (日給制) onto the weekday schedule, delete `getDailyMin()`

**Files:**
- Modify: `index.html` (`calcS8`, currently lines 1276-1441)

**Interfaces:**
- Consumes: `getAvgDailyMin()` from Task 1.
- Produces: same return shapes as before for `'8-1'`/`'8-2'`/`'8-3'`.

This is the last remaining caller of `getDailyMin()`; once it's migrated, the old function is deleted.

- [ ] **Step 1: Replace the `dm`/`dailyH` declaration at the top of `calcS8`**

In `index.html`, find (currently lines 1276-1280):

```js
function calcS8() {
  const minWage = parseFloat(document.getElementById('minWage').value) || 0;
  const wageType = document.getElementById('wageType').value;
  const dm = getDailyMin();
  const dailyH = dm / 60;
```

Replace with:

```js
function calcS8() {
  const minWage = parseFloat(document.getElementById('minWage').value) || 0;
  const wageType = document.getElementById('wageType').value;
  const dailyH = getAvgDailyMin() / 60;
```

(`dailyH` keeps its name — every downstream reference in this function stays correct — but now holds the average across work days instead of the single common day's hours.)

- [ ] **Step 2: Update the daily-wage branch's label text**

In `index.html`, find (currently lines 1421-1426):

```js
    setP('proc8-1',
      '適用最低賃金: ' + pref + ' ' + minWage.toLocaleString() + '円/時間\n' +
      '日給: ' + daily.toLocaleString() + '円\n\n' +
      '【時間換算額の計算】\n' +
      '  1日の所定労働時間: ' + dailyH.toFixed(2) + '時間\n' +
      '  ' + daily.toLocaleString() + '円 ÷ ' + dailyH.toFixed(2) + '時間 = ' + hourlyRate.toFixed(2) + '円/時間\n\n' +
```

Replace with:

```js
    setP('proc8-1',
      '適用最低賃金: ' + pref + ' ' + minWage.toLocaleString() + '円/時間\n' +
      '日給: ' + daily.toLocaleString() + '円\n\n' +
      '【時間換算額の計算】\n' +
      '  平均1日所定労働時間: ' + dailyH.toFixed(2) + '時間\n' +
      '  ' + daily.toLocaleString() + '円 ÷ ' + dailyH.toFixed(2) + '時間 = ' + hourlyRate.toFixed(2) + '円/時間\n\n' +
```

- [ ] **Step 3: Delete the now-unused `getDailyMin()`**

In `index.html`, find (currently lines 821-828, immediately followed by the NOTE comment added in Task 1):

```js
function getDailyMin() {
  const start = timeToMin(document.getElementById('startTime').value);
  let end = timeToMin(document.getElementById('endTime').value);
  const brk = parseInt(document.getElementById('breakTime').value) || 0;
  if (end < start) end += 24 * 60; // overnight
  return Math.max(0, end - start - brk);
}
// NOTE: getDailyMin() above is the pre-weekday-schedule accessor. It is
// still used by calcS1/calcS3/calcS8 until Tasks 2-4 migrate them onto
// getDailyMinFor(i) below; it will be deleted once nothing calls it.

function getWorkDayIndices() {
```

Replace with:

```js
function getWorkDayIndices() {
```

(This removes the dead function and its NOTE comment. Confirm with `grep -n "getDailyMin()" index.html` that no call sites remain before proceeding — only the definition of `getDailyMinFor` should match, not `getDailyMin()` with no suffix.)

- [ ] **Step 4: Verify — regression (common mode unchanged)**

1. `browser_navigate` to `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html`.
2. Go to ⑧最低賃金, select 賃金の種類=日給制, enter a daily salary (e.g. 8000), select 都道府県=大阪.
3. `browser_snapshot`. Confirm 8-1's numeric result matches what the same inputs produced before this task (平均1日所定労働時間 = 7.00時間 since 月〜金/土 aren't yet weekday-customized, and default ④ holidays are 土・日 → work days 月〜金 all 7h).

- [ ] **Step 5: Verify — weekday schedule affects ⑧'s average**

1. `browser_click` "曜日ごとに設定する" in ①. In ④ ensure 月〜土 are work days (日 only holiday). Set 土 to 09:00/12:00/0 (3h), leave 月〜金 at 7h each.
2. Go to ⑧, wageType=日給制, daily salary e.g. 8000, 都道府県=大阪.
3. `browser_snapshot`. Expected: 平均1日所定労働時間 = (7×5+3)/6 = 6.33時間, and the displayed hourly rate = 8000/6.33.
4. Switch wageType=月給制 with a monthly salary and fixed OT hours set (⑧-3 fields), confirm 8-1/8-2/8-3 still compute without errors and their "年間総労働時間" line uses the same 6.33h/day average (cross-check against 1-3's value from Task 2's verification).

- [ ] **Step 6: Run a full-file syntax sanity check**

```bash
node --check index.html 2>&1 | head -5 || true
```

This will report a syntax error location if the `<script>` content is malformed; since `index.html` is not pure JS, expect it to fail on the HTML around it — instead extract just the script block and check that:

```bash
node -e "
const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');
const m = html.match(/<script>([\s\S]*?)<\/script>\s*<\/body>/);
if (!m) { console.error('script block not found'); process.exit(1); }
new Function(m[1]);
console.log('OK: script block parses');
"
```

Expected output: `OK: script block parses` (a `SyntaxError` here means a stray bracket/paren was introduced by one of the edits above — fix it before continuing).

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "$(cat <<'EOF'
⑧最低賃金: 日給制の時間換算を平均1日所定労働時間ベースに変更

getAvgDailyMin()を使うことで、曜日別スケジュール設定時も1-3と同じ
基礎値（週合計÷週就業日数）で時間単価を算出する。calcS1/calcS3/calcS8
すべての移行が完了したため、未使用になったgetDailyMin()を削除。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 5: Persist weekday schedule (localStorage + Firestore history) with migration

**Files:**
- Modify: `index.html` (`getCurrentInputs`, currently line 1546; `saveAll`, currently line 1720; `restoreCheck`, currently line 1657; `loadAll`, currently line 1742)

**Interfaces:**
- Produces: `useWeekdaySchedule: boolean` and `daySchedule: Array<{start,end,break}>` (length 7) added to the saved-data shape used by both localStorage and the Firestore-backed cloud history (both go through `getCurrentInputs()`/`restoreCheck()`).

- [ ] **Step 1: Add fields to `getCurrentInputs()`**

In `index.html`, find (currently lines 1546-1564, the closing part of the object):

```js
function getCurrentInputs() {
  return {
    businessName:gv('businessName'), checkDate:gv('checkDate'), specialMeasure:gc('specialMeasure'),
    startTime:gv('startTime'), endTime:gv('endTime'), breakTime:gv('breakTime'),
```

Replace with:

```js
function getCurrentInputs() {
  return {
    businessName:gv('businessName'), checkDate:gv('checkDate'), specialMeasure:gc('specialMeasure'),
    startTime:gv('startTime'), endTime:gv('endTime'), breakTime:gv('breakTime'),
    useWeekdaySchedule:gc('useWeekdaySchedule'),
    daySchedule: Array.from({length:7}, (_,i) => ({
      start: gv('dayStart'+i), end: gv('dayEnd'+i), break: gv('dayBreak'+i)
    })),
```

- [ ] **Step 2: Add fields to `saveAll()`**

In `index.html`, find (currently lines 1720-1723):

```js
function saveAll() {
  const d = {
    businessName:gv('businessName'), checkDate:gv('checkDate'), specialMeasure:gc('specialMeasure'),
    startTime:gv('startTime'), endTime:gv('endTime'), breakTime:gv('breakTime'),
```

Replace with:

```js
function saveAll() {
  const d = {
    businessName:gv('businessName'), checkDate:gv('checkDate'), specialMeasure:gc('specialMeasure'),
    startTime:gv('startTime'), endTime:gv('endTime'), breakTime:gv('breakTime'),
    useWeekdaySchedule:gc('useWeekdaySchedule'),
    daySchedule: Array.from({length:7}, (_,i) => ({
      start: gv('dayStart'+i), end: gv('dayEnd'+i), break: gv('dayBreak'+i)
    })),
```

- [ ] **Step 3: Restore fields in `restoreCheck()`**

In `index.html`, find (currently lines 1657-1661):

```js
function restoreCheck(inputs) {
  if (!inputs) return;
  sv('businessName',inputs.businessName); sv('checkDate',inputs.checkDate); sc('specialMeasure',inputs.specialMeasure);
  sv('startTime',inputs.startTime); sv('endTime',inputs.endTime); sv('breakTime',inputs.breakTime);
```

Replace with:

```js
function restoreCheck(inputs) {
  if (!inputs) return;
  sv('businessName',inputs.businessName); sv('checkDate',inputs.checkDate); sc('specialMeasure',inputs.specialMeasure);
  sv('startTime',inputs.startTime); sv('endTime',inputs.endTime); sv('breakTime',inputs.breakTime);
  sc('useWeekdaySchedule', inputs.useWeekdaySchedule);
  if (Array.isArray(inputs.daySchedule)) inputs.daySchedule.forEach((d,i) => {
    if (!d) return;
    sv('dayStart'+i, d.start); sv('dayEnd'+i, d.end); sv('dayBreak'+i, d.break);
  });
```

- [ ] **Step 4: Restore fields in `loadAll()`**

In `index.html`, find (currently lines 1742-1747):

```js
function loadAll() {
  let d;
  try { d = JSON.parse(localStorage.getItem(LS_KEY)); } catch(e) { return; }
  if (!d) return;
  sv('businessName',d.businessName); sv('checkDate',d.checkDate); sc('specialMeasure',d.specialMeasure);
  sv('startTime',d.startTime); sv('endTime',d.endTime); sv('breakTime',d.breakTime);
```

Replace with:

```js
function loadAll() {
  let d;
  try { d = JSON.parse(localStorage.getItem(LS_KEY)); } catch(e) { return; }
  if (!d) return;
  sv('businessName',d.businessName); sv('checkDate',d.checkDate); sc('specialMeasure',d.specialMeasure);
  sv('startTime',d.startTime); sv('endTime',d.endTime); sv('breakTime',d.breakTime);
  sc('useWeekdaySchedule', d.useWeekdaySchedule);
  if (Array.isArray(d.daySchedule)) d.daySchedule.forEach((day,i) => {
    if (!day) return;
    sv('dayStart'+i, day.start); sv('dayEnd'+i, day.end); sv('dayBreak'+i, day.break);
  });
```

- [ ] **Step 5: Verify — save/reload round-trip (localStorage)**

1. `browser_navigate` to `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html`.
2. `browser_click` "曜日ごとに設定する". In ④, ensure 月〜土 are work days. Set 土 to 09:00/12:00/0.
3. `browser_navigate` to the same `file://` URL again (reload). `browser_snapshot`.
4. Expected: "曜日ごとに設定する" is still checked, the day table is visible, and 土's row shows 09:00/12:00/0 — not the common 09:00/17:00/60.

- [ ] **Step 6: Verify — old saved data still loads (migration)**

1. Using `browser_evaluate`, run (against the same page, before step 2's changes are made, or in a fresh navigation): set `localStorage` directly to an old-shape payload (no `useWeekdaySchedule`/`daySchedule` keys) matching the pre-Task-5 `saveAll()` output, e.g.:

```js
localStorage.setItem('hwLaborChecker_v1', JSON.stringify({
  businessName: '', checkDate: '2026-08-23', specialMeasure: false,
  startTime: '09:00', endTime: '17:00', breakTime: '60',
  weekdays: [false,false,false,false,false,true,true]
}));
```

2. `browser_navigate` (reload) to the same URL. `browser_snapshot`.
3. Expected: page loads without a JS error, "曜日ごとに設定する" is unchecked, common startTime/endTime/breakTime show 09:00/17:00/60, and 1-1/1-2 compute normally (same as before this whole plan started).

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "$(cat <<'EOF'
①労働時間: 曜日別スケジュールをlocalStorage/クラウド保存に対応

useWeekdaySchedule・daySchedule[7]をgetCurrentInputs/saveAll/
restoreCheck/loadAllに追加。旧形式データ（フィールドなし）は
共通モードのまま読み込めるため、移行処理は不要。

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```

---

### Task 6: Full regression + edge-case verification pass

**Files:** none (verification only; fix inline if something's off, then commit any fix)

This task walks every item in the spec's テスト観点 list end-to-end against the finished feature, using the Playwright browser tools against `file:///mnt/c/Users/y_uej/CC-RodoJokenCheck/index.html`.

- [ ] **Step 1: Common mode parity**

Leave "曜日ごとに設定する" unchecked. Set 始業09:00/終業17:00/休憩60分, ④で土・日を休日. `browser_snapshot` and confirm 1-1/1-2/1-3/3-1/⑧ all match the numbers this exact input produced before this plan (7h/day, 42h→ recompute for 月〜金のみ=5日: 35h/週 ≤ 40 → ○; adjust expectation to whatever ④'s current default holiday set is and confirm internally consistent, not against a stale memory of the number).

- [ ] **Step 2: All-days-same weekday mode ≡ common mode**

Check "曜日ごとに設定する", leave every day at the prefilled common values. `browser_snapshot`. Expected: every judgment and every number is identical to Step 1's result.

- [ ] **Step 3: Partial short day (土曜3時間)**

From Step 2's state, set 土 to 09:00/12:00/0分 (ensure 土 is a work day in ④). `browser_snapshot`. Expected (from Task 2/3's own verification, re-confirmed here in combination with everything else): 1-1 ○, 1-2's weekly total drops accordingly, 3-1 shows 土 needs no break and has none set (○ for that day), ⑧'s average daily hours reflects the shorter 土.

- [ ] **Step 4: A weekday exceeding 8h**

From Step 3's state, set 月 to 09:00/19:00/60分 (9h). `browser_snapshot`. Expected: 1-1 turns ✗ naming 月; other judgments unaffected structurally (1-2 recomputes with the larger 月).

- [ ] **Step 5: Holiday day is locked out of the table**

In ④, check 水 as a holiday. Go to ①. `browser_snapshot`. Expected: `dayRow2`（水）is greyed out (`is-holiday-row`), its 3 inputs are disabled, and 1-1/1-2/3-1's per-day breakdown no longer lists 水.

- [ ] **Step 6: Toggle off/on preserves edits**

Uncheck "曜日ごとに設定する" (table hides), re-check it (table shows). `browser_snapshot`. Expected: all previously-entered per-day values (月19:00, 土09:00-12:00, etc.) are still present — not reset to the common 09:00/17:00/60.

- [ ] **Step 7: Save → reload → history restore, end to end**

1. With Step 6's state, wait for `saveAll()` to fire (it fires on every `oninput`/`onchange` already exercised above) and reload the page. `browser_snapshot`. Expected: all state from Step 6 persists (localStorage path, already covered narrowly in Task 5 — this step confirms it holds under the fuller Step 1-6 scenario).
2. If a Firebase-authenticated cloud-history flow is reachable without live credentials in this environment, skip that part and note it as untested (Firestore requires a real Google sign-in this environment cannot perform) — the localStorage path is the one to fully verify here.

- [ ] **Step 8: Fix anything found, or confirm clean**

If any of Steps 1-7 surfaced a discrepancy, fix it in `index.html` now and re-run the specific step that failed until it passes. If everything passed, no code change is needed for this task.

- [ ] **Step 9: Commit (only if Step 8 made a fix)**

If Step 8 found nothing to fix, skip this step — there is nothing to commit.

Otherwise, commit with a message whose body states which of Steps 1-7 failed and what the actual code change was (e.g. "Step 4 showed 1-2's weekly total using stale `wMin` from before the per-day edit — `calcS1` now re-reads `getDaySchedule` on every call instead of caching it"). Use this shape:

```bash
git add index.html
git commit -m "$(cat <<'EOF'
①労働時間: 曜日別スケジュールの回帰確認で見つかった不具合を修正

<Step 8で実際に直した内容を1-2文で、失敗したステップ番号と症状・原因を明記して書く>

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_01HR16YRbqPstKNBmEN3ZNMs
EOF
)"
```
