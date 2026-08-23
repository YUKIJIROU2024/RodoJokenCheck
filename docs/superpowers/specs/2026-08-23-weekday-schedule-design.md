# ①労働時間：曜日別の所定労働時間 設計

- 日付: 2026-08-23
- 対象: `index.html`（Web版）。`generate_check_sheet.py`（Excel版）は対象外（後続タスク）。

## 背景・目的

現状、①労働時間の始業・終業・休憩は全曜日共通の1組しか設定できない。曜日によって所定労働時間を短く設定している事業所（例: 土曜だけ9:00-12:00・休憩なし）を検証できるようにする。

## 要件（ユーザー確認済み）

1. 曜日ごとに設定できるのは「始業・終業・休憩」の3項目（時間数だけでなく時刻ベース）。
2. 入力UIは「共通1組」と「曜日別7組」を切り替え式にする。デフォルトは共通（既存の見た目のまま）。切り替えチェックボックスをONにしたときだけ曜日別入力を表示する。
3. 年間総労働時間（①-3・⑧の月給/日給の時間換算の基礎）は「週合計労働時間 ÷ 週就業日数 ＝ 平均1日時間」を出し、年間就業日数（365－④-2の年間休日日数）に掛けて算出する。
4. Excel版（`generate_check_sheet.py`）は今回のスコープ外。

## データモデル

### 新規input要素

`wd0`〜`wd6`（0=月〜6=日、既存の休日チェックボックスと同じ並び）に対応させ、各曜日ごとに:

- `dayStart{i}`（time型、i=0..6）
- `dayEnd{i}`（time型）
- `dayBreak{i}`（number型、分）

を追加する。表示は「④ 休日」の曜日チェックボックス群と同様に、①パネル内に7列（または7行）のテーブルとして配置する。

加えて切り替え用チェックボックス `useWeekdaySchedule`（曜日ごとに所定労働時間を設定する）を追加する。

既存の `startTime` / `endTime` / `breakTime` は「共通モード」の入力欄として維持する（削除しない）。

### 単一アクセサ関数

計算コードは曜日別かどうかを意識しない。新設する関数:

```js
function getDaySchedule(i) {
  // i: 0=月..6=日
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
  const { start, end, break: brk } = getDaySchedule(i);
  const s = timeToMin(start);
  let e = timeToMin(end);
  if (e < s) e += 24 * 60; // overnight
  return Math.max(0, e - s - brk);
}
```

`getDailyMin()`（引数なし・現行関数）は削除し、呼び出し元をすべて `getDailyMinFor(i)` または後述の週集計関数に置き換える。

### 週集計ヘルパー

```js
function getWorkDayIndices() {
  // 就業日（④の休日チェックが外れている曜日）のインデックス配列（0..6）
  const r = [];
  for (let i = 0; i < 7; i++) if (!document.getElementById('wd' + i).checked) r.push(i);
  return r;
}
function getWeeklyMinutes() {
  return getWorkDayIndices().reduce((sum, i) => sum + getDailyMinFor(i), 0);
}
function getAvgDailyMin() {
  const days = getWorkDayIndices();
  return days.length > 0 ? getWeeklyMinutes() / days.length : 0;
}
```

`getWeeklyWorkDays()`（既存、就業日数を返すだけの関数）はそのまま残す。

## UI変更（①労働時間パネル）

1. 既存の始業/終業/休憩の3項目はそのまま残す（共通モードの入力として機能）。
2. その下に `useWeekdaySchedule` チェックボックス「曜日ごとに所定労働時間を設定する」を追加。
3. チェックボックスの直下に曜日別テーブル（初期は `display:none`、ON時に `display:block`）を追加。各行: 曜日名（④の休日設定と連動して「休日」の曜日は入力欄をグレーアウト・入力不可にし「－（④で休日）」と表示）、始業・終業・休憩の入力欄。
4. `useWeekdaySchedule` をONにした瞬間、`dayStart{i}` 等が空の場合のみ、共通の `startTime`/`endTime`/`breakTime` の値で初期化する（`onchange` ハンドラで一度だけ実施。2回目以降のON/OFF切り替えでは上書きしない＝入力済みの曜日別値を保持する）。

## 各判定ロジックの変更

### 1-1：1日の所定労働時間（労基法32条2項・8時間）

- 就業日それぞれについて `getDailyMinFor(i)` を計算し、1日でも8時間（480分）を超えていれば ✗、全日8時間以内なら ○。
- 詳細欄（`proc1-1`）に曜日ごとの内訳（始業/終業/休憩/労働時間）を一覧表示し、超過している曜日を明示する。

### 1-2：週の所定労働時間（労基法32条1項・週40/44時間）

- `getWeeklyMinutes()` を週上限（`weeklyLimit()`）と比較。ロジック自体は現行と同じで、入力元だけ曜日別集計に置き換える。
- 詳細欄に曜日ごとの労働時間と合計を表示。

### 1-3：年間総労働時間（参考）

- `getAvgDailyMin()` で平均1日時間を算出。
- 年間休日日数（④-2と同じ値）→ 年間就業日数 = 365－年間休日日数、は現行ロジックのまま。
- 年間総労働時間 = 年間就業日数 × 平均1日時間。
- 詳細欄に「週合計 ÷ 週就業日数 = 平均1日時間」の内訳を追記する。

### 3-1：休憩時間の長さ（労基法34条1項）

- 就業日それぞれについて、その日の `getDailyMinFor(i)` から必要休憩（6h超→45分、8h超→60分、6h以下→不要）を求め、その日の `getDaySchedule(i).break` と比較。
- 1日でも不足があれば ✗。全日クリアなら ○。
- 詳細欄に曜日ごとの内訳（労働時間・必要休憩・設定休憩）を表示。

### ⑧ 最低賃金（日給制の時間換算）

- `dailyH`（現行、`getDailyMin()`ベース）を `getAvgDailyMin() / 60` に置き換える。1-3と同じ平均1日時間を使う。
- 詳細欄の文言「1日の所定労働時間」は「平均1日所定労働時間」に変更する。

## 保存・移行（localStorage / Firestore 共通）

`getCurrentInputs()` / `saveAll()` / `restoreCheck()` / `loadAll()` に以下を追加する:

- `useWeekdaySchedule: gc('useWeekdaySchedule')`
- `daySchedule: Array.from({length:7}, (_,i) => ({ start: gv('dayStart'+i), end: gv('dayEnd'+i), break: gv('dayBreak'+i) }))`

復元側（`restoreCheck` / `loadAll`）:

- `useWeekdaySchedule` が未定義（旧データ）の場合は false 扱い（既存の `sc()` は undefined を無視するのでチェックボックスは自然に未チェックのまま＝変更不要）。
- `daySchedule` が配列であれば各曜日の `dayStart{i}`/`dayEnd{i}`/`dayBreak{i}` に復元する。配列がない（旧データ）場合は何もしない＝曜日別入力欄は空のまま。次に `useWeekdaySchedule` をONにしたとき、前述の「空なら共通値で初期化」ロジックが効き、共通値から自動的に補完される。

## テスト観点

- 共通モード（デフォルト）で①1-1/1-2/1-3、③3-1、⑧の判定・数値が現行と完全に一致すること（回帰確認）。
- 曜日別モードON、全曜日同じ値 → 共通モードと同じ結果になること。
- 曜日別モードON、一部曜日だけ短時間（例: 土曜3時間・休憩0分）→ 1-1は○のまま、1-2の週合計が正しく減ること、③3-1で土曜だけ休憩要件が「不要」判定になること。
- 曜日別モードON、ある曜日が8時間超（例: 月曜9時間）→ 1-1が✗になり、詳細欄にその曜日が明示されること。
- ④で休日にしている曜日は曜日別テーブルで入力不可・計算対象外になること。
- 曜日別モードON→OFF→ONと切り替えても、入力済みの曜日別値が消えないこと。
- 保存→リロードで曜日別データが復元されること。旧形式データ（`daySchedule`なし）を読み込んでもエラーにならず共通モードで開けること。
- localStorage・Firestore（クラウド保存）両方の保存/復元パスに新フィールドが反映されていること。

## スコープ外

- `generate_check_sheet.py`（Excel版）の曜日別対応。
- ⑤変形労働時間制タブとの連動強化（現状どおり別タブでの独立設定のまま）。
