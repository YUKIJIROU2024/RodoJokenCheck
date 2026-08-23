#!/usr/bin/env python3
"""ハローワーク 労働条件チェッカー — 統合Excel生成スクリプト

index.html の①〜⑧全項目の判定ロジックを、1計算ステップ1行の
詳細な数式＋日本語注釈というスタイルでExcelに再現する。
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from datetime import date

# ── 祝日テーブル 2025-2027 ────────────────────────────────────
HOLIDAYS = [
    (date(2025,1,1),"元日"),(date(2025,1,13),"成人の日"),(date(2025,2,11),"建国記念の日"),
    (date(2025,2,23),"天皇誕生日"),(date(2025,2,24),"天皇誕生日 振替"),(date(2025,3,20),"春分の日"),
    (date(2025,4,29),"昭和の日"),(date(2025,5,3),"憲法記念日"),(date(2025,5,4),"みどりの日"),
    (date(2025,5,5),"こどもの日"),(date(2025,5,6),"こどもの日 振替"),(date(2025,7,21),"海の日"),
    (date(2025,8,11),"山の日"),(date(2025,9,15),"敬老の日"),(date(2025,9,23),"秋分の日"),
    (date(2025,10,13),"スポーツの日"),(date(2025,11,3),"文化の日"),(date(2025,11,23),"勤労感謝の日"),
    (date(2025,11,24),"勤労感謝の日 振替"),
    (date(2026,1,1),"元日"),(date(2026,1,12),"成人の日"),(date(2026,2,11),"建国記念の日"),
    (date(2026,2,23),"天皇誕生日"),(date(2026,3,20),"春分の日"),(date(2026,4,29),"昭和の日"),
    (date(2026,5,3),"憲法記念日"),(date(2026,5,4),"みどりの日"),(date(2026,5,5),"こどもの日"),
    (date(2026,5,6),"こどもの日 振替"),(date(2026,7,20),"海の日"),(date(2026,8,11),"山の日"),
    (date(2026,9,21),"敬老の日"),(date(2026,9,22),"国民の休日"),(date(2026,9,23),"秋分の日"),
    (date(2026,10,12),"スポーツの日"),(date(2026,11,3),"文化の日"),(date(2026,11,23),"勤労感謝の日"),
    (date(2027,1,1),"元日"),(date(2027,1,11),"成人の日"),(date(2027,2,11),"建国記念の日"),
    (date(2027,2,23),"天皇誕生日"),(date(2027,3,21),"春分の日"),(date(2027,4,29),"昭和の日"),
    (date(2027,5,3),"憲法記念日"),(date(2027,5,4),"みどりの日"),(date(2027,5,5),"こどもの日"),
    (date(2027,7,19),"海の日"),(date(2027,8,11),"山の日"),(date(2027,8,12),"山の日 振替"),
    (date(2027,9,20),"敬老の日"),(date(2027,9,23),"秋分の日"),(date(2027,10,11),"スポーツの日"),
    (date(2027,11,3),"文化の日"),(date(2027,11,23),"勤労感謝の日"),(date(2027,11,24),"勤労感謝の日 振替"),
]
# 合計55件 → '祝日'!$A$2:$A$56

# ── 都道府県別最低賃金（2024年10月改定値） ───────────────────
PREF_MIN_WAGES = [
    ("北海道",1010),("青森",953),("岩手",952),("宮城",1055),("秋田",951),
    ("山形",955),("福島",955),("茨城",1005),("栃木",1004),("群馬",985),
    ("埼玉",1078),("千葉",1076),("東京",1163),("神奈川",1162),("新潟",985),
    ("富山",1001),("石川",984),("福井",984),("山梨",988),("長野",998),
    ("岐阜",1001),("静岡",1034),("愛知",1077),("三重",1023),("滋賀",1017),
    ("京都",1058),("大阪",1177),("兵庫",1001),("奈良",986),("和歌山",980),
    ("鳥取",957),("島根",962),("岡山",982),("広島",1020),("山口",979),
    ("徳島",980),("香川",970),("愛媛",956),("高知",952),("福岡",992),
    ("佐賀",956),("長崎",953),("熊本",952),("大分",954),("宮崎",951),
    ("鹿児島",953),("沖縄",952),
]

# ── 年次有給休暇の法定付与日数テーブル（労基法第39条） ────────
FULL_LEAVE = [
    ("6ヶ月",10),("1年6ヶ月",11),("2年6ヶ月",12),("3年6ヶ月",14),
    ("4年6ヶ月",16),("5年6ヶ月",18),("6年6ヶ月以上",20),
]
PART_LEAVE = {
    4: [("6ヶ月",7),("1年6ヶ月",8),("2年6ヶ月",9),("3年6ヶ月",10),("4年6ヶ月",12),("5年6ヶ月",13),("6年6ヶ月以上",15)],
    3: [("6ヶ月",5),("1年6ヶ月",6),("2年6ヶ月",6),("3年6ヶ月",8),("4年6ヶ月",9),("5年6ヶ月",10),("6年6ヶ月以上",11)],
    2: [("6ヶ月",3),("1年6ヶ月",4),("2年6ヶ月",4),("3年6ヶ月",5),("4年6ヶ月",6),("5年6ヶ月",6),("6年6ヶ月以上",7)],
    1: [("6ヶ月",1),("1年6ヶ月",2),("2年6ヶ月",2),("3年6ヶ月",2),("4年6ヶ月",3),("5年6ヶ月",3),("6年6ヶ月以上",3)],
}

# ── スタイル定数 ──────────────────────────────────────────────
HDR_FILL  = PatternFill(start_color='1A3A6E', end_color='1A3A6E', fill_type='solid')
HDR_FONT  = Font(bold=True, color='FFFFFF', size=12)
SEC_FILL  = PatternFill(start_color='EEF2F7', end_color='EEF2F7', fill_type='solid')
SEC_FONT  = Font(bold=True, color='1A3A6E', size=10)
INP_FILL  = PatternFill(start_color='EBF4FA', end_color='EBF4FA', fill_type='solid')
CALC_FILL = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
CALC_C_FILL = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
SEP_FILL  = PatternFill(start_color='E8EDF2', end_color='E8EDF2', fill_type='solid')
INFO_FILL = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
OK_FILL   = PatternFill(start_color='D1FAE5', end_color='D1FAE5', fill_type='solid')
WARN_FILL = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
NG_FILL   = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
th = Side(style='thin', color='CBD5E0')
BORDER = Border(left=th, right=th, top=th, bottom=th)


def status_font(color, size):
    return Font(bold=True, color=color, size=size)


# ── 行アドレッシング・詳細行スタイルを両立するビルダー ─────────
class SheetBuilder:
    """名前付きの行カーソルで、1行1呼び出しの詳細スタイル（ラベル・数式・
    注釈をまとめて出力）と、後続の数式から前の行を名前で参照できる仕組みを両立する。"""

    def __init__(self, ws):
        self.ws = ws
        self._row = 0
        self.R = {}
        self.judgments = []  # [{code, title, name, detail}]
        self._anon = 0

    def _next_row(self, name=None):
        if name is not None and name in self.R:
            return self.R[name]  # already reserved via reserve()
        self._row += 1
        if name is None:
            self._anon += 1
            name = f'_anon{self._anon}'
        self.R[name] = self._row
        return self._row

    def reserve(self, name):
        """Pre-allocate a row number for `name` so a formula/narration built
        before the row's cb()/judgment_row() call can still self-reference it."""
        return self._next_row(name)

    def ref(self, name, col='B'):
        return f"${col}${self.R[name]}"

    def title(self, text):
        r = self._next_row()
        c = self.ws.cell(row=r, column=1)
        c.value = text
        c.font = Font(bold=True, color='FFFFFF', size=14)
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='left', vertical='center')
        self.ws.merge_cells(f'A{r}:C{r}')
        self.ws.row_dimensions[r].height = 30

    def section_hdr(self, text):
        r = self._next_row()
        c = self.ws.cell(row=r, column=1)
        c.value = text
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='left', vertical='center')
        self.ws.merge_cells(f'A{r}:C{r}')
        self.ws.row_dimensions[r].height = 22

    def input_sec(self, text):
        r = self._next_row()
        c = self.ws.cell(row=r, column=1)
        c.value = text
        c.font = SEC_FONT
        c.fill = SEC_FILL
        self.ws.merge_cells(f'A{r}:C{r}')
        self.ws.row_dimensions[r].height = 18

    def sep_row(self, text):
        r = self._next_row()
        c = self.ws.cell(row=r, column=1)
        c.value = text
        c.font = Font(bold=True, size=9, color='374151')
        c.fill = SEP_FILL
        self.ws.merge_cells(f'A{r}:C{r}')
        self.ws.row_dimensions[r].height = 16

    def note_row(self, text):
        r = self._next_row()
        c = self.ws.cell(row=r, column=1)
        c.value = text
        c.font = Font(size=9, color='4B5563')
        c.alignment = Alignment(wrap_text=True)
        self.ws.merge_cells(f'A{r}:C{r}')

    def _dv(self, addr, choices=None, choices_range=None):
        if choices is not None:
            v = DataValidation(type="list", formula1=f'"{",".join(choices)}"', showDropDown=False)
        else:
            v = DataValidation(type="list", formula1=choices_range, showDropDown=False)
        self.ws.add_data_validation(v)
        v.add(addr)

    def inp(self, name, label, value, fmt=None, note=None, choices=None, choices_range=None):
        r = self._next_row(name)
        self.ws.cell(row=r, column=1).value = label
        self.ws.cell(row=r, column=1).font = Font(size=10)
        c = self.ws.cell(row=r, column=2)
        c.value = value
        c.fill = INP_FILL
        c.border = BORDER
        if fmt:
            c.number_format = fmt
        if note:
            nc = self.ws.cell(row=r, column=3)
            nc.value = note
            nc.font = Font(size=9, color='718096')
            nc.fill = CALC_C_FILL
        if choices is not None:
            self._dv(f'B{r}', choices=choices)
        elif choices_range is not None:
            self._dv(f'B{r}', choices_range=choices_range)
        return c

    def cb(self, name, label, formula, narration=None, fmt='#,##0.00'):
        r = self._next_row(name)
        self.ws.cell(row=r, column=1).value = label
        self.ws.cell(row=r, column=1).font = Font(size=10)
        c = self.ws.cell(row=r, column=2)
        c.value = formula
        c.fill = CALC_FILL
        c.border = BORDER
        c.font = Font(size=10)
        c.number_format = fmt
        if narration:
            nc = self.ws.cell(row=r, column=3)
            nc.value = narration
            nc.fill = CALC_C_FILL
            nc.font = Font(size=9, color='4B5563')
            nc.alignment = Alignment(wrap_text=True)
        return c

    def judgment_row(self, name, code, title, formula, detail=None):
        r = self._next_row(name)
        ac = self.ws.cell(row=r, column=1)
        ac.value = f'▶ {code} 判定（{title}）'
        ac.font = Font(bold=True, size=11, color='1A3A6E')
        ac.fill = INFO_FILL
        bc = self.ws.cell(row=r, column=2)
        bc.value = formula
        bc.font = Font(bold=True, size=16)
        bc.alignment = Alignment(horizontal='center', vertical='center')
        bc.border = BORDER
        bc.fill = INFO_FILL
        self.ws.cell(row=r, column=3).fill = INFO_FILL
        self.ws.row_dimensions[r].height = 30
        apply_judgment_cf(self.ws, f'B{r}:B{r}', size=16)
        self.judgments.append({'code': code, 'title': title, 'name': name, 'detail': detail})
        return bc


def apply_judgment_cf(ws, rng, size):
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"○"'], fill=OK_FILL, font=status_font('065F46', size)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"△"'], fill=WARN_FILL, font=status_font('92400E', size)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"✗"'], fill=NG_FILL, font=status_font('991B1B', size)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"参考"'], fill=INFO_FILL, font=Font(color='4B5563', size=size)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"－"'], fill=INFO_FILL, font=Font(color='9CA3AF', size=size)))


def xref(b, name, col='B'):
    return f"'チェックシート'!${col}${b.R[name]}"


# ── チェックシート（本体） ───────────────────────────────────
def build_check_sheet(b):
    ws = b.ws
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 46

    HR = "'祝日'!$A$2:$A$56"

    b.title('ハローワーク 労働条件チェッカー（Excel版）')
    b.inp('bizName', '事業所名', '')
    b.inp('checkDate', 'チェック日', date.today(), fmt='YYYY/MM/DD')

    # ── 共通設定・共通計算 ──────────────────────────────────
    b.input_sec('◆ 共通設定・共通計算')
    b.inp('specialMeasure', '特例措置対象事業場（10人未満の特定業種）', 'いいえ',
          note='商業・映画演劇業・保健衛生業・接客娯楽業 → 週44時間まで',
          choices=['はい', 'いいえ'])
    b.cb('weeklyLimit', '週の法定労働時間の上限（時間）',
         f'=IF({b.ref("specialMeasure")}="はい",44,40)', fmt='0',
         narration=f'=IF({b.ref("specialMeasure")}="はい","特例措置対象 → 44時間","通常 → 40時間")')

    b.inp('startTime', '始業時刻', 9/24, fmt='h:mm')
    b.inp('endTime', '終業時刻', 18/24, fmt='h:mm')
    b.inp('breakTime', '休憩時間（分）', 60)
    b.cb('startMin', '始業（分）', f'={b.ref("startTime")}*24*60', fmt='0')
    b.cb('endMin', '終業（分）',
         f'=IF({b.ref("endTime")}<{b.ref("startTime")},({b.ref("endTime")}+1)*24*60,{b.ref("endTime")}*24*60)',
         fmt='0',
         narration=f'=IF({b.ref("endTime")}<{b.ref("startTime")},"翌日にかかる勤務として計算","通常勤務")')
    b.reserve('dailyMin')
    b.cb('dailyMin', '1日の所定労働時間（分）',
         f'=MAX(0,{b.ref("endMin")}-{b.ref("startMin")}-{b.ref("breakTime")})', fmt='0',
         narration=f'=({b.ref("endMin")}-{b.ref("startMin")})&"分 − 休憩"&{b.ref("breakTime")}&"分 = "&{b.ref("dailyMin")}&"分"')
    b.reserve('dailyH')
    b.cb('dailyH', '1日の所定労働時間（時間）', f'={b.ref("dailyMin")}/60', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間"')

    wd_labels = ['月', '火', '水', '木', '金', '土', '日']
    wd_defaults = ['いいえ', 'いいえ', 'いいえ', 'いいえ', 'いいえ', 'はい', 'はい']
    for idx, (day, dflt) in enumerate(zip(wd_labels, wd_defaults)):
        b.inp(f'wd{idx}', f'{day}曜日（休日）', dflt, choices=['はい', 'いいえ'])

    b.reserve('weeklyOffCount')
    woc_parts = '+'.join(f'IF({b.ref(f"wd{d}")}="はい",1,0)' for d in range(7))
    b.cb('weeklyOffCount', '週休日数', f'={woc_parts}', fmt='0',
         narration=f'={b.ref("weeklyOffCount")}&"日/週"')
    b.cb('weeklyWorkDays', '週労働日数', f'=7-{b.ref("weeklyOffCount")}', fmt='0')
    b.reserve('weeklyH')
    b.cb('weeklyH', '週の所定労働時間（時間）', f'={b.ref("dailyH")}*{b.ref("weeklyWorkDays")}', fmt='0.00',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間/日 × "&{b.ref("weeklyWorkDays")}&"日 = "&TEXT({b.ref("weeklyH")},"0.00")&"時間/週"')

    b.inp('holidayType', '祝日の扱い', '休日扱い', choices=['休日扱い', '出勤日'])
    b.inp('addHolidays', '追加休日日数（週休・祝日以外）', 0, note='夏季・年末年始等')
    b.inp('holidayRate', '法定休日労働の割増率（%）', 35,
          note='④-3・⑦-2で使用（法定最低35%）')

    b.cb('targetYear', '対象年度', '=YEAR(TODAY())', fmt='0',
         narration='="祝日カウントの基準年（必要に応じて上書き可）"')

    WD = [b.ref(f'wd{d}') for d in range(7)]
    pub_parts = '+'.join(
        f'SUMPRODUCT((YEAR({HR})={b.ref("targetYear")})*(WEEKDAY({HR},2)={d+1})*({WD[d]}<>"はい"))'
        for d in range(7)
    )
    b.reserve('pubHolOnWorkDays')
    b.cb('pubHolOnWorkDays', '祝日（就労日に当たる分）',
         f'=IF({b.ref("holidayType")}="休日扱い",{pub_parts},0)', fmt='0',
         narration=f'=IF({b.ref("holidayType")}="休日扱い",{b.ref("pubHolOnWorkDays")}&"日","出勤日扱い（加算なし）")')
    b.reserve('annualOffFromWeekly')
    b.cb('annualOffFromWeekly', '年間休日（週休分）', f'={b.ref("weeklyOffCount")}*52', fmt='0',
         narration=f'={b.ref("weeklyOffCount")}&"日 × 52週 = "&{b.ref("annualOffFromWeekly")}&"日"')
    b.reserve('annualOffDays')
    b.cb('annualOffDays', '年間休日日数（合計）',
         f'={b.ref("annualOffFromWeekly")}+{b.ref("pubHolOnWorkDays")}+{b.ref("addHolidays")}', fmt='0',
         narration=f'={b.ref("annualOffFromWeekly")}&"日＋"&{b.ref("pubHolOnWorkDays")}&"日（祝日）＋"&{b.ref("addHolidays")}&"日（追加）＝"&{b.ref("annualOffDays")}&"日"')
    b.reserve('annualWorkDays')
    b.cb('annualWorkDays', '年間就業日数', f'=MAX(0,365-{b.ref("annualOffDays")})', fmt='0',
         narration=f'="365 − "&{b.ref("annualOffDays")}&" = "&{b.ref("annualWorkDays")}&"日"')
    b.reserve('annualWorkH')
    b.cb('annualWorkH', '年間総労働時間', f'={b.ref("dailyH")}*{b.ref("annualWorkDays")}', fmt='0.0',
         narration=f'=TEXT({b.ref("dailyH")},"0.00")&"時間 × "&{b.ref("annualWorkDays")}&"日 = "&TEXT({b.ref("annualWorkH")},"0.0")&"時間"')
    b.reserve('legalAnnualH')
    b.cb('legalAnnualH', '年間法定労働時間の上限', f'={b.ref("weeklyLimit")}*52', fmt='0',
         narration=f'={b.ref("weeklyLimit")}&"時間 × 52週 = "&{b.ref("legalAnnualH")}&"時間"')

    # ── ① 労働時間 ───────────────────────────────────────────
    b.section_hdr('① 労働時間')
    b.judgment_row('j1_1', '1-1', '1日の所定労働時間',
        f'=IF({b.ref("dailyH")}<=8,"○","✗")',
        detail=f'="1日労働時間: "&TEXT({xref(b,"dailyH")},"0.00")&"時間（上限8時間）"')
    b.judgment_row('j1_2', '1-2', '週の所定労働時間',
        f'=IF({b.ref("weeklyH")}<={b.ref("weeklyLimit")},"○","✗")',
        detail=f'="週労働時間: "&TEXT({xref(b,"weeklyH")},"0.00")&"時間（上限"&{xref(b,"weeklyLimit")}&"時間）"')
    b.judgment_row('j1_3', '1-3', '年間総労働時間（参考）', '="参考"',
        detail=f'="年間就業: "&{xref(b,"annualWorkDays")}&"日 / 年間労働: "&TEXT({xref(b,"annualWorkH")},"0.0")&"時間"')

    # ── ② 時間外・休日労働（36協定） ───────────────────────
    b.section_hdr('② 時間外・休日労働（36協定）')
    b.inp('s36', '36協定', '時間外なし', choices=['時間外なし', '届出済', '未届出'])
    b.inp('monthlyOT', '月間時間外上限（時間）', 45, note='法定上限45h')
    b.inp('annualOT', '年間時間外上限（時間）', 360, note='法定上限360h')
    b.inp('specialClause', '特別条項', 'なし', choices=['なし', 'あり'])
    b.inp('specialMonthOT', '特別条項 月間上限（時間）', 80, note='100h未満')
    b.inp('specialAnnualOT', '特別条項 年間上限（時間）', 720, note='720h以内')
    b.note_row('※ 参考: 特別条項適用時は「複数月（2〜6ヶ月）平均80時間以内」「年6回以内」の制限もあります（本表では未算定）')

    b.judgment_row('j2_1', '2-1', '36協定の締結・届出',
        f'=IF({b.ref("s36")}="未届出","✗","○")',
        detail=f'={xref(b,"s36")}')
    b.judgment_row('j2_2', '2-2', '時間外労働の上限（通常条項）',
        f'=IF({b.ref("s36")}="時間外なし","○",IF(AND({b.ref("monthlyOT")}<=45,{b.ref("annualOT")}<=360),"○","✗"))',
        detail=f'="月: "&{xref(b,"monthlyOT")}&"h（45h以内） / 年: "&{xref(b,"annualOT")}&"h（360h以内）"')
    b.judgment_row('j2_3', '2-3', '時間外労働の上限（特別条項）',
        f'=IF(OR({b.ref("s36")}="時間外なし",{b.ref("specialClause")}="なし"),"○",'
        f'IF(AND({b.ref("specialMonthOT")}<100,{b.ref("specialAnnualOT")}<=720),"○","✗"))',
        detail=f'=IF(OR({xref(b,"s36")}="時間外なし",{xref(b,"specialClause")}="なし"),"該当なし",'
               f'"月: "&{xref(b,"specialMonthOT")}&"h（100h未満） / 年: "&{xref(b,"specialAnnualOT")}&"h（720h以内）")')

    # ── ③ 休憩 ───────────────────────────────────────────────
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

    # ── ④ 休日 ───────────────────────────────────────────────
    b.section_hdr('④ 休日')
    b.judgment_row('j4_1', '4-1', '法定休日（週1日以上）',
        f'=IF({b.ref("weeklyOffCount")}>=1,"○","✗")',
        detail=f'="週休: "&{xref(b,"weeklyOffCount")}&"日/週"')
    b.judgment_row('j4_2', '4-2', '年間休日日数（参考）',
        f'=IF({b.ref("annualOffDays")}>=105,"○",IF({b.ref("annualOffDays")}>=52,"△","✗"))',
        detail=f'="年間休日: "&{xref(b,"annualOffDays")}&"日（目安105日以上）"')
    b.judgment_row('j4_3', '4-3', '法定休日労働の割増賃金',
        f'=IF({b.ref("holidayRate")}>=35,"○","✗")',
        detail=f'="休日割増率: "&{xref(b,"holidayRate")}&"%（法定35%以上）"')

    # ── ⑤ 変形労働時間制 ─────────────────────────────────────
    b.section_hdr('⑤ 変形労働時間制')
    b.inp('flexType', '変形制の種類', 'なし', choices=['なし', '1ヶ月単位', '1年単位', 'フレックス'])

    b.input_sec('1ヶ月単位変形制の設定')
    b.inp('flexMonAgreement', '就業規則/労使協定', '有・対応済', choices=['有・対応済', 'なし'])
    b.inp('flexMonReport', '労基署への届出', '届出済', choices=['届出済', '未届出'])
    b.inp('flexMonTotal', '1ヶ月の総所定労働時間（時間）', 0, note='0=未設定')
    b.reserve('flexMonMaxH')
    b.cb('flexMonMaxH', '1ヶ月単位の月間上限（時間）', f'={b.ref("weeklyLimit")}*4.348', fmt='0.0',
         narration=f'={b.ref("weeklyLimit")}&"時間 × 4.348週 ≒ "&TEXT({b.ref("flexMonMaxH")},"0.0")&"時間"')

    b.input_sec('1年単位変形制の設定')
    b.inp('flexAnnAgreement', '労使協定・届出', '有・届出済', choices=['有・届出済', 'なし/未届出'])
    b.inp('flexAnnMaxDay', '1日の最長労働時間（時間）', 0, note='上限10h、0=未設定')
    b.inp('flexAnnMaxWeek', '週の最長労働時間（時間）', 0, note='上限52h、0=未設定')

    b.input_sec('フレックスタイム制の設定')
    b.inp('flexFlexPeriod', '清算期間（月数）', '1', choices=['1', '2', '3'])
    b.inp('flexFlexTotal', '清算期間の総所定労働時間（時間）', 0, note='0=未設定')
    b.reserve('flexFlexLegalMax')
    b.cb('flexFlexLegalMax', 'フレックスの清算期間上限（時間）',
         f'=40*VALUE({b.ref("flexFlexPeriod")})*30/7', fmt='0.0',
         narration=f'="40時間 × "&{b.ref("flexFlexPeriod")}&"ヶ月×30日÷7 ≒ "&TEXT({b.ref("flexFlexLegalMax")},"0.0")&"時間"')

    b.cb('j5_1m', '5-1判定（1ヶ月単位のみ）',
         f'=IF(AND({b.ref("flexMonAgreement")}="有・対応済",{b.ref("flexMonReport")}="届出済"),"○","✗")')
    b.cb('j5_1a', '5-1判定（1年単位のみ）',
         f'=IF({b.ref("flexAnnAgreement")}="有・届出済","○","✗")')
    b.judgment_row('j5_1', '5-1', '変形制の必要書類の整備',
        f'=IF({b.ref("flexType")}="なし","○",'
        f'IF({b.ref("flexType")}="1ヶ月単位",{b.ref("j5_1m")},'
        f'IF({b.ref("flexType")}="1年単位",{b.ref("j5_1a")},"○")))',
        detail=f'="変形制の種類: "&{xref(b,"flexType")}')

    b.cb('j5_2m', '5-2判定（1ヶ月単位のみ）',
         f'=IF({b.ref("flexMonTotal")}=0,"△",IF({b.ref("flexMonTotal")}<={b.ref("flexMonMaxH")},"○","✗"))')
    b.cb('j5_2a', '5-2判定（1年単位のみ）',
         f'=IF(AND(OR({b.ref("flexAnnMaxDay")}=0,{b.ref("flexAnnMaxDay")}<=10),'
         f'OR({b.ref("flexAnnMaxWeek")}=0,{b.ref("flexAnnMaxWeek")}<=52)),"○","✗")')
    b.cb('j5_2f', '5-2判定（フレックスのみ）',
         f'=IF({b.ref("flexFlexTotal")}=0,"△",IF({b.ref("flexFlexTotal")}<={b.ref("flexFlexLegalMax")},"○","✗"))')
    b.judgment_row('j5_2', '5-2', '変形制の1日・週の上限時間',
        f'=IF({b.ref("flexType")}="なし","○",'
        f'IF({b.ref("flexType")}="1ヶ月単位",{b.ref("j5_2m")},'
        f'IF({b.ref("flexType")}="1年単位",{b.ref("j5_2a")},{b.ref("j5_2f")})))',
        detail=f'=IF({xref(b,"flexType")}="なし","該当なし",'
               f'IF({xref(b,"flexType")}="1ヶ月単位","月間上限: "&TEXT({xref(b,"flexMonMaxH")},"0.0")&"h",'
               f'IF({xref(b,"flexType")}="1年単位","日上限10h／週上限52h",'
               f'"清算期間上限: "&TEXT({xref(b,"flexFlexLegalMax")},"0.0")&"h")))')
    b.judgment_row('j5_3', '5-3', '変形制の平均週労働時間', '="○"',
        detail=f'=IF({xref(b,"flexType")}="なし","該当なし","就業規則・労使協定の内容に基づき管理")')

    # ── ⑥ 年次有給休暇 ───────────────────────────────────────
    b.section_hdr('⑥ 年次有給休暇（参考）')
    b.inp('empType', '雇用形態', '正社員・フルタイム', choices=['正社員・フルタイム', 'パート・アルバイト'])
    b.inp('partDays', '週所定労働日数（パートのみ）', '4', choices=['4', '3', '2', '1'])
    b.inp('hourlyLeave', '時間単位年休の採用', 'なし', choices=['なし', 'あり'])

    for t, (tenure, full_days) in enumerate(FULL_LEAVE):
        part_days_by_n = {n: PART_LEAVE[n][t][1] for n in (1, 2, 3, 4)}
        formula = (
            f'=IF({b.ref("empType")}="正社員・フルタイム",{full_days},'
            f'IF({b.ref("partDays")}="4",{part_days_by_n[4]},'
            f'IF({b.ref("partDays")}="3",{part_days_by_n[3]},'
            f'IF({b.ref("partDays")}="2",{part_days_by_n[2]},{part_days_by_n[1]}))))'
        )
        name = f'leaveDays{t}'
        b.reserve(name)
        b.cb(name, f'法定付与日数（勤続{tenure}）', formula, fmt='0',
             narration=f'="勤続{tenure}: "&{b.ref(name)}&"日"')

    b.judgment_row('j6_1', '6-1', '法定付与日数（参考）', '="参考"',
        detail=f'="雇用形態: "&{xref(b,"empType")}&" / パート週日数: "&{xref(b,"partDays")}')
    b.judgment_row('j6_2', '6-2', '時間単位年休', '="参考"',
        detail=f'="時間単位年休: "&{xref(b,"hourlyLeave")}')

    # ── ⑦ 割増賃金 ───────────────────────────────────────────
    b.section_hdr('⑦ 割増賃金')
    b.inp('otRate', '時間外割増率（%）', 25, note='法定最低25%。⑧-3の固定残業代計算でも使用')
    b.inp('nightRate', '深夜割増率（%）', 25, note='法定最低25%')
    b.inp('ot60Rate', '月60時間超割増率（%）', 50, note='法定最低50%（中小企業も対象）')

    b.judgment_row('j7_1', '7-1', '時間外・深夜割増率',
        f'=IF(AND({b.ref("otRate")}>=25,{b.ref("nightRate")}>=25,{b.ref("otRate")}+{b.ref("nightRate")}>=50),"○","✗")',
        detail=f'="時間外: "&{xref(b,"otRate")}&"% / 深夜: "&{xref(b,"nightRate")}&"% / 合計: "&({xref(b,"otRate")}+{xref(b,"nightRate")})&"%"')
    b.judgment_row('j7_2', '7-2', '法定休日労働の割増率',
        f'=IF(AND({b.ref("holidayRate")}>=35,{b.ref("holidayRate")}+{b.ref("nightRate")}>=60),"○","✗")',
        detail=f'="休日: "&{xref(b,"holidayRate")}&"% / 休日+深夜: "&({xref(b,"holidayRate")}+{xref(b,"nightRate")})&"%"')
    b.judgment_row('j7_3', '7-3', '月60時間超の割増率',
        f'=IF({b.ref("ot60Rate")}>=50,"○","✗")',
        detail=f'="月60h超割増: "&{xref(b,"ot60Rate")}&"%（法定50%以上）"')

    # ── ⑧ 最低賃金・固定残業代 ───────────────────────────────
    b.section_hdr('⑧ 最低賃金・固定残業代')
    b.inp('prefecture', '都道府県', '大阪', note='選択すると最低賃金を自動取得',
          choices_range="'設定'!$A$2:$A$48")
    b.cb('minWage', '最低賃金（円/時間）',
         f"=IFERROR(VLOOKUP({b.ref('prefecture')},'設定'!$A:$B,2,FALSE),0)", fmt='#,##0',
         narration='="自動取得（B列を直接上書きすれば変更可）"')
    b.inp('wageType', '賃金の種類', '月給制', choices=['月給制', '日給制', '時給制'])
    b.inp('monthlySalary', '月給（所定内賃金合計）（円）', 0, fmt='#,##0', note='固定残業代を含む合計額')
    b.inp('fixedOT', 'うち固定残業代（月額）（円）', 0, fmt='#,##0', note='0=なし')
    b.inp('fixedOTHours', '固定残業の設定時間/月（時間）', 0, fmt='0.0', note='0=なし')
    b.inp('dailySalary', '日給（円）', 0, fmt='#,##0')
    b.inp('hourlySalary', '時給（円）', 0, fmt='#,##0')

    b.sep_row('── 8-1：基本給のみの最低賃金比較 ──')
    b.cb('basicMonthly', '月額基本給（固定残業代除く）',
         f'=IF({b.ref("wageType")}="月給制",{b.ref("monthlySalary")}-IF({b.ref("fixedOT")}>0,{b.ref("fixedOT")},0),'
         f'IF({b.ref("wageType")}="日給制",{b.ref("dailySalary")},{b.ref("hourlySalary")}))', fmt='#,##0')
    b.cb('annualBasic', '年額基本給（月給制のみ）',
         f'=IF({b.ref("wageType")}="月給制",{b.ref("basicMonthly")}*12,0)', fmt='#,##0')
    b.reserve('hourlyRateBasic')
    b.cb('hourlyRateBasic', '基本給の時間換算額（円/時間）',
         f'=IF({b.ref("wageType")}="月給制",IF({b.ref("annualWorkH")}>0,{b.ref("annualBasic")}/{b.ref("annualWorkH")},0),'
         f'IF({b.ref("wageType")}="日給制",IF({b.ref("dailyH")}>0,{b.ref("dailySalary")}/{b.ref("dailyH")},0),{b.ref("hourlySalary")}))',
         fmt='#,##0.00',
         narration=f'=TEXT({b.ref("hourlyRateBasic")},"0.00")&"円/時間（最低賃金 "&{b.ref("minWage")}&"円）"')
    b.judgment_row('j8_1', '8-1', '基本給のみの最低賃金比較',
        f'=IF({b.ref("wageType")}="月給制",'
        f'IF(OR({b.ref("monthlySalary")}=0,{b.ref("monthlySalary")}=""),"△",'
        f'IF({b.ref("annualWorkH")}>0,IF({b.ref("hourlyRateBasic")}>={b.ref("minWage")},"○","✗"),"△")),'
        f'IF({b.ref("wageType")}="日給制",'
        f'IF(OR({b.ref("dailySalary")}=0,{b.ref("dailySalary")}=""),"△",'
        f'IF({b.ref("dailyH")}>0,IF({b.ref("hourlyRateBasic")}>={b.ref("minWage")},"○","✗"),"△")),'
        f'IF(OR({b.ref("hourlySalary")}=0,{b.ref("hourlySalary")}=""),"△",'
        f'IF({b.ref("hourlySalary")}>={b.ref("minWage")},"○","✗"))))',
        detail=f'="時間換算額: "&TEXT({xref(b,"hourlyRateBasic")},"0.00")&"円/h（最低賃金 "&{xref(b,"minWage")}&"円）"')

    b.sep_row('── 8-2：総額（固定残業代込み）の最低賃金比較 ──')
    b.cb('annualTotal', '年額総給与（月給制・固定残業代あり時のみ）',
         f'=IF(AND({b.ref("wageType")}="月給制",{b.ref("fixedOT")}>0),{b.ref("monthlySalary")}*12,0)', fmt='#,##0')
    b.reserve('hourlyRateTotal')
    b.cb('hourlyRateTotal', '総給与の時間換算額（円/時間）',
         f'=IF(AND({b.ref("wageType")}="月給制",{b.ref("fixedOT")}>0,{b.ref("annualWorkH")}>0),'
         f'{b.ref("annualTotal")}/{b.ref("annualWorkH")},0)', fmt='#,##0.00',
         narration=f'=IF(AND({b.ref("wageType")}="月給制",{b.ref("fixedOT")}>0),'
                    f'TEXT({b.ref("hourlyRateTotal")},"0.00")&"円/時間","該当なし（固定残業代なし）")')
    b.reserve('j8_2')
    b.judgment_row('j8_2', '8-2', '総額（固定残業代込み）の最低賃金比較',
        f'=IF(AND({b.ref("wageType")}="月給制",{b.ref("fixedOT")}>0),'
        f'IF({b.ref("annualWorkH")}>0,IF({b.ref("hourlyRateTotal")}>={b.ref("minWage")},"○","✗"),"△"),"－")',
        detail=f'=IF({xref(b,"j8_2")}="－","該当なし（固定残業代なし）",'
               f'"総給与時換算: "&TEXT({xref(b,"hourlyRateTotal")},"0.00")&"円/h（最低賃金 "&{xref(b,"minWage")}&"円）")')

    b.sep_row('── 8-3：固定残業代の適正額チェック ──')
    b.reserve('fixedOTRequired')
    b.cb('fixedOTRequired', '必要な固定残業代額（8-1の時間換算額ベース）',
         f'={b.ref("fixedOTHours")}*{b.ref("hourlyRateBasic")}*(1+{b.ref("otRate")}/100)', fmt='#,##0',
         narration=(f'=TEXT({b.ref("fixedOTHours")},"0.0")&"h × "&TEXT({b.ref("hourlyRateBasic")},"0.00")&"円 × "&'
                    f'TEXT(1+{b.ref("otRate")}/100,"0.00")&"（割増"&{b.ref("otRate")}&"%） = "&TEXT({b.ref("fixedOTRequired")},"#,##0")&"円"'))
    b.reserve('j8_3')
    b.judgment_row('j8_3', '8-3', '固定残業代の適正額チェック',
        f'=IF(AND({b.ref("wageType")}="月給制",{b.ref("fixedOT")}>0,{b.ref("fixedOTHours")}>0),'
        f'IF({b.ref("fixedOT")}>={b.ref("fixedOTRequired")},"○","✗"),"－")',
        detail=f'=IF({xref(b,"j8_3")}="－","該当なし",'
               f'"必要額: "&TEXT({xref(b,"fixedOTRequired")},"#,##0")&"円 / 設定額: "&TEXT({xref(b,"fixedOT")},"#,##0")&"円")')


# ── チェック結果シート（集計） ───────────────────────────────
def build_results_sheet(ws, b):
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 56

    def hrow(row, text, merge='A:D'):
        c = ws.cell(row=row, column=1)
        c.value = text
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.merge_cells(f'{merge[0]}{row}:{merge[-1]}{row}')
        ws.row_dimensions[row].height = 24

    hrow(1, 'ハローワーク 労働条件チェッカー — チェック結果')
    ws['A2'] = f"=IFERROR({xref(b,'bizName')}&\" ／ チェック日: \"&TEXT({xref(b,'checkDate')},\"YYYY/MM/DD\"),\"\")"
    ws['A2'].font = Font(bold=True, size=11)
    ws.merge_cells('A2:D2')

    n = len(b.judgments)
    first_row, last_row = 5, 4 + n
    j_range = f"C{first_row}:C{last_row}"
    ws['A3'] = (f'="判定集計: ○ "&(COUNTIF({j_range},"○")+COUNTIF({j_range},"参考"))&"件  '
                f'△ "&COUNTIF({j_range},"△")&"件  ✗ "&COUNTIF({j_range},"✗")&"件"')
    ws['A3'].font = Font(bold=True, size=10)
    ws.merge_cells('A3:D3')
    ws.row_dimensions[3].height = 20

    for col, hdr in enumerate(['No.', 'チェック項目', '判定', '詳細'], start=1):
        cell = ws.cell(row=4, column=col)
        cell.value = hdr
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill(start_color='2D3748', end_color='2D3748', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')

    for row_idx, item in enumerate(b.judgments, start=first_row):
        ws.cell(row=row_idx, column=1).value = item['code']
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal='center')
        ws.cell(row=row_idx, column=1).font = Font(size=10, color='4A5568')
        ws.cell(row=row_idx, column=2).value = item['title']
        ws.cell(row=row_idx, column=2).font = Font(size=10)

        j_cell = ws.cell(row=row_idx, column=3)
        j_cell.value = f"={xref(b, item['name'])}"
        j_cell.alignment = Alignment(horizontal='center', vertical='center')
        j_cell.font = Font(bold=True, size=13)

        d_cell = ws.cell(row=row_idx, column=4)
        d_cell.value = item['detail'] if item['detail'] else f"={xref(b, item['name'])}"
        d_cell.font = Font(size=10)
        d_cell.alignment = Alignment(wrap_text=True)

        for col in range(1, 5):
            ws.cell(row=row_idx, column=col).border = Border(bottom=Side(style='thin', color='E2E8F0'))
        ws.row_dimensions[row_idx].height = 18

    apply_judgment_cf(ws, j_range, size=13)


# ── 祝日シート ────────────────────────────────────────────────
def create_holidays_sheet(wb):
    ws = wb.create_sheet('祝日')
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 22
    ws['A1'] = '日付'; ws['A1'].font = Font(bold=True)
    ws['B1'] = '祝日名'; ws['B1'].font = Font(bold=True)
    for r, (d, name) in enumerate(HOLIDAYS, start=2):
        ws.cell(row=r, column=1).value = d
        ws.cell(row=r, column=1).number_format = 'YYYY/MM/DD'
        ws.cell(row=r, column=2).value = name


# ── 設定シート（都道府県別最低賃金） ───────────────────────────
def create_settings_sheet(wb):
    ws = wb.create_sheet('設定')
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 14
    hdr_fill = PatternFill(start_color='2D3748', end_color='2D3748', fill_type='solid')
    ws['A1'] = '都道府県'; ws['A1'].font = Font(bold=True, color='FFFFFF'); ws['A1'].fill = hdr_fill
    ws['B1'] = '最低賃金'; ws['B1'].font = Font(bold=True, color='FFFFFF'); ws['B1'].fill = hdr_fill
    for r, (pref, wage) in enumerate(PREF_MIN_WAGES, start=2):
        ws.cell(row=r, column=1).value = pref
        ws.cell(row=r, column=2).value = wage
        ws.cell(row=r, column=2).number_format = '#,##0'


# ── メイン ────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    ws_check = wb.active
    ws_check.title = 'チェックシート'

    create_holidays_sheet(wb)
    create_settings_sheet(wb)

    b = SheetBuilder(ws_check)
    build_check_sheet(b)

    ws_results = wb.create_sheet('チェック結果', 0)
    build_results_sheet(ws_results, b)

    path = '/mnt/c/Users/y_uej/CC-RodoJokenCheck/rodo-joken-check.xlsx'
    wb.save(path)
    print(f'生成完了: {path}（項目数: {len(b.judgments)}）')


if __name__ == '__main__':
    main()
