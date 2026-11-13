"""工数予実管理表のテンプレート(Excel)を生成するスクリプト。

使い方:
    uv run python src/create_template.py [YYYY] [M] [Excelファイル]
    ファイルが無ければ新規作成、あればそのブックに月シートを追加する。
    例) uv run python src/create_template.py 2026 11 工数予実管理表.xlsx

シート名は「YY年MM月」(例: 26年11月)。
メンバーごとに「予定」「実績」「差異」の3行を持ち、週単位で工数(人時)を入力する。
3~5行目に1日の定時工数の入力欄と、週ごとの稼働日数(日本の平日-祝日)・定時予定工数を持つ。
"""
import calendar
import re
import sys
from datetime import date, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
import holidays

CONTROL_SHEET = "シート追加"
FONT_NAME = "游ゴシック"
HOLIDAY_SHEET = "祝日"
HOLIDAY_NAME = "祝日リスト"  # 稼働日数の計算で除外する日付の範囲名
HOLIDAY_YEARS = range(2025, 2036)
DEFAULT_DAILY_HOURS = 7.5  # 1日の定時工数(h)の初期値
MEMBER_BLOCKS = 12  # メンバー記入枠の数
HEADER_ROW = 6
FIRST_ROW = HEADER_ROW + 1

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_HEADER = PatternFill("solid", fgColor="1F4E78")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")  # 入力セル
FILL_CALC = PatternFill("solid", fgColor="F2F2F2")  # 計算セル
FILL_TOTAL = PatternFill("solid", fgColor="DDEBF7")


def sheet_name(year: int, month: int) -> str:
    """シート名「YY年MM月」を返す。"""
    return f"{year % 100:02d}年{month:02d}月"


def month_weeks(year: int, month: int):
    """月に含まれる週(月曜始まり)の(開始日, 終了日)を月内に切り詰めて返す。"""
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])
    start = first
    weeks = []
    while start <= last:
        # 日曜までをその週の終わりとする
        end = min(start + timedelta(days=6 - start.weekday()), last)
        weeks.append((start, end))
        start = end + timedelta(days=1)
    return weeks


def _style(cell, bold=False, color="000000", fill=None, align="center", wrap=False):
    cell.font = Font(name=FONT_NAME, size=10, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border = BORDER
    if fill:
        cell.fill = fill


def holiday_dates():
    """日本の祝日(日付, 名称)を日付順で返す。"""
    return sorted(holidays.Japan(years=list(HOLIDAY_YEARS)).items())


def ensure_holiday_sheet(wb: Workbook):
    """祝日シートと範囲名「祝日リスト」を用意する(既にあれば何もしない)。"""
    if HOLIDAY_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(HOLIDAY_SHEET)
        ws["A1"], ws["B1"] = "日付", "名称"
        for c in ("A1", "B1"):
            _style(ws[c], bold=True, color="FFFFFF", fill=FILL_HEADER)
        for r, (d, nm) in enumerate(holiday_dates(), 2):
            ws.cell(r, 1, d).number_format = "yyyy/m/d"
            ws.cell(r, 2, nm)
        ws["D1"] = "稼働日数の計算で除外する日付です。会社休業日(年末年始など)や次の年の祝日は末尾に追記してください(A2:A400が対象)。"
        ws["D1"].font = Font(name=FONT_NAME, size=9, color="595959")
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 18
    if HOLIDAY_NAME not in wb.defined_names:
        wb.defined_names[HOLIDAY_NAME] = DefinedName(
            HOLIDAY_NAME, attr_text=f"{HOLIDAY_SHEET}!$A$2:$A$400")


def _build_capacity(ws, year, month, weeks, col_first):
    """1日の定時工数の入力欄と、週ごとの稼働日数・定時予定工数の行を作る(3~5行目)。"""
    L = get_column_letter
    n = len(weeks)
    col_total = col_first + n
    _style(ws.cell(3, 1, "1日の定時工数(h)"), bold=True, fill=FILL_CALC)
    c = ws.cell(3, 2, DEFAULT_DAILY_HOURS)
    _style(c, color="0000FF", fill=FILL_INPUT)
    c.number_format = "0.0"
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="24",
                        allow_blank=False, showErrorMessage=True,
                        errorTitle="入力エラー", error="0~24の数値(h)を入力してください")
    ws.add_data_validation(dv)
    dv.add("B3")
    m1 = ws.cell(3, 3, date(year, month, 1))
    m1.number_format = '"月初:"yyyy/m/d'
    m1.font = Font(name=FONT_NAME, size=9, color="999999")
    m1.alignment = Alignment(horizontal="center")
    for r, label in ((4, "稼働日数(平日)"), (5, "定時予定工数(h)")):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        _style(ws.cell(r, 1, label), bold=True, fill=FILL_CALC)
        ws.cell(r, 2).border = BORDER
    for j in range(n):
        cl = L(col_first + j)
        mon = f"$C$3-WEEKDAY($C$3,3)+7*{j}"
        ws.cell(4, col_first + j).value = (
            f"=NETWORKDAYS(MAX($C$3,{mon}),MIN(EOMONTH($C$3,0),{mon}+6),{HOLIDAY_NAME})")
        ws.cell(5, col_first + j).value = f"={cl}4*$B$3"
    for r, fmt in ((4, "0"), (5, "0.0")):
        rng = f"{L(col_first)}{r}:{L(col_total - 1)}{r}"
        ws.cell(r, col_total).value = f"=SUM({rng})"
        for cc in range(col_first, col_total + 1):
            cell = ws.cell(r, cc)
            cell.number_format = fmt
            _style(cell, bold=(cc == col_total), fill=FILL_CALC)


def build_sheet(wb: Workbook, year: int, month: int, ws=None, members=None):
    """月シートを作る。membersを渡すとメンバー名を引き継ぎ、記入例は入れない。"""
    ws = ws or wb.create_sheet()
    ws.title = sheet_name(year, month)
    weeks = month_weeks(year, month)
    n = len(weeks)
    col_first = 3  # C列から週
    col_total = col_first + n
    L = get_column_letter

    # タイトル・凡例
    ws["A1"] = f"{year}年{month}月 工数予実管理表"
    ws["A1"].font = Font(name=FONT_NAME, size=14, bold=True)
    ws["A2"] = "単位: 人時(h)。黄色のセルに入力してください(青文字=入力値、黒文字=自動計算)。"
    ws["E3"] = "差異 = 実績 − 予定(プラスは超過)。" + (
        "メンバー名は前月シートから引き継ぎました。" if members is not None
        else "1行目のブロックは記入例です。上書きして使ってください。")
    for a in ("A2", "E3"):
        ws[a].font = Font(name=FONT_NAME, size=9, color="595959")
    _build_capacity(ws, year, month, weeks, col_first)

    # ヘッダー
    heads = ["メンバー", "区分"] + [
        f"第{i}週\n{s.month}/{s.day}~{e.month}/{e.day}" for i, (s, e) in enumerate(weeks, 1)
    ] + ["月合計"]
    for i, h in enumerate(heads, 1):
        c = ws.cell(HEADER_ROW, i, h)
        _style(c, bold=True, color="FFFFFF", fill=FILL_HEADER, wrap=True)
    ws.row_dimensions[HEADER_ROW].height = 32

    # メンバーブロック(予定/実績/差異)
    sample = {"山田 太郎": ([0, 40, 40, 36, 40, 8], [0, 44, 38, 40, 36, 6])}
    for b in range(MEMBER_BLOCKS):
        r0 = FIRST_ROW + b * 3
        ws.merge_cells(start_row=r0, start_column=1, end_row=r0 + 2, end_column=1)
        name_cell = ws.cell(r0, 1)
        _style(name_cell, color="0000FF", fill=FILL_INPUT, wrap=True)
        for k in (1, 2):  # 結合範囲の罫線用
            ws.cell(r0 + k, 1).border = BORDER
        plan_vals = act_vals = None
        if members is not None:
            if b < len(members):
                name_cell.value = members[b]
        elif b == 0:
            name_cell.value = "山田 太郎"
            plan_vals, act_vals = sample["山田 太郎"]
        for k, label in enumerate(["予定", "実績", "差異"]):
            r = r0 + k
            _style(ws.cell(r, 2, label), bold=True, fill=FILL_CALC if k == 2 else None)
            for j in range(n):
                c = ws.cell(r, col_first + j)
                cl = L(col_first + j)
                if k == 2:
                    c.value = f'=IF(AND({cl}{r0}="",{cl}{r0 + 1}=""),"",{cl}{r0 + 1}-{cl}{r0})'
                    _style(c, fill=FILL_CALC)
                else:
                    vals = plan_vals if k == 0 else act_vals
                    if vals and j < len(vals):
                        c.value = vals[j]
                    _style(c, color="0000FF", fill=FILL_INPUT)
                c.number_format = '0.0;-0.0;0.0'
            tc = ws.cell(r, col_total)
            rng = f"{L(col_first)}{r}:{L(col_total - 1)}{r}"
            tc.value = f'=IF(COUNT({rng})=0,"",SUM({rng}))'
            tc.number_format = '0.0;-0.0;0.0'
            _style(tc, bold=True, fill=FILL_CALC)

    # 全体合計
    last = FIRST_ROW + MEMBER_BLOCKS * 3 - 1
    for k, label in enumerate(["予定", "実績", "差異"]):
        r = last + 1 + k
        if k == 0:
            ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=1)
            _style(ws.cell(r, 1, "全体合計"), bold=True, fill=FILL_TOTAL)
        for kk in (1, 2):
            ws.cell(last + 1 + kk, 1).border = BORDER
        _style(ws.cell(r, 2, label), bold=True, fill=FILL_TOTAL)
        for j in range(n + 1):
            cl = L(col_first + j)
            c = ws.cell(r, col_first + j)
            if k < 2:
                c.value = f'=SUMIF($B${FIRST_ROW}:$B${last},$B{r},{cl}${FIRST_ROW}:{cl}${last})'
            else:
                c.value = f"={cl}{r - 1}-{cl}{r - 2}"
            c.number_format = '0.0;-0.0;0.0'
            _style(c, bold=True, fill=FILL_TOTAL)

    # 実績が予定を下回る場合は太赤字(条件付き書式)。全体合計のブロックも対象
    red_bold = Font(name=FONT_NAME, size=10, bold=True, color="FF0000")
    first, lastc = L(col_first), L(col_total)
    for r_plan in [FIRST_ROW + b * 3 for b in range(MEMBER_BLOCKS)] + [last + 1]:
        ra, rd = r_plan + 1, r_plan + 2
        ws.conditional_formatting.add(
            f"{first}{ra}:{lastc}{ra}",
            FormulaRule(formula=[f"AND(ISNUMBER({first}{ra}),ISNUMBER({first}{r_plan}),{first}{ra}<{first}{r_plan})"],
                        font=red_bold))
        ws.conditional_formatting.add(
            f"{first}{rd}:{lastc}{rd}",
            FormulaRule(formula=[f"AND(ISNUMBER({first}{rd}),{first}{rd}<0)"], font=red_bold))

    # 入力規則(0以上の数値)
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="入力エラー", error="0以上の数値(h)を入力してください")
    ws.add_data_validation(dv)
    for b in range(MEMBER_BLOCKS):
        r0 = FIRST_ROW + b * 3
        dv.add(f"{L(col_first)}{r0}:{L(col_total - 1)}{r0 + 1}")

    # 体裁
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 8
    for j in range(n + 1):
        ws.column_dimensions[L(col_first + j)].width = 12
    ws.freeze_panes = ws.cell(FIRST_ROW, col_first)
    ws.sheet_view.showGridLines = False
    return ws


def month_sheets(wb):
    """「YY年MM月」形式の月シートだけを返す。"""
    return [w for w in wb.worksheets if re.fullmatch(r"\d{2}年\d{2}月", w.title)]


def build_control_sheet(wb: Workbook, year: int, month: int, ws=None):
    """シート追加用の操作シートを作る(年月入力欄とボタン設置場所)。"""
    ws = ws or wb.create_sheet(index=0)
    ws.title = CONTROL_SHEET
    ws["A1"] = "月シートの追加"
    ws["A1"].font = Font(name=FONT_NAME, size=14, bold=True)
    ws["A2"] = "追加したい年月を入力し、下の「シートを追加」ボタンを押してください(マクロ有効ブックが必要)。"
    ws["A2"].font = Font(name=FONT_NAME, size=9, color="595959")
    for r, (label, val, fmt) in enumerate([("年(西暦)", year, "0"), ("月(1~12)", month, "0")], 4):
        _style(ws.cell(r, 2, label), bold=True, fill=FILL_CALC)
        c = ws.cell(r, 3, val)
        _style(c, color="0000FF", fill=FILL_INPUT)
        c.number_format = fmt
    ws["B8"] = "(ここにボタンが設置されます)"
    ws["B8"].font = Font(name=FONT_NAME, size=9, color="999999")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.sheet_view.showGridLines = False
    return ws


def read_members(ws):
    """シートのメンバー名(A列の各ブロック先頭)を読み取る。"""
    names = []
    for b in range(MEMBER_BLOCKS):
        v = ws.cell(FIRST_ROW + b * 3, 1).value
        if v:
            names.append(v)
    return names


def add_month_sheet(path: str, year: int, month: int) -> str:
    """既存ブックに月シートを追加して保存する。ファイルが無ければ新規作成する。
    直近の月シートがあればメンバー名を引き継ぐ。同名シートがある場合はValueError。
    """
    name = sheet_name(year, month)
    try:
        wb = load_workbook(path)
    except FileNotFoundError:
        wb = Workbook()
        build_control_sheet(wb, year, month, ws=wb.active)
        ensure_holiday_sheet(wb)
        build_sheet(wb, year, month)
        wb.save(path)
        return name
    if name in wb.sheetnames:
        raise ValueError(f"シート「{name}」は既に存在します")
    ensure_holiday_sheet(wb)
    existing = month_sheets(wb)
    members = read_members(max(existing, key=lambda w: w.title)) if existing else None
    build_sheet(wb, year, month, members=members)
    # 月シートを年月順に並べ替える(操作シートなどその他のシートは位置を保つ)
    slots = [i for i, w in enumerate(wb._sheets) if w in month_sheets(wb)]
    for i, w in zip(slots, sorted(month_sheets(wb), key=lambda w: w.title)):
        wb._sheets[i] = w
    if CONTROL_SHEET not in wb.sheetnames:
        build_control_sheet(wb, year, month)
    wb.save(path)
    return name


def main():
    today = date.today()
    year = int(sys.argv[1]) if len(sys.argv) > 1 else today.year
    month = int(sys.argv[2]) if len(sys.argv) > 2 else today.month
    out = sys.argv[3] if len(sys.argv) > 3 else "工数予実管理表.xlsx"
    try:
        name = add_month_sheet(out, year, month)
    except ValueError as e:
        sys.exit(str(e))
    print(f"シートを追加しました: {out} (シート: {name})")


if __name__ == "__main__":
    main()
