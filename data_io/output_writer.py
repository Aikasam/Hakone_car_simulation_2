"""配車結果の書き出し(Excel形式、1ファイル3シート)。

元コード(logic/allocator.py の save_plan_to_csv)との違いは2点:
  ① 計算に使った入力データ(参加者一覧)も出力する
  ② 「入力データ」「区間ごとのランナー一覧」「区間ごとの配車」を1つの表に
     混在させず、1つのExcelファイル(.xlsx)の3枚のシートに分けて出力する
     (CSVには複数シートという概念が無いため、xlsxにした)
"""

import colorsys
import os
import re
import sys
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Participant, SectionState
from logic.car_pool import ALL_CAR_IDS, section_label
from validator import compute_individual_summary

INPUT_HEADER = ["名前", "学年", "宿泊", "運転", "大", "山",
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
                "希望区間数", "離脱区間", "特に走りたい区間"]

CARS_FIXED_HEADER = ["区間", "車ID", "車種", "山行き", "先行", "運転手"]

SECTION_BORDER = Border(bottom=Side(style="medium"))
THIN_SIDE = Side(style="thin", color="FFB0B0B0")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

FILL_WHITE = PatternFill(fill_type="solid", fgColor="FFFFFFFF")
FILL_BLACK = PatternFill(fill_type="solid", fgColor="FF000000")
FILL_RED = PatternFill(fill_type="solid", fgColor="FFFF0000")

# 0/1をそのまま数値で見せる代わりに、セルの塗り色だけで表現する列(白=0/黒=1)
INPUT_BOOL_COLS = [
    ("宿泊", lambda p: p.leaves_after_section is None),
    ("運転", lambda p: p.can_drive),
    ("大", lambda p: p.can_drive_large),
    ("山", lambda p: p.can_drive_mountain),
]

INPUT_LEGEND = [
    (FILL_WHITE, "＝いいえ／該当なし"),
    (FILL_BLACK, "＝はい／希望している"),
    (FILL_RED, "＝希望区間のうち「特に走りたい区間」"),
]

# 個人別まとめシート用の役割別の塗り色(薄い色。セル内のテキストが読めるように)
FILL_RUNNER = PatternFill(fill_type="solid", fgColor="FFFFE699")
FILL_DRIVER = PatternFill(fill_type="solid", fgColor="FFBDD7EE")
FILL_PASSENGER = PatternFill(fill_type="solid", fgColor="FFC6E0B4")

HUE_DRIVER = 210 / 360     # 運転手=青系
HUE_PASSENGER = 120 / 360  # 同乗者=緑系

INDIVIDUAL_LEGEND = [
    (FILL_RUNNER, "＝ランナー"),
    (FILL_DRIVER, "＝運転手（車によって濃淡が変わります。下の「車ごとの色」参照）"),
    (FILL_PASSENGER, "＝同乗者（車によって濃淡が変わります。下の「車ごとの色」参照）"),
    (FILL_WHITE, "＝その区間は不参加(離脱済み等)"),
]

_CAR_ID_RE = re.compile(r"\(([^)]+)\)")


def _shade_fill(hue: float, idx: int, n: int) -> PatternFill:
    """同じ色相(hue)のまま、車ごとに明度を変えた塗り色を作る(車の見分け用)。"""
    light_min, light_max = 0.55, 0.88
    lightness = light_max if n <= 1 else light_max - (light_max - light_min) * idx / (n - 1)
    r, g, b = colorsys.hls_to_rgb(hue, lightness, 0.55)
    return PatternFill(fill_type="solid", fgColor="FF{:02X}{:02X}{:02X}".format(round(r * 255), round(g * 255), round(b * 255)))


def _build_car_fills(used_car_ids: List[str]) -> Tuple[Dict[str, PatternFill], Dict[str, PatternFill]]:
    n = len(used_car_ids)
    driver_fills = {car_id: _shade_fill(HUE_DRIVER, i, n) for i, car_id in enumerate(used_car_ids)}
    passenger_fills = {car_id: _shade_fill(HUE_PASSENGER, i, n) for i, car_id in enumerate(used_car_ids)}
    return driver_fills, passenger_fills


def _add_legend(ws, start_row: int, start_col: int, items: List[Tuple[PatternFill, str]], title: str = "凡例") -> None:
    """色付きセル+説明を1行ずつ並べた凡例を追加する。
    メインの表の列とは別の列(start_col)に置くことで、_autosize()が凡例の長い説明文に
    引っ張られてメインの表の列幅まで不自然に広がらないようにしている。"""
    ws.cell(row=start_row, column=start_col, value=title).font = Font(bold=True)
    for i, (fill, label) in enumerate(items, start=1):
        r = start_row + i
        swatch = ws.cell(row=r, column=start_col)
        swatch.fill = fill
        swatch.border = THIN_BORDER
        ws.cell(row=r, column=start_col + 1, value=label)


def _participant_row(p: Participant) -> List:
    return [
        p.name,
        p.grade,
        None, None, None, None,  # 宿泊/運転/大/山: 値は入れず、塗り色だけで表現する
        *[None for _ in p.preferred_sections],  # 1〜10: 同上(希望=黒 or 特に希望=赤、非希望=白)
        p.remaining_sections,
        None if p.leaves_after_section is None else p.leaves_after_section,
        ", ".join(f"{i+1}区" for i, v in enumerate(p.priority_sections) if v) or None,
    ]


def _runner_rows(plan: List[SectionState], participants: Dict[str, Participant]) -> List[List]:
    """区間別配車シートと同じ形式(1区間=1行、ランナーは横に並べる)にする。"""
    rows = []
    for section in plan:
        label = section_label(section.section_id)
        names = [participants[pid].name for pid in section.runner_ids if pid in participants]
        rows.append([label, *names])
    return rows


def _runners_header(rows: List[List]) -> List[str]:
    max_runners = max((len(row) - 1 for row in rows), default=0)
    return ["区間"] + [f"ランナー{i}" for i in range(1, max_runners + 1)]


def _car_blocks(plan: List[SectionState], participants: Dict[str, Participant]) -> List[Tuple[str, List[List]]]:
    """区間ごとにグループ化した配車行を返す(1区間=1ブロック、各行は先頭の区間ラベルを含まない)。"""
    blocks = []
    for section in plan:
        label = section_label(section.section_id)
        section_rows = []
        for car in section.cars:
            driver_name = participants[car.driver_id].name if car.driver_id in participants else "エラー"
            passengers = [participants[pid].name for pid in car.passenger_ids if pid in participants]
            car_type = "大型" if car.car_type == "large" else "普通"
            is_mt = "★山行き" if car.is_mountain_goer else ("🏨ホテル組" if car.group == "hotel" else "")
            is_adv = "🚀先行" if car.is_advance else ""
            section_rows.append([car.car_id, car_type, is_mt, is_adv, driver_name, *passengers])
        blocks.append((label, section_rows))
    return blocks


def _cars_header(blocks: List[Tuple[str, List[List]]]) -> List[str]:
    n_fixed = len(CARS_FIXED_HEADER) - 1  # 区間列を除いた固定列数(車ID/車種/山行き/先行/運転手)
    max_passengers = max(
        (len(row) - n_fixed for _, rows in blocks for row in rows),
        default=0,
    )
    return CARS_FIXED_HEADER + [f"同乗者{i}" for i in range(1, max_passengers + 1)]


def _driver_matrix(plan: List[SectionState], participants: Dict[str, Participant]) -> Tuple[List[str], List[List]]:
    """行=区間、列=車ID として運転手のみをまとめる表(どの区間で運転手が交代したか一目で分かる)。"""
    used_car_ids = {car.car_id for section in plan for car in section.cars}
    car_ids = [c for c in ALL_CAR_IDS if c in used_car_ids]

    header = ["区間"] + car_ids
    rows = []
    for section in plan:
        label = section_label(section.section_id)
        driver_by_car = {}
        for car in section.cars:
            driver_by_car[car.car_id] = participants[car.driver_id].name if car.driver_id in participants else "エラー"
        rows.append([label] + [driver_by_car.get(car_id) for car_id in car_ids])
    return header, rows


def _write_input_sheet(ws, participants: Dict[str, Participant]) -> None:
    ws.append(INPUT_HEADER)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER

    section_col_start = INPUT_HEADER.index("1") + 1  # 1-based列番号
    n_cols = len(INPUT_HEADER)

    for p in participants.values():
        ws.append(_participant_row(p))
        r = ws.max_row

        for col in range(1, n_cols + 1):
            ws.cell(row=r, column=col).border = THIN_BORDER

        for col_name, flag_fn in INPUT_BOOL_COLS:
            col_idx = INPUT_HEADER.index(col_name) + 1
            ws.cell(row=r, column=col_idx).fill = FILL_BLACK if flag_fn(p) else FILL_WHITE

        for i in range(10):
            col_idx = section_col_start + i
            if not p.preferred_sections[i]:
                fill = FILL_WHITE
            elif p.priority_sections[i]:
                fill = FILL_RED
            else:
                fill = FILL_BLACK
            ws.cell(row=r, column=col_idx).fill = fill

    _add_legend(ws, 1, n_cols + 2, INPUT_LEGEND)


def _individual_summary_rows(
    plan: List[SectionState], participants: Dict[str, Participant]
) -> Tuple[List[str], List[List]]:
    summary = compute_individual_summary(plan, participants)
    labels = [section_label(s) for s in list(range(1, 11)) + [11]]
    header = ["名前"] + labels
    rows = []
    for pid, p in participants.items():
        per_section = summary.get(pid, {})
        rows.append([p.name] + [per_section.get(label) for label in labels])
    return header, rows


def _role_fill(
    text: Optional[str],
    driver_fills: Dict[str, PatternFill],
    passenger_fills: Dict[str, PatternFill],
) -> PatternFill:
    if text is None:
        return FILL_WHITE
    if text.startswith("🏃"):
        return FILL_RUNNER
    m = _CAR_ID_RE.search(text)
    car_id = m.group(1) if m else None
    if text.startswith("🚘"):
        return driver_fills.get(car_id, FILL_DRIVER)
    if text.startswith("👥"):
        return passenger_fills.get(car_id, FILL_PASSENGER)
    return FILL_WHITE


def _write_individual_sheet(ws, plan: List[SectionState], participants: Dict[str, Participant]) -> None:
    header, rows = _individual_summary_rows(plan, participants)
    n_cols = len(header)

    used_car_ids = [c for c in ALL_CAR_IDS if c in {car.car_id for section in plan for car in section.cars}]
    driver_fills, passenger_fills = _build_car_fills(used_car_ids)

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER

    for row in rows:
        ws.append(row)
        r = ws.max_row
        ws.cell(row=r, column=1).border = THIN_BORDER

        # 区間を跨いで同じ役割・車が続く場合は、その範囲のセルを結合する。
        # 塗り・罫線は結合範囲の全セルに設定してから結合する(結合後だと先頭セル以外は
        # MergedCellになりopenpyxl再読み込み時にスタイルが見えなくなるが、書き込まれる
        # xlsx自体には各セルのスタイルが正しく保存されており、Excel側では正しく描画される。
        # 逆に先頭セルだけに設定すると、Excelで結合範囲の右側の罫線が欠けて見える)
        col = 2
        while col <= n_cols:
            text = row[col - 1]
            end_col = col
            while end_col + 1 <= n_cols and row[end_col] == text:
                end_col += 1

            fill = _role_fill(text, driver_fills, passenger_fills)
            for c in range(col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = THIN_BORDER
                cell.fill = fill
                if c != col:
                    cell.value = None

            if end_col > col:
                ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=end_col)
                ws.cell(row=r, column=col).alignment = Alignment(horizontal="center", vertical="center")

            col = end_col + 1

    _add_legend(ws, 1, n_cols + 2, INDIVIDUAL_LEGEND)
    if used_car_ids:
        car_legend_items = []
        for car_id in used_car_ids:
            car_legend_items.append((driver_fills[car_id], f"＝{car_id} 運転手"))
            car_legend_items.append((passenger_fills[car_id], f"＝{car_id} 同乗者"))
        _add_legend(ws, len(INDIVIDUAL_LEGEND) + 3, n_cols + 2, car_legend_items, title="車ごとの色")


def _autosize(ws) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 8), 40)


def write_plan_xlsx(plan: List[SectionState], participants: Dict[str, Participant], output_path: str) -> str:
    """入力データ・区間別ランナー・区間別配車・個人別まとめの4シートを持つ1つのxlsxファイルとして書き出す。"""
    wb = Workbook()

    ws_input = wb.active
    ws_input.title = "入力データ"
    _write_input_sheet(ws_input, participants)
    _autosize(ws_input)
    ws_input.freeze_panes = "B2"  # 横にスクロールしても「名前」列が見え続けるようにする

    ws_runners = wb.create_sheet("区間別ランナー")
    runner_rows = _runner_rows(plan, participants)
    ws_runners.append(_runners_header(runner_rows))
    for row in runner_rows:
        ws_runners.append(row)
    _autosize(ws_runners)
    ws_runners.freeze_panes = "B2"  # 横にスクロールしても「区間」列が見え続けるようにする

    ws_cars = wb.create_sheet("区間別配車")
    _write_cars_sheet(ws_cars, plan, participants)
    _autosize(ws_cars)
    ws_cars.freeze_panes = "C2"  # 横にスクロールしても「区間」「車ID」列が見え続けるようにする

    ws_individual = wb.create_sheet("個人別まとめ")
    _write_individual_sheet(ws_individual, plan, participants)
    _autosize(ws_individual)
    ws_individual.freeze_panes = "B2"  # 横にスクロールしても「名前」列が見え続けるようにする

    wb.save(output_path)
    print(f"\n✅ Excelファイル '{output_path}' を作成しました（入力データ/区間別ランナー/区間別配車/個人別まとめの4シート）。")
    return output_path


def _write_cars_sheet(ws, plan: List[SectionState], participants: Dict[str, Participant]) -> None:
    blocks = _car_blocks(plan, participants)
    header = _cars_header(blocks)
    n_cols = len(header)

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for label, section_rows in blocks:
        start_row = ws.max_row + 1
        if not section_rows:
            # ランナー無し等でその区間の配車が0台の場合も、区切りが分かるように1行だけ出す
            ws.append([label])
            end_row = ws.max_row
        else:
            for i, row in enumerate(section_rows):
                padded = row + [None] * (n_cols - 1 - len(row))
                ws.append([label if i == 0 else None, *padded])
            end_row = ws.max_row

        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        label_cell = ws.cell(row=start_row, column=1)
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(vertical="center")

        for col in range(1, n_cols + 1):
            ws.cell(row=end_row, column=col).border = SECTION_BORDER

    # 運転手交代早見表(行=区間、列=車ID)
    ws.append([])
    ws.append([])
    title_row = ws.max_row + 1
    ws.append(["■ 運転手一覧（車ID別・区間ごと）"])
    ws.cell(row=title_row, column=1).font = Font(bold=True)

    matrix_header, matrix_rows = _driver_matrix(plan, participants)
    header_row = ws.max_row + 1
    ws.append(matrix_header)
    for col in range(1, len(matrix_header) + 1):
        ws.cell(row=header_row, column=col).font = Font(bold=True)

    prev_by_car: Dict[str, str] = {}
    for row in matrix_rows:
        ws.append(row)
        r = ws.max_row
        for col_idx, car_id in enumerate(matrix_header[1:], start=2):
            driver = row[col_idx - 1]
            cell = ws.cell(row=r, column=col_idx)
            if driver is not None and prev_by_car.get(car_id) not in (None, driver):
                cell.font = Font(bold=True, color="C00000")  # 直前の区間から運転手が交代した箇所を強調
            if driver is not None:
                prev_by_car[car_id] = driver
