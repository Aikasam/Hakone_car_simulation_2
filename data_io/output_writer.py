"""配車結果の書き出し(Excel形式、1ファイル3シート)。

元コード(logic/allocator.py の save_plan_to_csv)との違いは2点:
  ① 計算に使った入力データ(参加者一覧)も出力する
  ② 「入力データ」「区間ごとのランナー一覧」「区間ごとの配車」を1つの表に
     混在させず、1つのExcelファイル(.xlsx)の3枚のシートに分けて出力する
     (CSVには複数シートという概念が無いため、xlsxにした)
"""

import colorsys
import os
import sys
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import CarState, Participant, SectionState
from logic.car_pool import ALL_CAR_IDS, CAR_TYPE, section_label

INPUT_HEADER = ["名前", "学年", "宿泊", "運転", "大", "山",
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
                "希望区間数", "離脱区間", "特に走りたい区間"]

CARS_FIXED_HEADER = ["区間", "車ID", "車種", "山行き", "先行", "運転手"]

SECTION_BORDER = Border(bottom=Side(style="medium"))
THIN_SIDE = Side(style="thin", color="FFB0B0B0")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

# 区間別配車で「直前の区間でランナーだった(=走者回収された)人」のセルを囲む赤枠。
PICKUP_SIDE = Side(style="thin", color="FFFF0000")
PICKUP_BORDER = Border(left=PICKUP_SIDE, right=PICKUP_SIDE, top=PICKUP_SIDE, bottom=PICKUP_SIDE)
# 区間ブロックの最終行では下辺だけSECTION_BORDER(区切り線)を優先させる。
PICKUP_BORDER_AT_SECTION_END = Border(
    left=PICKUP_SIDE, right=PICKUP_SIDE, top=PICKUP_SIDE, bottom=Side(style="medium"))


def _solid_fill(hex_color: str) -> PatternFill:
    """fgColorとbgColorの両方を設定する。条件付き書式(dxf)の塗りつぶしはExcel上で
    bgColorの方が実際の描画に使われる癖があり、fgColorしか設定しないと通常のセルの
    塗りでは見えるのに条件付き書式では色が出ない、という現象になるため。"""
    return PatternFill(fill_type="solid", fgColor=hex_color, bgColor=hex_color)


FILL_WHITE = _solid_fill("FFFFFFFF")
FILL_BLACK = _solid_fill("FF000000")
FILL_RED = _solid_fill("FFFF0000")

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
FILL_RUNNER = _solid_fill("FFFFE699")
FILL_DRIVER = _solid_fill("FFBDD7EE")
FILL_PASSENGER = _solid_fill("FFC6E0B4")
FILL_LOCAL_JOIN = _solid_fill("FFD9B3FF")  # 現地集合(1区を走らず2区を走る人)専用の色

HUE_DRIVER = 210 / 360     # 運転手=青系
HUE_PASSENGER = 120 / 360  # 同乗者=緑系

INDIVIDUAL_LEGEND = [
    (FILL_RED, "＝重複あり(区間別配車と区間別ランナーの両方にこの人がいる。要確認)"),
    (FILL_RUNNER, "＝ランナー"),
    (FILL_DRIVER, "＝運転手（車によって濃淡が変わります。下の「車ごとの色」参照）"),
    (FILL_PASSENGER, "＝同乗者（車によって濃淡が変わります。下の「車ごとの色」参照）"),
    (FILL_LOCAL_JOIN, "＝現地集合(1区は走らず2区から参加、1区の車には乗らない)"),
    (FILL_WHITE, "＝その区間は不参加(離脱済み等)"),
]

# 個人別まとめは区間別ランナー・区間別配車を「数式」で参照する(手動編集が自動反映されるように)。

# 区間別ランナー: 実際の人数より少し余裕を持たせておく(手動で1人追加したいときに
# 列が足りず、隣接シートの別の列を間違って使ってしまう事故を防ぐため)。
RUNNER_BUFFER = 5
# 区間別配車: 同乗者は車の定員上8人までなので、実データがそれ未満でも8列は確保しておく。
PASSENGER_CAP = 8
# メモ欄の手前に空ける列数(データ列とメモ欄を視覚的に区別するため)。
GAP_COLS = 3


def _shade_fill(hue: float, idx: int, n: int) -> PatternFill:
    """同じ色相(hue)のまま、車ごとに明度を変えた塗り色を作る(車の見分け用)。"""
    light_min, light_max = 0.55, 0.88
    lightness = light_max if n <= 1 else light_max - (light_max - light_min) * idx / (n - 1)
    r, g, b = colorsys.hls_to_rgb(hue, lightness, 0.55)
    return _solid_fill("FF{:02X}{:02X}{:02X}".format(round(r * 255), round(g * 255), round(b * 255)))


def _build_car_fills(used_car_ids: List[str]) -> Tuple[Dict[str, PatternFill], Dict[str, PatternFill]]:
    n = len(used_car_ids)
    driver_fills = {car_id: _shade_fill(HUE_DRIVER, i, n) for i, car_id in enumerate(used_car_ids)}
    passenger_fills = {car_id: _shade_fill(HUE_PASSENGER, i, n) for i, car_id in enumerate(used_car_ids)}
    return driver_fills, passenger_fills


# 区間別配車で「運転手・同乗者が次に走る区間」を示すための、区間ごとの色分け(1〜10区)。
# UIの区間別配車タブ(app.py)とExcel出力の両方から参照し、表示を完全に一致させる。
SECTION_COLORS = {
    1: "#DFBAB1",
    2: "#EDCDCC",
    3: "#F8E6D0",
    4: "#FDF3D0",
    5: "#DCE9D5",
    6: "#D3DFE2",
    7: "#CCDAF6",
    8: "#D3E1F2",
    9: "#D8D3E7",
    10: "#E6D2DC",
}


def compute_car_table_overrides(plan: List[SectionState], participants: Dict[str, Participant], header: List[str]):
    """区間別配車の「先行」列の上書きテキストと、運転手・同乗者セルの背景色(#RRGGBB)、
    直前の区間でランナーだった人(走者回収された人)のセルを計算する。
    UIの区間別配車タブ(app.py)とExcel出力(_write_cars_sheet)の両方がこの関数を使うことで、
    2つの表示が食い違わないようにしている。

    戻り値はどれも、planを「区間ごとに、車が無ければ1行・あれば車の数だけ」フラット化した
    行の順序(_car_blocksが作る行の順序と同じ)に対応するリスト:
      advance_overrides: 各行の「先行」列に表示すべきテキスト(車が無い行は"")
      cell_fills: 各行の {列インデックス(headerに対応、0始まり): 背景色} の辞書
                  (運転手・同乗者の列のみ。次に走る区間が無い人のセルには含めない)
      pickup_cols: 各行の、直前の区間でランナーだったため回収された人がいる
                   列インデックス(headerに対応、0始まり)のset
    """
    n_cols = len(header)
    driver_col = header.index("運転手")
    advance_col = header.index("先行")
    passenger_cols = list(range(driver_col + 1, n_cols))

    row_section_ids: List[int] = []
    row_car_ids: List[Optional[str]] = []
    row_person_ids: List[List[Optional[str]]] = []
    for section in plan:
        if not section.cars:
            row_section_ids.append(section.section_id)
            row_car_ids.append(None)
            row_person_ids.append([None] * n_cols)
            continue
        for car in section.cars:
            row_section_ids.append(section.section_id)
            row_car_ids.append(car.car_id)
            pids: List[Optional[str]] = [None] * n_cols
            pids[driver_col] = car.driver_id if car.driver_id in participants else None
            for j, pid in enumerate(car.passenger_ids):
                col = driver_col + 1 + j
                if col < n_cols and pid in participants:
                    pids[col] = pid
            row_person_ids.append(pids)

    running_by_person: Dict[str, List[int]] = {}
    runner_ids_by_section: Dict[int, set] = {}
    for section in plan:
        rids = {pid for pid in section.runner_ids if pid in participants}
        runner_ids_by_section[section.section_id] = rids
        for pid in rids:
            running_by_person.setdefault(pid, []).append(section.section_id)
    for lst in running_by_person.values():
        lst.sort()

    car_appearances: Dict[str, list] = {}
    for sec_id, car_id, pids in zip(row_section_ids, row_car_ids, row_person_ids):
        if car_id is None:
            continue
        roster = frozenset(p for p in pids if p is not None)
        car_appearances.setdefault(car_id, []).append((sec_id, roster))

    change_sections: Dict[str, List[int]] = {}
    for car_id, appearances in car_appearances.items():
        appearances.sort(key=lambda t: t[0])
        change_sections[car_id] = [
            cur_sec for (_, prev_roster), (cur_sec, cur_roster) in zip(appearances, appearances[1:])
            if cur_roster != prev_roster
        ]

    advance_overrides: List[str] = []
    cell_fills: List[Dict[int, str]] = []
    pickup_cols: List[set] = []
    for i, car_id in enumerate(row_car_ids):
        sec_id = row_section_ids[i]
        pids = row_person_ids[i]
        fills: Dict[int, str] = {}
        pickups: set = set()
        if car_id is None:
            advance_overrides.append("")
            cell_fills.append(fills)
            pickup_cols.append(pickups)
            continue

        current_people = [p for p in pids if p is not None]
        prev_runner_ids = runner_ids_by_section.get(sec_id - 1, set())
        next_change = next((t for t in change_sections.get(car_id, []) if t > sec_id), None)
        is_pickup = any(pid in prev_runner_ids for pid in current_people)
        if is_pickup:
            advance_overrides.append(f"走者回収&{section_label(next_change)}" if next_change else "走者回収")
        elif next_change is not None:
            advance_overrides.append(section_label(next_change))
        else:
            advance_overrides.append("")

        for col in [driver_col] + passenger_cols:
            pid = pids[col] if col < len(pids) else None
            if pid is None:
                continue
            if pid in prev_runner_ids:
                pickups.add(col)
            future = [s for s in running_by_person.get(pid, []) if s > sec_id]
            if future:
                fills[col] = SECTION_COLORS.get(min(future), "#EEEEEE")
        cell_fills.append(fills)
        pickup_cols.append(pickups)

    return advance_overrides, cell_fills, pickup_cols


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
    """実際の最大人数に、手動で追加編集する余地(RUNNER_BUFFER)を足した列数を確保する。"""
    max_runners = max((len(row) - 1 for row in rows), default=0)
    total_slots = max_runners + RUNNER_BUFFER
    return ["区間"] + [f"ランナー{i}" for i in range(1, total_slots + 1)]


def _write_runners_sheet(ws, plan: List[SectionState], participants: Dict[str, Participant]) -> Tuple[int, int]:
    """区間別ランナーシートを書く。個人別まとめは各セル(ランナー1〜N)を直接参照するので、
    このシート自体には隠し列を持たせない(隠し列を持たせると、ブロック単位のコピー&
    ペーストで一部の行だけ隠し列の値が消えてしまい、その行が個人別まとめから見えなく
    なる不具合が実際に起きたため)。"""
    rows = _runner_rows(plan, participants)
    header = _runners_header(rows)
    n_cols = len(header)

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append(row)

    memo_col = n_cols + GAP_COLS + 1
    ws.cell(row=1, column=memo_col, value="メモ").font = Font(bold=True)
    return n_cols, memo_col


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
    """同乗者は車の定員上PASSENGER_CAP人までなので、実データがそれ未満でも
    PASSENGER_CAP列は確保しておく(万一実データがそれを超える場合は切り捨てない)。"""
    n_fixed = len(CARS_FIXED_HEADER) - 1  # 区間列を除いた固定列数(車ID/車種/山行き/先行/運転手)
    max_passengers = max(
        (len(row) - n_fixed for _, rows in blocks for row in rows),
        default=0,
    )
    total_passengers = max(max_passengers, PASSENGER_CAP)
    return CARS_FIXED_HEADER + [f"同乗者{i}" for i in range(1, total_passengers + 1)]


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


_RUNNER_ROW_FOR_SECTION = {s: s + 1 for s in range(1, 11)}
_RUNNER_ROW_FOR_SECTION[11] = 12  # 帰路


def _runner_presence_formula(name_norm_ref: str, runner_row: int, runner_col_letters: List[str]) -> str:
    """区間別ランナーの指定行(runner_row)のランナー1〜N各セルを直接見て、
    そのいずれかがname_norm_refと一致するかを判定する式(隠し列に頼らない)。"""
    if not runner_col_letters:
        return "FALSE"
    terms = "+".join(
        f'({_normalize_expr(f"区間別ランナー!${c}${runner_row}")}={name_norm_ref})'
        for c in runner_col_letters
    )
    return f"({terms})>0"


def _individual_helper_formulas(
    name_norm_ref: str,
    section_idx: int,
    section_start_row: int,
    section_end_row: int,
    driver_col_letter: str,
    passenger_col_letters: List[str],
    car_id_col_letter: str,
    runner_row: int,
    runner_col_letters: List[str],
) -> Tuple[str, str, str]:
    """個人別まとめの隠し列3つ(ランナーか/運転手なら車ID/同乗者なら車ID)の数式を組み立てる。
    区間別ランナー・区間別配車の「見える列」だけを直接見るので、それらを手動編集すれば
    開き直すたびに再計算される。以前は区間別配車側に区間キー・正規化済み運転手・
    正規化済み同乗者結合という専用の隠し列を持たせていたが、区間ブロック単位の
    コピー&ペーストで一部の行だけそれらの隠し列の値が消えてしまい、その行が
    個人別まとめから見えなくなる不具合が実際に起きたため、区間別配車・区間別ランナー
    側には何も追加で持たせない設計にした(区間の範囲は生成時点で確定しているので、
    区間別配車側の行番号の範囲(section_start_row〜section_end_row)を直接この式に
    埋め込む)。複数条件でのユニーク行検索は、CSE(配列数式)入力が要らないSUMPRODUCT
    ベースの「一致した行の相対位置を合計する」方式にしている(INDEX(array,0)を使う
    MATCH方式はExcel互換の実装によって挙動が不安定になることを確認したため採用しなかった)。
    名前の正規化はこの関数の外(呼び出し側の隠し列)で1回だけ計算し、ここではその結果
    (name_norm_ref)を使い回す。3つの判定+重複検知の組み立て先(可視セル)で毎回
    正規化式を埋め込むと数式が数千文字・ネスト60段超まで膨れ上がり、Excelが条件付き
    書式や数式を「読み取れない内容」として削除してしまう不具合が実際に起きたため。"""
    driver_range = f"区間別配車!${driver_col_letter}${section_start_row}:${driver_col_letter}${section_end_row}"
    car_id_range = f"区間別配車!${car_id_col_letter}${section_start_row}:${car_id_col_letter}${section_end_row}"
    row_offset = f"(ROW({driver_range})-ROW({driver_range.split(':')[0]})+1)"

    driver_pos = f'SUMPRODUCT(({_normalize_expr(driver_range)}={name_norm_ref})*{row_offset})'
    driver_formula = f'=IF({driver_pos}=0,"",INDEX({car_id_range},{driver_pos}))'

    if passenger_col_letters:
        passenger_terms = "+".join(
            f'({_normalize_expr(f"区間別配車!${c}${section_start_row}:${c}${section_end_row}")}={name_norm_ref})'
            for c in passenger_col_letters
        )
        passenger_pos = f'SUMPRODUCT(({passenger_terms})*{row_offset})'
        passenger_formula = f'=IF({passenger_pos}=0,"",INDEX({car_id_range},{passenger_pos}))'
    else:
        passenger_formula = '=""'

    runner_formula = "=" + _runner_presence_formula(name_norm_ref, runner_row, runner_col_letters)

    return runner_formula, driver_formula, passenger_formula


def _individual_visible_formula(
    row: int,
    section_idx: int,
    runner_flag_col_letter: str,
    driver_car_col_letter: str,
    passenger_car_col_letter: str,
    local_join_flag_col_letter: str,
) -> str:
    """個人別まとめの可視セルの数式。隠し列(ランナーか/運転手車ID/同乗者車ID)を参照するだけ
    なので短く、区間別配車・区間別ランナーの両方に同じ人がいる「重複」も検知できる。"""
    rf = f"{runner_flag_col_letter}{row}"
    dc = f"{driver_car_col_letter}{row}"
    pc = f"{passenger_car_col_letter}{row}"

    dup_count = f'(--({rf}))+(--({dc}<>""))+(--({pc}<>""))'
    warn_text = (
        '"⚠️重複: "&TRIM('
        f'IF({rf},"ランナー ","")&'
        f'IF({dc}<>"","運転手("&{dc}&") ","")&'
        f'IF({pc}<>"","同乗者("&{pc}&") ","")'
        ")"
    )

    if section_idx == 1:
        # 現地集合の判定(2区のランナーかどうか)も、他の判定と同じく専用の隠し列を
        # 1回だけ計算したものを参照する(可視セルに直接埋め込むと、この判定だけ
        # ランナー列を全て舐める長い式になり、ネストが深くなりすぎるため)。
        lj = f"{local_join_flag_col_letter}{row}"
        fallback = f'IF({lj},"現地集合","")'
    else:
        fallback = '""'

    return (
        f'=IF({dup_count}>=2,{warn_text},'
        f'IF({rf},"🏃 ランナー",'
        f'IF({dc}<>"","🚘 運転手("&{dc}&")",'
        f'IF({pc}<>"","👥 同乗者("&{pc}&")",'
        f"{fallback}))))"
    )


def _write_individual_sheet(
    ws, plan: List[SectionState], participants: Dict[str, Participant],
    cars_n_cols: int, runners_n_cols: int,
) -> None:
    """区間別ランナー・区間別配車を数式で参照する個人別まとめシート。
    それらのシートを手動編集すると、このシートも開き直すたびに自動で再計算される。
    (数式は先頭セル1つにしか設定できないため、区間を跨いだセル結合とは両立しない)
    名前の正規化と、区間ごとの「ランナーか/運転手なら車ID/同乗者なら車ID」の判定は
    非表示の隠し列に1回だけ計算させ、可視セルの数式はそれらを参照するだけの
    短い式にする(理由は_individual_helper_formulasのコメント参照)。"""
    sections = [(s, section_label(s)) for s in list(range(1, 11)) + [11]]
    header = ["名前"] + [lbl for _, lbl in sections]
    n_cols = len(header)

    bounds_by_label = _section_row_bounds(plan)
    driver_col_letter = get_column_letter(len(CARS_FIXED_HEADER))
    car_id_col_letter = get_column_letter(2)
    passenger_col_letters = [get_column_letter(c) for c in range(len(CARS_FIXED_HEADER) + 1, cars_n_cols + 1)]
    runner_col_letters = [get_column_letter(c) for c in range(2, runners_n_cols + 1)]

    name_norm_col = n_cols + 1
    name_norm_col_letter = get_column_letter(name_norm_col)
    local_join_flag_col = name_norm_col + 1
    local_join_flag_col_letter = get_column_letter(local_join_flag_col)
    helper_start_col = local_join_flag_col + 1
    n_helper_cols = len(sections) * 3
    helper_last_col = helper_start_col + n_helper_cols - 1
    runner2_row = _RUNNER_ROW_FOR_SECTION[2]

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER

    for p in participants.values():
        ws.append([p.name])
        r = ws.max_row
        ws.cell(row=r, column=1).border = THIN_BORDER

        ws.cell(row=r, column=name_norm_col, value="=" + _normalize_expr(f"$A{r}"))
        name_norm_ref = f"${name_norm_col_letter}{r}"
        # 現地集合(1区を走らず2区から参加)の判定も、他の判定と同じく専用の隠し列に
        # 1回だけ計算しておく(理由は_individual_visible_formulaのコメント参照)。
        ws.cell(row=r, column=local_join_flag_col,
                value="=" + _runner_presence_formula(name_norm_ref, runner2_row, runner_col_letters))

        for i, (s, label) in enumerate(sections):
            runner_flag_col = helper_start_col + i * 3
            driver_car_col = runner_flag_col + 1
            passenger_car_col = runner_flag_col + 2

            section_start_row, section_end_row = bounds_by_label[label]
            runner_row = _RUNNER_ROW_FOR_SECTION[s]
            runner_formula, driver_formula, passenger_formula = _individual_helper_formulas(
                name_norm_ref, s, section_start_row, section_end_row,
                driver_col_letter, passenger_col_letters, car_id_col_letter,
                runner_row, runner_col_letters,
            )
            ws.cell(row=r, column=runner_flag_col, value=runner_formula)
            ws.cell(row=r, column=driver_car_col, value=driver_formula)
            ws.cell(row=r, column=passenger_car_col, value=passenger_formula)

            formula = _individual_visible_formula(
                r, s,
                get_column_letter(runner_flag_col), get_column_letter(driver_car_col),
                get_column_letter(passenger_car_col), local_join_flag_col_letter,
            )
            cell = ws.cell(row=r, column=2 + i, value=formula)
            cell.border = THIN_BORDER

    for c in range(name_norm_col, helper_last_col + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # 数式の計算結果に応じて自動で色が変わるよう、固定の塗りではなく条件付き書式を使う
    last_row = ws.max_row
    data_range = f"B2:{get_column_letter(n_cols)}{last_row}"
    anchor = "B2"
    used_car_ids = [c for c in ALL_CAR_IDS if c in {car.car_id for section in plan for car in section.cars}]
    driver_fills, passenger_fills = _build_car_fills(used_car_ids)

    ws.conditional_formatting.add(
        data_range, FormulaRule(formula=[f'ISNUMBER(SEARCH("⚠️重複",{anchor}))'], fill=FILL_RED, stopIfTrue=True))
    ws.conditional_formatting.add(
        data_range, FormulaRule(formula=[f'{anchor}="現地集合"'], fill=FILL_LOCAL_JOIN, stopIfTrue=True))
    ws.conditional_formatting.add(
        data_range, FormulaRule(formula=[f'ISNUMBER(SEARCH("ランナー",{anchor}))'], fill=FILL_RUNNER, stopIfTrue=True))
    for car_id in used_car_ids:
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'ISNUMBER(SEARCH("運転手({car_id})",{anchor}))'],
                        fill=driver_fills[car_id], stopIfTrue=True))
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(formula=[f'ISNUMBER(SEARCH("同乗者({car_id})",{anchor}))'],
                        fill=passenger_fills[car_id], stopIfTrue=True))
    ws.conditional_formatting.add(
        data_range, FormulaRule(formula=[f'{anchor}=""'], fill=FILL_WHITE, stopIfTrue=True))

    legend_col = helper_last_col + 2
    _add_legend(ws, 1, legend_col, INDIVIDUAL_LEGEND)
    if used_car_ids:
        car_legend_items = []
        for car_id in used_car_ids:
            car_legend_items.append((driver_fills[car_id], f"＝{car_id} 運転手"))
            car_legend_items.append((passenger_fills[car_id], f"＝{car_id} 同乗者"))
        _add_legend(ws, len(INDIVIDUAL_LEGEND) + 3, legend_col, car_legend_items, title="車ごとの色")


def _autosize(ws) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 8), 40)


def write_plan_xlsx(
    plan: List[SectionState], participants: Dict[str, Participant], output_path: str,
    use_formulas: bool = True,
) -> str:
    """入力データ・区間別ランナー・区間別配車・個人別まとめの4シートを持つ1つのxlsxファイルとして書き出す。
    use_formulas=False にすると、個人別まとめを(数式ではなく)「修理」機能と同じその場計算の
    値で書き出す(区間別配車・区間別ランナーを手動編集しても壊れない代わり、自動追従しない)。"""
    wb = Workbook()

    ws_input = wb.active
    ws_input.title = "入力データ"
    _write_input_sheet(ws_input, participants)
    _autosize(ws_input)
    ws_input.freeze_panes = "B2"  # 横にスクロールしても「名前」列が見え続けるようにする

    ws_runners = wb.create_sheet("区間別ランナー")
    runners_n_cols, runners_memo_col = _write_runners_sheet(ws_runners, plan, participants)
    _autosize(ws_runners)
    ws_runners.column_dimensions[get_column_letter(runners_memo_col)].width = 30  # _autosizeで縮むため上書き
    ws_runners.freeze_panes = "B2"  # 横にスクロールしても「区間」列が見え続けるようにする

    ws_cars = wb.create_sheet("区間別配車")
    cars_n_cols, cars_memo_col = _write_cars_sheet(ws_cars, plan, participants)
    _autosize(ws_cars)
    ws_cars.column_dimensions[get_column_letter(cars_memo_col)].width = 30  # _autosizeで縮むため上書き
    ws_cars.freeze_panes = "C2"  # 横にスクロールしても「区間」「車ID」列が見え続けるようにする

    ws_individual = wb.create_sheet("個人別まとめ")
    if use_formulas:
        _write_individual_sheet(ws_individual, plan, participants, cars_n_cols, runners_n_cols)
    else:
        _write_individual_sheet_static(ws_individual, plan, participants)
    _autosize(ws_individual)
    ws_individual.freeze_panes = "B2"  # 横にスクロールしても「名前」列が見え続けるようにする

    wb.save(output_path)
    formula_note = "数式あり" if use_formulas else "数式なし"
    print(f"\n✅ Excelファイル '{output_path}' を作成しました（入力データ/区間別ランナー/区間別配車/個人別まとめの4シート、{formula_note}）。")
    return output_path


def compute_individual_summary_static(
    plan: List[SectionState], participants: Dict[str, Participant]
) -> Dict[str, Dict[str, str]]:
    """個人別まとめの内容を、数式ではなくPythonでその場計算する(「修理」機能専用)。
    区間別配車・区間別ランナーの両方に同じ人がいる「重複」の検知も、write_plan_xlsx側の
    数式版(_individual_visible_formula)と同じルールで行う: ランナー・運転手・同乗者の
    うち2つ以上に該当したら「⚠️重複: ...」にする。"""
    summary: Dict[str, Dict[str, str]] = {pid: {} for pid in participants}
    for section in plan:
        label = section_label(section.section_id)
        is_runner = {pid for pid in section.runner_ids if pid in participants}
        driver_car: Dict[str, str] = {}
        passenger_car: Dict[str, str] = {}
        for car in section.cars:
            if car.driver_id in participants and car.driver_id not in driver_car:
                driver_car[car.driver_id] = car.car_id
            for pid in car.passenger_ids:
                if pid in participants and pid not in passenger_car:
                    passenger_car[pid] = car.car_id

        for pid in is_runner | set(driver_car) | set(passenger_car):
            rf = pid in is_runner
            dc = driver_car.get(pid)
            pc = passenger_car.get(pid)
            if sum([rf, dc is not None, pc is not None]) >= 2:
                parts = []
                if rf:
                    parts.append("ランナー")
                if dc:
                    parts.append(f"運転手({dc})")
                if pc:
                    parts.append(f"同乗者({pc})")
                summary[pid][label] = "⚠️重複: " + " ".join(parts)
            elif rf:
                summary[pid][label] = "🏃 ランナー"
            elif dc:
                summary[pid][label] = f"🚘 運転手({dc})"
            elif pc:
                summary[pid][label] = f"👥 同乗者({pc})"

    label_1, label_2 = section_label(1), section_label(2)
    for per_section in summary.values():
        if label_1 not in per_section and per_section.get(label_2, "").startswith("🏃"):
            per_section[label_1] = "現地集合"
    return summary


def _write_individual_sheet_static(
    ws, plan: List[SectionState], participants: Dict[str, Participant],
    summary: Optional[Dict[str, Dict[str, str]]] = None,
) -> None:
    """個人別まとめを、数式を一切使わずにその場の内容をそのまま書き込む(「修理」機能専用)。
    write_plan_xlsx側の数式版(_write_individual_sheet)とは別実装で、区間別配車・
    区間別ランナーを手で編集しても自動更新はされない代わり、セルのドラッグ移動などの
    構造変更で数式が壊れる問題自体が起こらない。summaryを渡すとcompute_individual_summary_static
    の代わりにそれを使う(修理のUIで、一部の変更だけ元の値に戻して書き出したい場合用)。"""
    sections = [(s, section_label(s)) for s in list(range(1, 11)) + [11]]
    header = ["名前"] + [lbl for _, lbl in sections]
    n_cols = len(header)

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = THIN_BORDER

    if summary is None:
        summary = compute_individual_summary_static(plan, participants)
    used_car_ids = [c for c in ALL_CAR_IDS if c in {car.car_id for section in plan for car in section.cars}]
    driver_fills, passenger_fills = _build_car_fills(used_car_ids)

    def _cell_fill(text: str) -> PatternFill:
        if text.startswith("⚠️重複"):
            return FILL_RED
        if text == "現地集合":
            return FILL_LOCAL_JOIN
        if text.startswith("🏃"):
            return FILL_RUNNER
        if text.startswith("🚘"):
            for car_id, fill in driver_fills.items():
                if f"運転手({car_id})" in text:
                    return fill
        if text.startswith("👥"):
            for car_id, fill in passenger_fills.items():
                if f"同乗者({car_id})" in text:
                    return fill
        return FILL_WHITE

    for pid, p in participants.items():
        row_values = [p.name] + [summary.get(pid, {}).get(label, "") for _, label in sections]
        ws.append(row_values)
        r = ws.max_row
        ws.cell(row=r, column=1).border = THIN_BORDER
        for i in range(len(sections)):
            cell = ws.cell(row=r, column=2 + i)
            cell.border = THIN_BORDER
            cell.fill = _cell_fill(cell.value or "")

    legend_col = n_cols + GAP_COLS
    _add_legend(ws, 1, legend_col, INDIVIDUAL_LEGEND)
    if used_car_ids:
        car_legend_items = []
        for car_id in used_car_ids:
            car_legend_items.append((driver_fills[car_id], f"＝{car_id} 運転手"))
            car_legend_items.append((passenger_fills[car_id], f"＝{car_id} 同乗者"))
        _add_legend(ws, len(INDIVIDUAL_LEGEND) + 3, legend_col, car_legend_items, title="車ごとの色")


def _scan_cars_sheet(ws) -> Optional[dict]:
    """区間別配車シートの列構成・(区間ラベル,車ID)ごとの行番号・区間ブロックの
    終端行を読み取る(読み取り専用、書き込みは行わない)。describe_car_row_changesと
    _patch_cars_sheet_in_placeの両方から使う共通のスキャン処理。"""
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if "運転手" not in header or "先行" not in header or "車ID" not in header:
        return None  # 想定外の形式
    driver_col = header.index("運転手") + 1
    advance_col = header.index("先行") + 1
    car_id_col = header.index("車ID") + 1
    car_type_col = header.index("車種") + 1 if "車種" in header else None
    mountain_col = header.index("山行き") + 1 if "山行き" in header else None
    passenger_cols = [i + 1 for i, v in enumerate(header) if v and str(v).startswith("同乗者")]

    title_row = None
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith("■"):
            title_row = r
            break
    last_row = (title_row - 1) if title_row else ws.max_row

    row_by_key: Dict[Tuple[str, str], int] = {}
    block_end_row: Dict[str, int] = {}
    current_label = None
    for r in range(2, last_row + 1):
        label_cell = ws.cell(row=r, column=1).value
        if label_cell:
            current_label = label_cell
        if current_label is None:
            continue
        block_end_row[current_label] = r
        car_id = ws.cell(row=r, column=car_id_col).value
        if car_id:
            row_by_key[(current_label, str(car_id))] = r

    return {
        "header": header, "driver_col": driver_col, "advance_col": advance_col,
        "car_id_col": car_id_col, "car_type_col": car_type_col, "mountain_col": mountain_col,
        "passenger_cols": passenger_cols,
        "row_by_key": row_by_key, "block_end_row": block_end_row, "last_row": last_row,
    }


def describe_car_row_changes(
    ws, plan: List[SectionState], participants: Dict[str, Participant]
) -> Optional[List[dict]]:
    """区間別配車シートの各(区間,車ID)行について、修理/裏シュミの新形式出力が
    実際にどう書き換えることになるかを、書き込む前に一覧として返す(UIで
    「この修正を適用するか」を選ぶための元データ)。各要素:
      label, car_id, excel_row(元シートに無ければ新規追加行なのでNone),
      old_advance, new_advance, old_names(元シートの運転手・同乗者名一覧),
      new_names(planに基づく新しい一覧), changed(bool: 何か変わるか)。
    シートの見出し(運転手/先行/車ID)が想定と違う場合はNoneを返す(空リストとは
    区別する。空リストは「解析できて、差分が無かった」、Noneは「解析できず、
    比較・修理ができなかった」を意味する)。"""
    from data_io.plan_importer import _name_to_id_map, _resolve_name

    scan = _scan_cars_sheet(ws)
    if scan is None:
        return None
    driver_col, advance_col = scan["driver_col"], scan["advance_col"]
    passenger_cols, row_by_key = scan["passenger_cols"], scan["row_by_key"]
    name_to_id = _name_to_id_map(participants)

    blocks = _car_blocks(plan, participants)
    virtual_header = _cars_header(blocks)
    advance_overrides, _cell_fills, _pickup_cols = compute_car_table_overrides(plan, participants, virtual_header)

    results = []
    row_idx = 0
    for section in plan:
        label = section_label(section.section_id)
        if not section.cars:
            row_idx += 1
            continue
        for car in section.cars:
            excel_row = row_by_key.get((label, car.car_id))
            old_advance = ws.cell(row=excel_row, column=advance_col).value if excel_row else None
            new_advance = advance_overrides[row_idx] or None

            old_driver_name = str(ws.cell(row=excel_row, column=driver_col).value or "").strip() if excel_row else ""
            old_passenger_names = []
            if excel_row:
                for real_col in passenger_cols:
                    v = ws.cell(row=excel_row, column=real_col).value
                    if v:
                        old_passenger_names.append(str(v).strip())
            old_names = ([old_driver_name] if old_driver_name else []) + old_passenger_names

            new_driver_name = participants[car.driver_id].name if car.driver_id in participants else ""
            new_passenger_names = [participants[p].name for p in car.passenger_ids if p in participants]
            new_names = ([new_driver_name] if new_driver_name else []) + new_passenger_names

            # 表記ゆれ(全角/半角スペースなど)だけの違いは変更として扱わない。
            # _patch_cars_sheet_in_place(実際の書き込み)がpid単位で比較しているのと
            # 同じ基準にするため、名前も正規化してpidに変換してから比較する
            # (運転手・同乗者の役割が入れ替わっただけの場合も、運転手だけは個別に
            # 比較するので見逃さない)。
            old_driver_pid = _resolve_name(name_to_id, old_driver_name) if excel_row else None
            old_passenger_pids = {_resolve_name(name_to_id, n) for n in old_passenger_names}
            old_passenger_pids.discard(None)
            new_driver_pid = car.driver_id if car.driver_id in participants else None
            new_passenger_pids = {p for p in car.passenger_ids if p in participants}
            occupants_changed = (
                excel_row is None
                or old_driver_pid != new_driver_pid
                or old_passenger_pids != new_passenger_pids
            )
            advance_changed = (old_advance or None) != (new_advance or None)
            results.append({
                "label": label, "car_id": car.car_id, "excel_row": excel_row,
                "old_advance": old_advance, "new_advance": new_advance,
                "old_names": old_names, "new_names": new_names,
                "changed": occupants_changed or advance_changed,
            })
            row_idx += 1
    return results


def _insert_new_car_row(ws, scan: dict, label: str, car: CarState) -> int:
    """区間別配車シートに、元ファイルには無かった新しい車の行を、その区間ブロックの
    末尾に1行追加する(裏シュミの新形式出力で、元ファイルに無い車IDが新たに使われた
    場合に発生しうる)。scanの行番号キャッシュ(row_by_key/block_end_row/last_row)も
    挿入に合わせてその場で更新する。"""
    car_id = car.car_id
    insert_at = scan["block_end_row"].get(label, scan["last_row"]) + 1

    # openpyxlのinsert_rowsは既存の結合セル範囲の再調整が不完全(挿入位置より下の
    # 結合セルがずれずに残ってしまい、ラベル列が壊れる)ため、先に全て解除してから
    # 挿入し、座標を補正して結合し直す。
    merged_ranges = [str(rng) for rng in ws.merged_cells.ranges]
    for rng in merged_ranges:
        ws.unmerge_cells(rng)

    ws.insert_rows(insert_at)
    scan["last_row"] += 1

    for rng in merged_ranges:
        min_col, min_row, max_col, max_row = range_boundaries(rng)
        if min_row >= insert_at:
            min_row += 1
        if max_row >= insert_at:
            max_row += 1
        ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
    for lbl, r in list(scan["block_end_row"].items()):
        if r >= insert_at:
            scan["block_end_row"][lbl] = r + 1
    for k, r in list(scan["row_by_key"].items()):
        if r >= insert_at:
            scan["row_by_key"][k] = r + 1

    ws.cell(row=insert_at, column=scan["car_id_col"], value=car_id)
    if scan["car_type_col"]:
        ws.cell(row=insert_at, column=scan["car_type_col"],
                value=("大型" if CAR_TYPE.get(car_id) == "large" else "普通"))
    if scan["mountain_col"]:
        mt_text = "★山行き" if car.is_mountain_goer else ("🏨ホテル組" if car.group == "hotel" else None)
        ws.cell(row=insert_at, column=scan["mountain_col"], value=mt_text)

    scan["block_end_row"][label] = insert_at
    scan["row_by_key"][(label, car_id)] = insert_at
    return insert_at


def _patch_cars_sheet_in_place(
    ws, plan: List[SectionState], participants: Dict[str, Participant],
    revert_keys: Optional[set] = None,
) -> None:
    """区間別配車シートの「セルの色付け」「先行列の内容」、そして(裏シュミの新形式出力の
    ように人員の交代がある場合は)運転手・同乗者の名前を書き換える。山行き・メモ・
    その見出し文字・運転手一覧・凡例など、それ以外の内容・書式は一切変更しない
    (元ファイルのまま)。表シュミの出力を丸ごと作り直すwrite_plan_xlsx/_write_cars_sheet
    とは別の実装。

    revert_keys((区間ラベル,車ID)のset)に含まれる行は、実際の変更内容に関わらず
    一切書き換えない(UIで「この修正は適用しない」と選ばれた行を元のまま保つため)。

    元シートの同乗者列に手動編集による空白(同乗者2が空欄で同乗者3に名前がある等)が
    あっても正しく塗り分けられるよう、どのpid(参加者)がどの実際の列にいるかは
    computeの仮想的な列位置(手前が空欄なら詰まる)ではなく、その行の実際のセルの
    名前を突き合わせて特定する。"""
    from data_io.plan_importer import _name_to_id_map, _resolve_name

    scan = _scan_cars_sheet(ws)
    if scan is None:
        return  # 想定外の形式なら何も変更しない(安全側)
    driver_col, advance_col = scan["driver_col"], scan["advance_col"]
    passenger_cols = scan["passenger_cols"]
    name_to_id = _name_to_id_map(participants)
    revert_keys = revert_keys or set()

    blocks = _car_blocks(plan, participants)
    virtual_header = _cars_header(blocks)
    v_driver_col = virtual_header.index("運転手")
    advance_overrides, cell_fills, pickup_cols = compute_car_table_overrides(plan, participants, virtual_header)

    row_idx = 0
    for section in plan:
        label = section_label(section.section_id)
        if not section.cars:
            row_idx += 1
            continue
        for car in section.cars:
            if (label, car.car_id) in revert_keys:
                row_idx += 1
                continue

            excel_row = scan["row_by_key"].get((label, car.car_id))
            if excel_row is None:
                excel_row = _insert_new_car_row(ws, scan, label, car)

            v_pids: List[Optional[str]] = [None] * len(virtual_header)
            v_pids[v_driver_col] = car.driver_id if car.driver_id in participants else None
            for j, pid in enumerate(car.passenger_ids):
                col = v_driver_col + 1 + j
                if col < len(v_pids) and pid in participants:
                    v_pids[col] = pid

            current_driver_pid = _resolve_name(name_to_id, ws.cell(row=excel_row, column=driver_col).value)
            current_passenger_pids = set()
            for real_col in passenger_cols:
                pid = _resolve_name(name_to_id, ws.cell(row=excel_row, column=real_col).value)
                if pid:
                    current_passenger_pids.add(pid)
            new_driver_pid = car.driver_id if car.driver_id in participants else None
            new_passenger_pids = {p for p in car.passenger_ids if p in participants}

            if current_driver_pid != new_driver_pid or current_passenger_pids != new_passenger_pids:
                # 運転手・同乗者どちらかの人員が変わった(役割の入れ替えも含む)行だけ、
                # 名前をplanの内容で書き直す(詰まった順で書く。変わっていない行は
                # 元の表記・空欄配置をそのまま保つ)。
                driver_name = participants[car.driver_id].name if car.driver_id in participants else None
                ws.cell(row=excel_row, column=driver_col, value=driver_name)
                passenger_names = [participants[p].name for p in car.passenger_ids if p in participants]
                for i, real_col in enumerate(passenger_cols):
                    ws.cell(row=excel_row, column=real_col,
                            value=passenger_names[i] if i < len(passenger_names) else None)

            ws.cell(row=excel_row, column=advance_col, value=advance_overrides[row_idx] or None)

            pid_fill = {v_pids[c]: color for c, color in cell_fills[row_idx].items() if v_pids[c]}
            pickup_pids = {v_pids[c] for c in pickup_cols[row_idx] if v_pids[c]}
            for real_col in [driver_col] + passenger_cols:
                cell = ws.cell(row=excel_row, column=real_col)
                pid = _resolve_name(name_to_id, cell.value)
                cell.fill = _solid_fill("FF" + pid_fill[pid].lstrip("#")) if pid in pid_fill \
                    else PatternFill(fill_type=None)
                existing_border = cell.border
                had_pickup_border = existing_border.left == PICKUP_SIDE
                if pid in pickup_pids:
                    cell.border = Border(
                        left=PICKUP_SIDE, right=PICKUP_SIDE, top=PICKUP_SIDE, bottom=existing_border.bottom)
                elif had_pickup_border:
                    # 前回の修理/裏シュミでは走者回収だったが今回は違う場合、赤枠だけ
                    # 外す(下辺の区切り線など、赤枠以外の既存の罫線は変更しない)。
                    cell.border = Border(bottom=existing_border.bottom)
            row_idx += 1


def _patch_runners_sheet_in_place(ws, plan: List[SectionState], participants: Dict[str, Participant]) -> None:
    """区間別ランナーシートのランナー名だけを、planの内容に合わせて書き直す。メモ列・
    その見出し文字・列幅など、それ以外は一切変更しない(元ファイルのまま)。同じ区間の
    ランナーの集合が変わっていなければその行は一切触らない(元シートに手動編集による
    空白セルが混じっていても、詰め直して書式・空欄配置を変えてしまわないように)。
    修理ではランナーの交代が無いので実質的に無変化、裏シュミの新形式出力では実際に
    交代した区間だけ内容が変わる(詰まった順で書き直す)。"""
    from data_io.plan_importer import _name_to_id_map, _resolve_name

    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if "区間" not in header:
        return
    runner_cols = [i + 1 for i, v in enumerate(header) if v and str(v).startswith("ランナー")]
    if not runner_cols:
        return
    name_to_id = _name_to_id_map(participants)

    sections_by_id = {s.section_id: s for s in plan}
    for s in list(range(1, 11)) + [11]:
        row = _RUNNER_ROW_FOR_SECTION[s]
        section = sections_by_id.get(s)
        new_pids = {pid for pid in (section.runner_ids if section else []) if pid in participants}

        current_pids = set()
        for col in runner_cols:
            pid = _resolve_name(name_to_id, ws.cell(row=row, column=col).value)
            if pid:
                current_pids.add(pid)

        if current_pids == new_pids:
            continue  # 誰も交代していないので元の表記・空欄配置のまま何も変えない

        # 書き直す順序はsection.runner_ids(plan)の元の順序に合わせる。
        ordered_pids = [pid for pid in (section.runner_ids if section else []) if pid in new_pids]
        names = [participants[pid].name for pid in ordered_pids]
        for i, col in enumerate(runner_cols):
            ws.cell(row=row, column=col, value=names[i] if i < len(names) else None)


def _apply_plan_to_workbook_no_formulas(
    wb, plan: List[SectionState], participants: Dict[str, Participant],
    revert_car_keys: Optional[set] = None,
    individual_summary: Optional[Dict[str, Dict[str, str]]] = None,
) -> None:
    """既に開いているworkbook(4シート形式)に対し、区間別ランナー・区間別配車の内容を
    planに合わせて更新し(実際に変わった箇所だけ書き換え、それ以外は元のまま保持する。
    修理のようにplanがsourceからそのまま読み込んだものなら実質無変化)、個人別まとめは
    数式を使わず全面的に書き直す。write_repaired_xlsx(修理)とwrite_updated_output_xlsx
    (裏シュミの新形式出力)の共通処理。"""
    if "区間別ランナー" in wb.sheetnames:
        _patch_runners_sheet_in_place(wb["区間別ランナー"], plan, participants)
    if "区間別配車" in wb.sheetnames:
        _patch_cars_sheet_in_place(wb["区間別配車"], plan, participants, revert_keys=revert_car_keys)

    if "個人別まとめ" in wb.sheetnames:
        del wb["個人別まとめ"]
    ws_individual = wb.create_sheet("個人別まとめ")
    _write_individual_sheet_static(ws_individual, plan, participants, summary=individual_summary)
    _autosize(ws_individual)
    ws_individual.freeze_panes = "B2"


def write_repaired_xlsx(
    source, plan: List[SectionState], participants: Dict[str, Participant], output_path: str,
    revert_car_keys: Optional[set] = None,
    individual_summary: Optional[Dict[str, Dict[str, str]]] = None,
) -> str:
    """「修理」機能専用の書き出し。write_plan_xlsx(表シュミ・裏シュミの出力)とは別物で、
    人員の交代や入力データとの不整合の調整は一切行わない。source(元ファイルのパスまたは
    file-likeオブジェクト)をそのまま土台として読み込み、修理の対象である
    (1)区間別配車のセルの色付け・先行列の内容 と (2)個人別まとめの内容 だけを書き換え、
    入力データ・区間別ランナー・区間別配車のそれ以外の部分(山行き・メモ・運転手一覧・
    凡例・列幅など)は一切変更せず元ファイルのまま保持する。個人別まとめだけは数式を
    使わずその場計算した値で全面的に書き直す(理由は_write_individual_sheet_staticの
    コメント参照)。

    revert_car_keys/individual_summaryを渡すと、UIで見せた修正候補のうち一部だけを
    「適用しない」と選んだ場合に、その箇所を元の値のまま出力できる
    (revert_car_keysは区間別配車の(区間ラベル,車ID)のset、individual_summaryは
    個人別まとめのpid→区間ラベル→表示文字列で、一部のセルだけ元の値に戻したもの)。"""
    wb = load_workbook(source)
    _apply_plan_to_workbook_no_formulas(
        wb, plan, participants, revert_car_keys=revert_car_keys, individual_summary=individual_summary)
    wb.save(output_path)
    print(f"\n✅ 修理済みExcelファイル '{output_path}' を作成しました（数式なし・修理対象外の内容は元ファイルのまま）。")
    return output_path


def write_updated_output_xlsx(source, plan: List[SectionState], participants: Dict[str, Participant], output_path: str) -> str:
    """裏シュミの結果を、修理と同じ形式(数式なし)で出力する。source(裏シュミの入力元
    として使った既存の4シート形式ファイル)を土台として読み込み、区間別ランナー・
    区間別配車のうち実際に裏シュミで変わった箇所だけを書き換える。入力データ・
    メモ・運転手一覧・凡例など、それ以外は元ファイルのまま保持する。個人別まとめは
    数式を使わず全面的に書き直す。"""
    wb = load_workbook(source)
    _apply_plan_to_workbook_no_formulas(wb, plan, participants)
    wb.save(output_path)
    print(f"\n✅ 裏シュミの結果Excelファイル(新形式) '{output_path}' を作成しました（数式なし・変更のない箇所は元ファイルのまま）。")
    return output_path


def _normalize_expr(ref: str) -> str:
    """名前セルの表記ゆれを吸収する正規化の数式断片を作る。
    「山田太郎（4区も走る）」のような括弧以降の注記を切り捨て、「山田　太郎」のような
    半角/全角スペースの混入も取り除く。COUNTIF等ではセル単位の変換ができないため、
    値そのものを正規化した状態で使う。
    括弧の位置の判定にMIN()ではなくIF()を使っているのは、MIN()は複数セルの範囲を
    渡すと範囲全体の最小値1つに潰れてしまい(セルごとの計算にならない)、refとして
    単一セルだけでなく範囲(例: $B$2:$B$100)を渡したときにも正しく1セルずつ判定
    できるようにするため(比較演算子・IF()は範囲を渡すと要素ごとに配列として
    計算されるが、MIN()等の集計関数は配列を丸ごと集約してしまう、というExcelの
    仕様の違いによる)。"""
    find_paren = f'IFERROR(FIND("(",{ref}),9999)'
    find_zenkaku_paren = f'IFERROR(FIND("（",{ref}),9999)'
    paren_pos = f'IF({find_paren}<{find_zenkaku_paren},{find_paren},{find_zenkaku_paren})'
    base = f'IF({paren_pos}=9999,{ref},LEFT({ref},{paren_pos}-1))'
    return f'SUBSTITUTE(SUBSTITUTE({base}," ",""),"　","")'


def _section_row_bounds(plan: List[SectionState]) -> Dict[str, Tuple[int, int]]:
    """区間別配車で各区間ブロックが実際に占める(開始行, 終了行)を区間ラベルごとに返す。
    _write_cars_sheetが書き込む行配置(1区間=section.cars件、0台でも1行)と必ず一致させる。"""
    bounds = {}
    row_cursor = 2
    for section in plan:
        label = section_label(section.section_id)
        n_rows = len(section.cars) if section.cars else 1
        bounds[label] = (row_cursor, row_cursor + n_rows - 1)
        row_cursor += n_rows
    return bounds


def _write_cars_sheet(ws, plan: List[SectionState], participants: Dict[str, Participant]) -> Tuple[int, int]:
    """区間別配車シートを書く。個人別まとめは運転手・同乗者の各列を直接参照するので、
    このシート自体には隠し列を持たせない(隠し列を持たせると、ブロック単位のコピー&
    ペーストで一部の行だけ隠し列の値が消えてしまい、その行が個人別まとめから見えなく
    なる不具合が実際に起きたため)。"""
    blocks = _car_blocks(plan, participants)
    header = _cars_header(blocks)
    n_cols = len(header)
    advance_col = header.index("先行")

    advance_overrides, cell_fills, pickup_cols = compute_car_table_overrides(plan, participants, header)
    row_idx = 0

    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for label, section_rows in blocks:
        start_row = ws.max_row + 1
        last_row_pickups: set = set()
        if not section_rows:
            # ランナー無し等でその区間の配車が0台の場合も、区切りが分かるように1行だけ出す
            ws.append([label])
            end_row = ws.max_row
            row_idx += 1
        else:
            for i, row in enumerate(section_rows):
                padded = list(row) + [None] * (n_cols - 1 - len(row))
                override = advance_overrides[row_idx]
                if override:
                    padded[advance_col - 1] = override
                ws.append([label if i == 0 else None, *padded])

                excel_row = ws.max_row
                for col_idx, color in cell_fills[row_idx].items():
                    ws.cell(row=excel_row, column=col_idx + 1).fill = _solid_fill("FF" + color.lstrip("#"))
                last_row_pickups = pickup_cols[row_idx]
                for col_idx in last_row_pickups:
                    ws.cell(row=excel_row, column=col_idx + 1).border = PICKUP_BORDER
                row_idx += 1
            end_row = ws.max_row

        if end_row > start_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        label_cell = ws.cell(row=start_row, column=1)
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(vertical="center")

        for col in range(1, n_cols + 1):
            if (col - 1) in last_row_pickups:
                ws.cell(row=end_row, column=col).border = PICKUP_BORDER_AT_SECTION_END
            else:
                ws.cell(row=end_row, column=col).border = SECTION_BORDER

    memo_col = n_cols + GAP_COLS + 1
    ws.cell(row=1, column=memo_col, value="メモ").font = Font(bold=True)

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

    # セルの色（次に走る区間）・「先行」列の意味の凡例(UIの区間別配車タブと同じ説明)
    ws.append([])
    ws.append([])
    legend_title_row = ws.max_row + 1
    ws.append(["■ セルの色・「先行」列の意味"])
    ws.cell(row=legend_title_row, column=1).font = Font(bold=True)
    for s in range(1, 11):
        r = ws.max_row + 1
        swatch = ws.cell(row=r, column=1)
        swatch.fill = _solid_fill("FF" + SECTION_COLORS[s].lstrip("#"))
        swatch.border = THIN_BORDER
        ws.cell(row=r, column=2, value=f"＝運転手・同乗者が次に走る区間が{s}区")
    note_row = ws.max_row + 1
    ws.cell(
        row=note_row, column=1,
        value="「先行」列＝その車の乗車メンバーが次に変わる区間。直前区間のランナーを"
              "新たに乗せた行は「走者回収&◯区」。",
    )

    return n_cols, memo_col
