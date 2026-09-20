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
from typing import Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Participant, SectionState
from logic.car_pool import ALL_CAR_IDS, section_label

INPUT_HEADER = ["名前", "学年", "宿泊", "運転", "大", "山",
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
                "希望区間数", "離脱区間", "特に走りたい区間"]

CARS_FIXED_HEADER = ["区間", "車ID", "車種", "山行き", "先行", "運転手"]

SECTION_BORDER = Border(bottom=Side(style="medium"))
THIN_SIDE = Side(style="thin", color="FFB0B0B0")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


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


def write_plan_xlsx(plan: List[SectionState], participants: Dict[str, Participant], output_path: str) -> str:
    """入力データ・区間別ランナー・区間別配車・個人別まとめの4シートを持つ1つのxlsxファイルとして書き出す。"""
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
    _write_individual_sheet(ws_individual, plan, participants, cars_n_cols, runners_n_cols)
    _autosize(ws_individual)
    ws_individual.freeze_panes = "B2"  # 横にスクロールしても「名前」列が見え続けるようにする

    wb.save(output_path)
    print(f"\n✅ Excelファイル '{output_path}' を作成しました（入力データ/区間別ランナー/区間別配車/個人別まとめの4シート）。")
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

    return n_cols, memo_col
