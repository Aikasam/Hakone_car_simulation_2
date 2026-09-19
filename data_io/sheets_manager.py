import re
import os
import sys
from typing import Dict, List

import gspread

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Participant, SectionState
from logic.car_pool import section_label

CREDENTIALS_DEFAULT = "credentials.json"


def _client(credentials_path: str) -> gspread.Client:
    # Streamlit Cloud では st.secrets から読む。CLI やローカルではファイルを使う。
    try:
        import json as _json
        import streamlit as st
        secrets = st.secrets  # SecretsManagerオブジェクトを取得
        if "credentials_json" in secrets:
            return gspread.service_account_from_dict(_json.loads(secrets["credentials_json"]))
        if "gcp_service_account" in secrets:
            return gspread.service_account_from_dict(dict(secrets["gcp_service_account"]))
    except Exception:
        pass  # CLI環境やSecrets未設定の場合はファイルにフォールバック

    return gspread.service_account(filename=credentials_path)


def _sheet_id_from_url(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError(f"URLからスプレッドシートIDを取得できませんでした: {url}")
    return match.group(1)


def _flag(val) -> bool:
    return str(val).strip().replace(".0", "") == "1"


def _int_or(val, default: int) -> int:
    s = str(val).strip().replace(".0", "")
    return int(s) if s else default


def _find_col(headers: List[str], keyword: str, exclude: str = "") -> str:
    """キーワードを含む列名を返す。見つからなければ空文字。excludeを含む列は除外する。"""
    return next((h for h in headers if keyword in h and (not exclude or exclude not in h)), "")


def load_participants_from_sheet(url: str, credentials_path: str = CREDENTIALS_DEFAULT) -> Dict[str, Participant]:
    gc = _client(credentials_path)
    sheet = gc.open_by_key(_sheet_id_from_url(url)).sheet1
    records = sheet.get_all_records()

    participants = {}
    for index, row in enumerate(records):
        p_id = f"p{index + 1}"

        preferred = [_flag(row.get(str(s), 0)) for s in range(1, 11)]

        is_large = _flag(row.get("大", 0))
        is_mountain = _flag(row.get("山", 0))
        is_drive = _flag(row.get("運転", 0)) or is_large or is_mountain

        leaves_str = str(row.get("離脱区間", "")).strip().replace(".0", "")
        leaves_after_section = int(leaves_str) if leaves_str else None

        participants[p_id] = Participant(
            id=p_id,
            name=str(row.get("名前", "")),
            preferred_sections=preferred,
            can_drive=is_drive,
            can_drive_large=is_large,
            can_drive_mountain=is_mountain,
            staying_overnight=_flag(row.get("宿泊", 0)),
            grade=_int_or(row.get("学年", 1), 1),
            remaining_sections=_int_or(row.get("希望区間数", 0), 0),
            leaves_after_section=leaves_after_section,
        )
    return participants


def load_participants_from_form_sheet(url: str, credentials_path: str = CREDENTIALS_DEFAULT) -> Dict[str, Participant]:
    """Googleフォームの回答スプレッドシートからParticipantを読み込む。"""
    gc = _client(credentials_path)
    all_values = gc.open_by_key(_sheet_id_from_url(url)).sheet1.get_all_values()

    if len(all_values) < 2:
        raise ValueError("フォームの回答が0件です。")

    headers = all_values[0]
    col_name     = _find_col(headers, "名前")       # "名前" は "お名前" にも "名前" にもマッチ
    col_grade    = _find_col(headers, "学年")
    col_sections = _find_col(headers, "走りたい区間", exclude="特に")
    col_count    = _find_col(headers, "何区間")
    col_drive    = _find_col(headers, "普通自動車")
    col_large    = _find_col(headers, "大型")
    col_mountain = _find_col(headers, "山道")
    col_stay     = _find_col(headers, "宿泊")
    col_leave    = _find_col(headers, "帰りますか")
    col_priority = _find_col(headers, "特に走りたい")

    participants = {}
    p_index = 0
    for row_vals in all_values[1:]:
        if not any(row_vals):
            continue
        row = dict(zip(headers, row_vals))

        def get(col: str) -> str:
            return str(row.get(col, "")).strip()

        def yes(col: str) -> bool:
            return get(col).startswith("はい")

        # 走りたい区間: カンマ・読点・中黒・スラッシュ・スペース（全角含む）など複数の区切り文字に対応
        def parse_sections(text: str) -> List[bool]:
            flags = [False] * 10
            for item in re.split(r'[,、・/\s　]+', text):
                item = item.strip().replace("区", "")
                if item.isdigit():
                    idx = int(item) - 1
                    if 0 <= idx < 10:
                        flags[idx] = True
            return flags

        preferred = parse_sections(get(col_sections))
        priority = parse_sections(get(col_priority)) if col_priority else [False] * 10

        grade_match = re.match(r"(\d+)", get(col_grade))
        grade = int(grade_match.group(1)) if grade_match else 1

        count_str = get(col_count)
        remaining = int(count_str) if count_str.isdigit() else preferred.count(True)

        leave_str = get(col_leave).replace("区", "")
        if leave_str.isdigit():
            leaves_after = int(leave_str)
        elif not yes(col_stay):
            # 日帰りだが「何区の後に帰るか」が未記入 -> 希望区間のうち最後の区間の後に帰るとみなす
            preferred_indices = [i for i, v in enumerate(preferred) if v]
            leaves_after = (preferred_indices[-1] + 1) if preferred_indices else None
        else:
            leaves_after = None

        is_large    = yes(col_large)
        is_mountain = yes(col_mountain)
        is_drive    = yes(col_drive) or is_large or is_mountain

        p_index += 1
        p_id = f"p{p_index}"
        participants[p_id] = Participant(
            id=p_id,
            name=get(col_name),
            preferred_sections=preferred,
            can_drive=is_drive,
            can_drive_large=is_large,
            can_drive_mountain=is_mountain,
            staying_overnight=yes(col_stay),
            grade=grade,
            remaining_sections=remaining,
            leaves_after_section=leaves_after,
            priority_sections=priority,
        )
        # 9・10区希望者のパース確認ログ
        if preferred[8] or preferred[9]:
            secs = [i + 1 for i, v in enumerate(preferred) if v]
            print(f"  ⛰️ 山行き候補: {get(col_name)} → 希望区間 {secs}  (生データ: '{get(col_sections)}')")

    return participants


RESULT_SHEET_NAME = "配車結果"
INPUT_SHEET_NAME = "入力データ"

INPUT_HEADER = ["名前", "学年", "宿泊", "運転", "大", "山",
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
                "希望区間数", "離脱区間", "特に走りたい区間"]


def _participant_row(p: Participant) -> List[str]:
    return [
        p.name,
        str(p.grade),
        "1" if p.leaves_after_section is None else "0",
        "1" if p.can_drive else "0",
        "1" if p.can_drive_large else "0",
        "1" if p.can_drive_mountain else "0",
        *["1" if v else "" for v in p.preferred_sections],
        str(p.remaining_sections),
        "" if p.leaves_after_section is None else str(p.leaves_after_section),
        ", ".join(f"{i+1}区" for i, v in enumerate(p.priority_sections) if v),
    ]


def _replace_worksheet(spreadsheet, title: str, rows: int, cols: int):
    existing = next((ws for ws in spreadsheet.worksheets() if ws.title == title), None)
    if existing:
        spreadsheet.del_worksheet(existing)
    return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


def save_plan_to_sheet(
    plan: List[SectionState],
    participants: Dict[str, Participant],
    url: str,
    credentials_path: str = CREDENTIALS_DEFAULT,
) -> str:
    """配車結果を「配車結果」シートに、計算に使った入力データを「入力データ」シートに書き出す。
    どちらも既存シートがあれば削除して作り直す(元コードと同じ挙動)。"""
    gc = _client(credentials_path)
    spreadsheet = gc.open_by_key(_sheet_id_from_url(url))

    # 入力データシート: この結果を計算した時点の参加者データをそのまま書き出す
    input_sheet = _replace_worksheet(spreadsheet, INPUT_SHEET_NAME, rows=len(participants) + 5, cols=len(INPUT_HEADER))
    input_rows = [INPUT_HEADER] + [_participant_row(p) for p in participants.values()]
    input_sheet.update(input_rows, "A1")

    # 配車結果シート
    sheet = _replace_worksheet(spreadsheet, RESULT_SHEET_NAME, rows=500, cols=7)
    rows = [["区間", "ランナー", "車ID", "車種", "山行き", "先行", "運転手", "同乗者"]]
    for section in plan:
        runners = ", ".join(participants[pid].name for pid in section.runner_ids if pid in participants)
        for car in section.cars:
            driver_name = participants[car.driver_id].name if car.driver_id in participants else "エラー"
            passengers = ", ".join(participants[pid].name for pid in car.passenger_ids if pid in participants)
            car_type = "大型" if car.car_type == "large" else "普通"
            is_mt = "★山行き" if car.is_mountain_goer else ("🏨ホテル組" if car.group == "hotel" else "")
            is_adv = "🚀先行" if car.is_advance else ""
            rows.append([section_label(section.section_id), runners, car.car_id, car_type, is_mt, is_adv, driver_name, passengers])

    sheet.update(rows, "A1")
    return spreadsheet.url
