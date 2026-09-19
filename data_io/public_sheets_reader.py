"""Google Cloud のサービスアカウント設定なしにスプレッドシートを読む方法。

スプレッドシートが「リンクを知っている全員が閲覧可」に共有されていれば、
gvizのCSVエクスポートエンドポイントを認証なしのHTTP GETだけで取得できる。
gspread(サービスアカウント認証)を使う data_io/sheets_manager.py の代わりに、
読み込み専用でこちらを使えば credentials.json は一切不要になる。

書き込み(結果をスプシに反映)はGoogleが必ず何らかの認証を要求するため、
この方式では行わない。結果はアプリの画面表示とCSVダウンロードで受け取る。
"""

import csv
import io
import os
import re
import sys
from typing import Dict, List, Optional

import requests

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Participant


def _extract_sheet_id_and_gid(url: str) -> (str, Optional[str]):
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        raise ValueError(f"URLからスプレッドシートIDを取得できませんでした: {url}")
    sheet_id = match.group(1)

    gid_match = re.search(r"[?#&]gid=(\d+)", url)
    gid = gid_match.group(1) if gid_match else None
    return sheet_id, gid


def _fetch_csv_rows(url: str) -> List[List[str]]:
    sheet_id, gid = _extract_sheet_id_and_gid(url)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv"
    if gid:
        export_url += f"&gid={gid}"

    resp = requests.get(export_url, timeout=20)
    if resp.status_code != 200 or "text/csv" not in resp.headers.get("content-type", ""):
        raise ValueError(
            "スプレッドシートを認証なしで読み込めませんでした。"
            "共有設定が「リンクを知っている全員が閲覧可」になっているか確認してください。"
            f"(status={resp.status_code})"
        )
    text = resp.content.decode("utf-8-sig")
    return list(csv.reader(io.StringIO(text)))


def _flag(val) -> bool:
    return str(val).strip().replace(".0", "") == "1"


def _int_or(val, default: int) -> int:
    s = str(val).strip().replace(".0", "")
    return int(s) if s else default


def _find_col(headers: List[str], keyword: str, exclude: str = "") -> str:
    return next((h for h in headers if keyword in h and (not exclude or exclude not in h)), "")


def load_participants_from_public_sheet(url: str) -> Dict[str, Participant]:
    """独自フォーマットのスプレッドシートを認証なしで読み込む。

    注意: このフォーマットのシートはgvizのCSVエクスポートで見出し行の一部
    (学年・宿泊・運転・大・山・希望区間数など)が空文字になって返ってくることが
    ある(セル書式や名前付き範囲の影響と見られる)。見出し名に頼った列特定は
    信頼できないため、所定のフォーマット
    (名前・学年・宿泊・運転・大・山・1〜10・希望区間数・離脱区間) の
    列の並び順を固定で決め打ちして読む。
    """
    rows = _fetch_csv_rows(url)
    if len(rows) < 2:
        raise ValueError("データが0件です。")

    def cell(row_vals: List[str], idx: int) -> str:
        return row_vals[idx] if idx < len(row_vals) else ""

    participants = {}
    for index, row_vals in enumerate(rows[1:]):
        if not any(row_vals):
            continue
        p_id = f"p{index + 1}"

        preferred = [_flag(cell(row_vals, 6 + s)) for s in range(10)]
        is_large = _flag(cell(row_vals, 4))
        is_mountain = _flag(cell(row_vals, 5))
        is_drive = _flag(cell(row_vals, 3)) or is_large or is_mountain

        leaves_str = cell(row_vals, 17).strip().replace(".0", "")
        leaves_after_section = int(leaves_str) if leaves_str else None

        participants[p_id] = Participant(
            id=p_id,
            name=cell(row_vals, 0),
            preferred_sections=preferred,
            can_drive=is_drive,
            can_drive_large=is_large,
            can_drive_mountain=is_mountain,
            staying_overnight=_flag(cell(row_vals, 2)),
            grade=_int_or(cell(row_vals, 1), 1),
            remaining_sections=_int_or(cell(row_vals, 16), 0),
            leaves_after_section=leaves_after_section,
        )
    return participants


def load_participants_from_public_form_sheet(url: str) -> Dict[str, Participant]:
    """Googleフォームの回答スプレッドシートを認証なしで読み込む。"""
    rows = _fetch_csv_rows(url)
    if len(rows) < 2:
        raise ValueError("フォームの回答が0件です。")

    headers = rows[0]
    col_name = _find_col(headers, "名前")
    col_grade = _find_col(headers, "学年")
    col_sections = _find_col(headers, "走りたい区間", exclude="特に")
    col_count = _find_col(headers, "何区間")
    col_drive = _find_col(headers, "普通自動車")
    col_large = _find_col(headers, "大型")
    col_mountain = _find_col(headers, "山道")
    col_stay = _find_col(headers, "宿泊")
    col_leave = _find_col(headers, "帰りますか")
    col_priority = _find_col(headers, "特に走りたい")

    participants = {}
    p_index = 0
    for row_vals in rows[1:]:
        if not any(row_vals):
            continue
        row = dict(zip(headers, row_vals))

        def get(col: str) -> str:
            return str(row.get(col, "")).strip()

        def yes(col: str) -> bool:
            return get(col).startswith("はい")

        def parse_sections(text: str) -> List[bool]:
            flags = [False] * 10
            for item in re.split(r"[,、・/\s　]+", text):
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

        is_large = yes(col_large)
        is_mountain = yes(col_mountain)
        is_drive = yes(col_drive) or is_large or is_mountain

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
    return participants
