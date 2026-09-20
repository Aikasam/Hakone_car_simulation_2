import contextlib
import io
import os
from typing import Dict, List

import pandas as pd
import streamlit as st

from data_io.public_sheets_reader import (
    load_participants_from_public_form_sheet,
    load_participants_from_public_sheet,
)
from data_io.sheets_manager import (
    load_participants_from_form_sheet,
    load_participants_from_sheet,
    save_plan_to_sheet,
)
from logic.milp_allocator_v3 import generate_full_plan_cpsat, DEFAULT_TIME_LIMIT
from logic.back_sim import BackSimConfig, run_back_sim, summarize_conditions
from logic.car_pool import section_label, LARGE_CAR_IDS, NORMAL_CAR_IDS
from data_io.output_writer import write_plan_xlsx, SECTION_COLORS
from data_io.plan_importer import (
    fetch_output_format_xlsx_from_google_sheet,
    import_participants_from_output_xlsx,
    import_plan_from_output_xlsx,
)
from validator import (
    validate_participants,
    validate_transitions,
    count_car_changes,
    compute_runner_satisfaction,
    compute_individual_summary,
)

CREDENTIALS_PATH = "credentials.json"
OUTPUT_XLSX_PATH = "hakone_result.xlsx"
BACK_OUTPUT_XLSX_PATH = "hakone_result_back.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

APP_VERSION = "v3-cpsat-2026-09-17"

# 裏シュミを開くための隠しドット。ページ上の複数箇所に見た目が同じドットを散らし、
# この順番(dot_2 → dot_4 → dot_1)でクリックしたときだけ開閉する。それ以外のドット
# (dot_3など)や順番を間違えたクリックは進捗をリセットする「罠」として機能する。
_HIDDEN_SEQUENCE = ["dot_2", "dot_4", "dot_1"]


def _register_hidden_click(spot_id: str) -> None:
    progress = st.session_state.get("hidden_click_progress", 0)
    expected = _HIDDEN_SEQUENCE[progress] if progress < len(_HIDDEN_SEQUENCE) else None
    if spot_id == expected:
        progress += 1
        if progress >= len(_HIDDEN_SEQUENCE):
            st.session_state.show_back_sim = not st.session_state.get("show_back_sim", False)
            progress = 0
        st.session_state.hidden_click_progress = progress
    else:
        # 間違えても、その一手がたまたま最初の一手と一致するならそこからやり直せる
        # ようにする(そうしないと「1手目」を忘れて先頭から全部押し直す羽目になる)。
        st.session_state.hidden_click_progress = 1 if spot_id == _HIDDEN_SEQUENCE[0] else 0


def _hidden_dot(spot_id: str, key: str) -> None:
    """裏シュミを開くための隠しドットを1つ描画する。見た目は他のドットと区別が
    つかない、右端に寄せた小さい灰色の「・」。"""
    st.markdown(
        f"""
        <style>
        div.st-key-{key} button {{
            background: transparent; border: none; color: #cccccc;
            font-size: 11px; padding: 0 2px; min-height: 0; line-height: 1.2;
        }}
        div.st-key-{key} button:hover {{ color: #999999; background: transparent; border: none; }}
        div.st-key-{key} button:focus:not(:active) {{ color: #999999; box-shadow: none; }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    _spacer, _dot_col = st.columns([40, 1])
    with _dot_col:
        if st.button("·", key=key):
            _register_hidden_click(spot_id)


def _car_id_of_cell(text):
    """個人別まとめの1セルのテキスト(例: "🚘 運転手(L1)")から車両IDを取り出す。
    車に乗っていない役割(ランナー・現地集合・不参加)ならNone。"""
    import re
    if not isinstance(text, str):
        return None
    m = re.search(r"\(([^)]+)\)\s*$", text)
    return m.group(1) if m else None


def _condition_cell_styles(table: pd.DataFrame, together_map: dict, apart_map: dict, labels: list) -> pd.DataFrame:
    """一緒にしたい/離したい条件が付いた人について、名前欄だけでなく、実際にその
    条件を満たせている区間のセルにも同じ色・太字を適用するためのスタイル表を作る
    (Styler.apply(axis=None)用。tableは元データそのまま、返り値は同じ形のCSS文字列表)。
    「一緒にしたい」は同じ車になれていれば満たしている、「離したい」は指定した
    相手全員と同じ車になっていなければ満たしている、とみなす。"""
    name_to_row = {name: idx for idx, name in zip(table.index, table["名前"])}
    styles = pd.DataFrame("", index=table.index, columns=table.columns)

    for idx, row in table.iterrows():
        name = row["名前"]
        together_targets = together_map.get(name, set())
        apart_targets = apart_map.get(name, set())
        if not together_targets and not apart_targets:
            continue
        for label in labels:
            my_car = _car_id_of_cell(row[label])
            if my_car is None:
                continue

            def _shares_car_with(target_name):
                t_idx = name_to_row.get(target_name)
                if t_idx is None:
                    return False
                return _car_id_of_cell(table.at[t_idx, label]) == my_car

            if together_targets and any(_shares_car_with(t) for t in together_targets):
                styles.at[idx, label] = "color: #CC3333; font-weight: bold"
            elif apart_targets and not any(_shares_car_with(t) for t in apart_targets):
                styles.at[idx, label] = "color: #1E5FCC; font-weight: bold"

    return styles


def _render_plan_result(plan, participants, xlsx_path, key_prefix, download_label, highlight=None):
    """表シュミ・裏シュミの結果表示で共通する部分(区間ごとの配車・個人別まとめ・
    メトリクス・Excelダウンロード)をまとめた描画関数。
    highlight: 裏シュミの条件を個人別まとめ表で見分けやすくするための追加情報
    (dict: together_names/apart_names/driver_range_names のset、
     together_map/apart_map の 人名->相手名setで区間セルごとの充足判定に使う)。
    Noneなら何も強調しない。"""
    used_cars = len({car.car_id for section in plan for car in section.cars})
    large_n = len({car.car_id for section in plan for car in section.cars if car.car_type == "large"})
    normal_n = used_cars - large_n
    st.success(f"計算完了！　大型 {large_n} 台 ＋ 普通 {normal_n} 台 ＝ 合計 {used_cars} 台")

    for section in plan:
        label = section_label(section.section_id)
        runners = [participants[pid].name for pid in section.runner_ids if pid in participants]
        runner_text = f"🏃 ランナー {len(runners)} 名: {', '.join(runners)}" if runners else "（ランナーなし）"

        with st.expander(f"【{label}】　{runner_text}"):
            for car in section.cars:
                driver_name = participants[car.driver_id].name if car.driver_id in participants else "エラー"
                passengers = [participants[pid].name for pid in car.passenger_ids if pid in participants]
                car_type = "大型" if car.car_type == "large" else "普通"
                mt_badge = "　⛰️ 山行き" if car.is_mountain_goer else ""
                hotel_badge = "　🏨 ホテル組" if car.group == "hotel" else ""
                adv_badge = "　🚀 先行" if car.is_advance else ""

                st.markdown(f"**🚘 車 {car.car_id}**（{car_type}）{mt_badge}{hotel_badge}{adv_badge}")
                st.write(f"　👨‍✈️ 運転手: {driver_name}")
                st.write(f"　👥 同乗者: {', '.join(passengers) if passengers else 'なし'}")
                st.divider()

    with st.expander("👤 個人別まとめ（区間ごとの役割）"):
        summary = compute_individual_summary(plan, participants)
        labels = [section_label(s) for s in list(range(1, 11)) + [11]]
        table = pd.DataFrame(
            [
                {"名前": p.name, **{label: summary.get(pid, {}).get(label, "") for label in labels}}
                for pid, p in participants.items()
            ]
        )

        def _cell_color(text):
            if text == "現地集合":
                return "background-color: #D9B3FF"
            if isinstance(text, str) and text.startswith("🏃"):
                return "background-color: #FFE699"
            if isinstance(text, str) and text.startswith("🚘"):
                return "background-color: #BDD7EE"
            if isinstance(text, str) and text.startswith("👥"):
                return "background-color: #C6E0B4"
            return "background-color: #FFFFFF"

        styled_table = table.style.map(_cell_color, subset=labels)
        caption = "🟡 ランナー　🔵 運転手　🟢 同乗者　🟣 現地集合(1区は走らず2区から参加)　⚪ 不参加"

        if highlight:
            together_names = highlight.get("together_names", set())
            apart_names = highlight.get("apart_names", set())
            driver_range_names = highlight.get("driver_range_names", set())

            def _name_color(name):
                in_together = name in together_names
                in_apart = name in apart_names
                if in_together and in_apart:
                    return "color: #A000A0; font-weight: bold"  # 両方指定されている人は紫で目立たせる
                if in_apart:
                    return "color: #1E5FCC; font-weight: bold"
                if in_together:
                    return "color: #CC3333; font-weight: bold"
                return ""

            def _bold_driver_row(row):
                if row["名前"] not in driver_range_names:
                    return ["" for _ in row.index]
                return [
                    "font-weight: bold" if (col in labels and isinstance(row[col], str) and row[col].startswith("🚘")) else ""
                    for col in row.index
                ]

            styled_table = styled_table.map(_name_color, subset=["名前"])
            styled_table = styled_table.apply(_bold_driver_row, axis=1)

            together_map = highlight.get("together_map", {})
            apart_map = highlight.get("apart_map", {})
            if together_map or apart_map:
                styled_table = styled_table.apply(
                    _condition_cell_styles, axis=None,
                    together_map=together_map, apart_map=apart_map, labels=labels,
                )

            caption += (
                "　｜　🔴 一緒にしたい人／一緒になれた区間　🔵 離したい人／離せた区間"
                "　🟣 両方指定　太字の🚘=運転区間数を指定した人"
            )

        st.dataframe(styled_table, hide_index=True, use_container_width=True, key=f"{key_prefix}_table")
        st.caption(caption)

    with st.expander("🚘 区間別配車（Excelの「区間別配車」シートと同じ形式）"):
        _render_car_assignment_table(plan, participants, key_prefix=key_prefix)

    transition_errors = validate_transitions(plan, participants)
    if transition_errors:
        st.warning("⚠️ 車両引き継ぎに問題があります: " + " / ".join(transition_errors))

    changes = count_car_changes(plan)
    n_wanted, n_satisfied, total_wanted, total_ran = compute_runner_satisfaction(plan, participants)
    col1, col2 = st.columns(2)
    col1.metric("ランナー希望充足", f"{n_satisfied}/{n_wanted}人", help="希望区間数を達成できた人の割合")
    col2.metric("走行区間数", f"{total_ran}/{total_wanted}区間", help="実際に走れた区間 / 希望合計")
    st.caption(f"ℹ️ 区間をまたいで車を乗り換えた回数: {changes}回")

    if os.path.exists(xlsx_path):
        with open(xlsx_path, "rb") as f:
            st.download_button(
                download_label,
                f,
                file_name=os.path.basename(xlsx_path),
                mime=XLSX_MIME,
                type="primary",
                key=f"{key_prefix}_download",
            )


# 区間ごとの色分け(1〜10区)。ユーザー指定の配色見本から実際のピクセル色を抽出した値。
def _render_section_color_legend():
    """「次に走る区間」の色分け(SECTION_COLORS)が、どの区間がどの色かを示す凡例。"""
    chips = "".join(
        f'<span style="background-color:{SECTION_COLORS[s]}; padding:2px 10px; '
        f'margin-right:6px; border-radius:4px; display:inline-block; font-size:0.85em;">{s}区</span>'
        for s in range(1, 11)
    )
    st.markdown(
        f'<div>🎨 セルの色（運転手・同乗者が次に走る区間）: {chips}</div>',
        unsafe_allow_html=True,
    )


def _render_car_assignment_table(plan, participants, key_prefix):
    """Excelの「区間別配車」シートと同じ形式(1区間=1ブロック、車ID・車種・山行き・
    先行・運転手・同乗者…)の表を画面上にも出す。列構成・基本の表示テキストは
    output_writer側の_car_blocks/_cars_headerをそのまま再利用し、「先行」列の上書き・
    セルの背景色もoutput_writer.compute_car_table_overridesを共用することで、
    Excel出力(区間別配車シート)と画面表示が完全に一致するようにしている。"""
    from data_io.output_writer import _car_blocks, _cars_header, compute_car_table_overrides

    blocks = _car_blocks(plan, participants)
    header = _cars_header(blocks)
    n_cols = len(header)
    advance_col = header.index("先行")

    rows = []
    for label, section_rows in blocks:
        if not section_rows:
            rows.append([label] + [""] * (n_cols - 1))
            continue
        for i, row in enumerate(section_rows):
            padded = list(row) + [""] * (n_cols - 1 - len(row))
            rows.append([label if i == 0 else ""] + padded)

    advance_overrides, cell_fills = compute_car_table_overrides(plan, participants, header)
    for i, override in enumerate(advance_overrides):
        if override:
            rows[i][advance_col] = override

    table = pd.DataFrame(rows, columns=header).fillna("")

    def _next_run_cell_styles(_data: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=table.index, columns=table.columns)
        for i, fills in enumerate(cell_fills):
            for col, color in fills.items():
                styles.iat[i, col] = f"background-color: {color}"
        return styles

    styled_table = table.style.apply(_next_run_cell_styles, axis=None)
    st.dataframe(styled_table, hide_index=True, use_container_width=True, key=f"{key_prefix}_car_table")
    st.caption(
        "「先行」列 = その車の乗車メンバーが次に変わる区間"
        "（直前区間のランナーを新たに乗せた行は「走者回収&◯区」）"
    )
    _render_section_color_legend()


def _render_back_sim_section():
    """裏シュミ(追加条件での再調整)。表シュミを実行していなくても、表シュミ出力形式の
    xlsxファイルを直接アップロードすれば単独で使える。"""
    st.subheader("🔧 裏シュミ（追加条件で再調整）")
    st.caption(
        "既存の配車結果を土台に、「できるだけ一緒にしたい」「離したい」「運転区間数の範囲」を"
        "追加のお願いとして加え、既存の割り当てをできるだけ変えずに調整します。"
        "免許・定員などの必須条件や、表シュミ側の希望(特に「特に走りたい区間」)はここでの新しい条件より優先されます。"
    )

    has_forward_result = bool(st.session_state.get("result"))
    has_saved_output = os.path.exists(OUTPUT_XLSX_PATH)

    source_options = []
    if has_forward_result:
        source_options.append("直前の表シュミ結果を使う")
    if has_saved_output:
        source_options.append("保存済みの表シュミ結果ファイルを使う")
    source_options.append("表シュミの結果xlsxをアップロード")
    source_options.append("表シュミの結果と同じ形式のスプレッドシートのURLを指定")

    source = st.radio("入力元", source_options, horizontal=True, key="back_source")

    plan = None
    participants = None

    if source == "直前の表シュミ結果を使う":
        plan = st.session_state["result"]["plan"]
        participants = st.session_state["result"]["participants"]
    elif source == "保存済みの表シュミ結果ファイルを使う":
        st.caption(f"前回保存された結果ファイル（{OUTPUT_XLSX_PATH}）を読み込みます。")
        try:
            participants = import_participants_from_output_xlsx(OUTPUT_XLSX_PATH)
            plan = import_plan_from_output_xlsx(OUTPUT_XLSX_PATH, participants)
            st.success(f"参加者 {len(participants)} 名のデータを読み込みました。")
        except Exception as e:
            st.error(f"ファイルの読み込みに失敗しました。\n\n{e}")
    elif source == "表シュミの結果と同じ形式のスプレッドシートのURLを指定":
        sheet_url = st.text_input(
            "スプレッドシートのURL（入力データ／区間別ランナー／区間別配車の3シートを含む、"
            "「リンクを知っている全員が閲覧可」に共有されたもの）",
            key="back_sheet_url",
        )
        if sheet_url:
            try:
                data = fetch_output_format_xlsx_from_google_sheet(sheet_url).getvalue()
                participants = import_participants_from_output_xlsx(io.BytesIO(data))
                plan = import_plan_from_output_xlsx(io.BytesIO(data), participants)
                st.success(f"参加者 {len(participants)} 名のデータを読み込みました。")
            except Exception as e:
                st.error(f"スプレッドシートの読み込みに失敗しました。表シュミの出力形式か、共有設定を確認してください。\n\n{e}")
    else:
        uploaded = st.file_uploader(
            "表シュミが出力したxlsxファイル（入力データ／区間別ランナー／区間別配車の3シートを含むもの）",
            type=["xlsx"], key="back_uploaded_file",
        )
        if uploaded is not None:
            try:
                data = uploaded.getvalue()
                participants = import_participants_from_output_xlsx(io.BytesIO(data))
                plan = import_plan_from_output_xlsx(io.BytesIO(data), participants)
                st.success(f"参加者 {len(participants)} 名のデータを読み込みました。")
            except Exception as e:
                st.error(f"ファイルの読み込みに失敗しました。表シュミの出力形式のxlsxか確認してください。\n\n{e}")

    if plan is None or participants is None:
        st.info("裏シュミを実行するには、上で入力元を指定してください。")
        return

    # 学年順(同学年内は名前順)に並べる。一括選択(学年ボタン)にも使う。
    participants_sorted = sorted(participants.values(), key=lambda p: (p.grade, p.name))
    participant_names = [p.name for p in participants_sorted]
    driver_names = [p.name for p in participants_sorted if p.can_drive]
    names_by_grade: Dict[int, List[str]] = {}
    for p in participants_sorted:
        names_by_grade.setdefault(p.grade, []).append(p.name)
    grade_options = sorted(names_by_grade.keys())

    def _add_grade_members(multiselect_key, grade, exclude_name):
        current = list(st.session_state.get(multiselect_key, []))
        additions = [n for n in names_by_grade.get(grade, []) if n != exclude_name and n not in current]
        st.session_state[multiselect_key] = current + additions

    def _remove_group(count_key: str, prefix: str, fields: List[str], idx: int, total: int):
        """グループidxを削除し、それ以降のグループの入力値を1つずつ前に詰める。"""
        for j in range(idx, total - 1):
            for field in fields:
                src, dst = f"{prefix}_{field}_{j + 1}", f"{prefix}_{field}_{j}"
                if src in st.session_state:
                    st.session_state[dst] = st.session_state[src]
                elif dst in st.session_state:
                    del st.session_state[dst]
        last = total - 1
        for field in fields:
            st.session_state.pop(f"{prefix}_{field}_{last}", None)
        st.session_state[count_key] = total - 1

    NO_SELECTION = "（未選択）"
    anchor_options = [NO_SELECTION] + participant_names

    with st.expander("裏シュミの条件を設定", expanded=not st.session_state.get("back_result")):
        st.markdown("**できるだけ一緒にしたい**")
        n_together = st.number_input(
            "グループ数", min_value=0, max_value=10, step=1, key="n_together",
        )
        together_pairs = []
        together_or_groups = []
        for i in range(int(n_together)):
            c1, c2 = st.columns(2)
            anchor = c1.selectbox("基準となる人", anchor_options, key=f"together_anchor_{i}")
            companions = c2.multiselect(
                "一緒にしたい相手（複数選択可）",
                [n for n in participant_names if n != anchor],
                key=f"together_companions_{i}",
            )
            gcol1, gcol2 = st.columns([3, 1])
            grade_pick = gcol1.selectbox(
                "学年で一括追加", grade_options, key=f"together_grade_pick_{i}", label_visibility="collapsed",
            )
            gcol2.button(
                f"{grade_pick}年を追加", key=f"together_grade_btn_{i}",
                on_click=_add_grade_members, args=(f"together_companions_{i}", grade_pick, anchor),
            )
            mode = st.radio(
                "複数選んだ場合の扱い",
                ["それぞれとできるだけ一緒に（AND）", "誰か一人とでも一緒ならOK（OR）"],
                key=f"together_mode_{i}", horizontal=True,
            )
            st.button(
                "🗑️ このグループを削除", key=f"together_remove_{i}",
                on_click=_remove_group, args=("n_together", "together", ["anchor", "companions", "mode"], i, int(n_together)),
            )
            if anchor != NO_SELECTION and companions:
                if mode.startswith("誰か"):
                    together_or_groups.append((anchor, companions))
                else:
                    together_pairs += [(anchor, comp) for comp in companions]
            st.divider()

        st.markdown("**できるだけ離したい**")
        n_apart = st.number_input(
            "グループ数", min_value=0, max_value=10, step=1, key="n_apart",
        )
        apart_pairs = []
        for i in range(int(n_apart)):
            c1, c2 = st.columns(2)
            anchor = c1.selectbox("基準となる人", anchor_options, key=f"apart_anchor_{i}")
            companions = c2.multiselect(
                "離したい相手（複数選択可）",
                [n for n in participant_names if n != anchor],
                key=f"apart_companions_{i}",
            )
            gcol1, gcol2 = st.columns([3, 1])
            grade_pick = gcol1.selectbox(
                "学年で一括追加", grade_options, key=f"apart_grade_pick_{i}", label_visibility="collapsed",
            )
            gcol2.button(
                f"{grade_pick}年を追加", key=f"apart_grade_btn_{i}",
                on_click=_add_grade_members, args=(f"apart_companions_{i}", grade_pick, anchor),
            )
            st.button(
                "🗑️ このグループを削除", key=f"apart_remove_{i}",
                on_click=_remove_group, args=("n_apart", "apart", ["anchor", "companions"], i, int(n_apart)),
            )
            if anchor != NO_SELECTION and companions:
                apart_pairs += [(anchor, comp) for comp in companions]
            st.divider()

        st.markdown("**運転区間数の範囲**")
        n_driver_ranges = st.number_input(
            "指定する人数", min_value=0, max_value=20, value=0, step=1, key="n_driver_ranges",
        )
        driver_ranges = {}
        DRIVER_RANGE_NO_LIMIT = 11
        for i in range(int(n_driver_ranges)):
            c1, c2, c3 = st.columns(3)
            person = c1.selectbox("対象者", driver_names, key=f"driver_range_person_{i}")
            min_r = c2.number_input(
                "最少運転区間数", min_value=0, max_value=DRIVER_RANGE_NO_LIMIT, value=0, step=1,
                key=f"driver_range_min_{i}",
            )
            max_r = c3.number_input(
                "最大運転区間数（11で無制限）", min_value=0, max_value=DRIVER_RANGE_NO_LIMIT,
                value=DRIVER_RANGE_NO_LIMIT, step=1, key=f"driver_range_max_{i}",
            )
            driver_ranges[person] = (int(min_r), None if int(max_r) >= DRIVER_RANGE_NO_LIMIT else int(max_r))

        back_time_limit = st.number_input(
            "裏シュミの計算制限時間（秒/フェーズ）", min_value=10, max_value=600, value=60, step=10,
            key="back_time_limit",
        )

        has_any_condition = bool(together_pairs or together_or_groups or apart_pairs or driver_ranges)
        run_back = st.button(
            "裏シュミを実行",
            type="primary",
            help=None if has_any_condition else "条件が指定されていないため、再計算せずそのまま結果を表示します（区間別配車の表の確認などに）。",
        )

    if run_back:
        back_log_buf = io.StringIO()
        config = BackSimConfig(
            together_pairs=together_pairs, apart_pairs=apart_pairs, driver_ranges=driver_ranges,
            together_or_groups=together_or_groups,
        )
        if not has_any_condition:
            back_log_buf.write("条件が指定されていないため、計算をせずに読み込んだ結果をそのまま表示します。\n")
            back_plan = plan
            write_plan_xlsx(back_plan, participants, BACK_OUTPUT_XLSX_PATH)
            condition_summary = summarize_conditions(back_plan, participants, config)
        else:
            with st.spinner("裏シュミを計算中..."):
                try:
                    with contextlib.redirect_stdout(back_log_buf):
                        back_plan = run_back_sim(plan, participants, config, time_limit=back_time_limit)
                        write_plan_xlsx(back_plan, participants, BACK_OUTPUT_XLSX_PATH)
                        condition_summary = summarize_conditions(back_plan, participants, config)
                except Exception as e:
                    st.error(f"裏シュミの計算に失敗しました。\n\n{e}")
                    with st.expander("詳細ログ"):
                        st.text(back_log_buf.getvalue())
                    st.stop()

        together_map: Dict[str, set] = {}
        for a, b in together_pairs:
            together_map.setdefault(a, set()).add(b)
            together_map.setdefault(b, set()).add(a)
        for or_anchor, or_group in together_or_groups:
            for g in or_group:
                together_map.setdefault(or_anchor, set()).add(g)
                together_map.setdefault(g, set()).add(or_anchor)

        apart_map: Dict[str, set] = {}
        for a, b in apart_pairs:
            apart_map.setdefault(a, set()).add(b)
            apart_map.setdefault(b, set()).add(a)

        st.session_state.back_result = {
            "plan": back_plan,
            "participants": participants,
            "log": back_log_buf.getvalue(),
            "summary": condition_summary,
            "highlight": {
                "together_names": set(together_map.keys()),
                "apart_names": set(apart_map.keys()),
                "driver_range_names": set(driver_ranges.keys()),
                "together_map": together_map,
                "apart_map": apart_map,
            },
        }

    back_result = st.session_state.get("back_result")
    if back_result:
        st.markdown("#### 裏シュミの結果")

        summary = back_result.get("summary", {})
        if summary.get("together") or summary.get("together_or") or summary.get("apart") or summary.get("driver_ranges"):
            st.markdown("**条件の充足状況**")
            if summary.get("together"):
                st.caption("できるだけ一緒にしたい（同乗できた区間数）")
                st.dataframe(
                    pd.DataFrame([
                        {"人物A": t["a"], "人物B": t["b"], "同乗できた区間数": t["shared_sections"]}
                        for t in summary["together"]
                    ]),
                    hide_index=True, use_container_width=True, key="back_summary_together",
                )
            if summary.get("together_or"):
                st.caption("できるだけ一緒にしたい（候補のうち誰か一人とでも同乗できた区間数）")
                st.dataframe(
                    pd.DataFrame([
                        {
                            "基準となる人": t["anchor"],
                            "候補": "、".join(t["group"]),
                            "満たせた区間数": t["satisfied_sections"],
                        }
                        for t in summary["together_or"]
                    ]),
                    hide_index=True, use_container_width=True, key="back_summary_together_or",
                )
            if summary.get("apart"):
                st.caption("できるだけ離したい（同乗してしまった区間数。0が理想）")
                st.dataframe(
                    pd.DataFrame([
                        {"人物A": t["a"], "人物B": t["b"], "同乗してしまった区間数": t["shared_sections"]}
                        for t in summary["apart"]
                    ]),
                    hide_index=True, use_container_width=True, key="back_summary_apart",
                )
            if summary.get("driver_ranges"):
                st.caption("運転区間数の範囲")
                st.dataframe(
                    pd.DataFrame([
                        {
                            "名前": d["name"],
                            "指定範囲": f"{d['min']}〜{d['max'] if d['max'] is not None else '無制限'}",
                            "実際の運転区間数": d["actual"],
                            "達成": "✅" if d["satisfied"] else "❌",
                        }
                        for d in summary["driver_ranges"]
                    ]),
                    hide_index=True, use_container_width=True, key="back_summary_driver_ranges",
                )

        _render_plan_result(
            back_result["plan"], back_result["participants"], BACK_OUTPUT_XLSX_PATH, key_prefix="back",
            download_label="📊 裏シュミの結果Excelをダウンロード",
            highlight=back_result.get("highlight"),
        )
        with st.expander("裏シュミの詳細ログ"):
            st.text(back_result["log"])


# =============================================================================
# 画面構成
# =============================================================================

st.set_page_config(page_title="箱根駅伝配車シミュレーター (CP-SAT版)", page_icon="🚗")
st.title("🚗 箱根駅伝配車シミュレーター")
st.caption(f"ver {APP_VERSION}　(計算エンジン: OR-Tools CP-SAT)")
_hidden_dot("dot_1", "hidden_dot_1")
st.divider()

# --- 表シュミ ---
input_format = st.radio(
    "入力データの形式",
    ["Googleフォームの回答", "独自フォーマット"],
    horizontal=True,
)

if input_format == "Googleフォームの回答":
    st.caption("Googleフォームの回答が集まったスプレッドシートのURLを貼り付けてください。")
else:
    st.caption("所定のフォーマット（名前・1〜10・運転・大・山・宿泊・学年・希望区間数・離脱区間）のスプレッドシートのURLを貼り付けてください。")

url = st.text_input(
    "スプレッドシートのURL",
    placeholder="https://docs.google.com/spreadsheets/d/...",
)

auth_mode = st.radio(
    "スプレッドシートへのアクセス方法",
    ["認証なし（読み込み専用・推奨）", "サービスアカウント認証（読み込み＋書き戻し）"],
    horizontal=True,
    help=(
        "「認証なし」は、スプレッドシートが「リンクを知っている全員が閲覧可」に共有されていれば"
        "Google Cloudの設定が一切不要です。結果は画面表示とExcelダウンロードで受け取ります。\n\n"
        "「サービスアカウント認証」は結果をスプレッドシートに書き戻せますが、"
        "credentials.json（Google Cloudのサービスアカウントキー）が必要です。"
    ),
)
use_auth = auth_mode.startswith("サービスアカウント")
if use_auth and not os.path.exists(CREDENTIALS_PATH):
    st.warning(f"credentials.json が見つかりません（{os.path.abspath(CREDENTIALS_PATH)}）。認証なしモードに切り替えてください。")
_hidden_dot("dot_3", "hidden_dot_3")

st.subheader("車両スロット設定")
col_l, col_n = st.columns(2)
n_large_slots = col_l.number_input(
    "大型車（8人乗り）スロット数", min_value=0, max_value=len(LARGE_CAR_IDS), value=0, step=1
)
n_normal_slots = col_n.number_input(
    "普通車（4人乗り）スロット数", min_value=0, max_value=len(NORMAL_CAR_IDS), value=len(NORMAL_CAR_IDS), step=1
)
active_car_ids = LARGE_CAR_IDS[:n_large_slots] + NORMAL_CAR_IDS[:n_normal_slots]
st.caption(f"最大定員: 大型 {n_large_slots}台×8人 + 普通 {n_normal_slots}台×4人 = {n_large_slots*8 + n_normal_slots*4}人")
_hidden_dot("dot_2", "hidden_dot_2")

NO_LIMIT = 99  # 「無制限」を表す便宜上の大きな値(実際の参加人数を超えるので事実上上限なしになる)


def _apply_bulk_limits():
    bulk_min = st.session_state.get("bulk_limit_min", 1)
    bulk_max = st.session_state.get("bulk_limit_max", NO_LIMIT)
    for s in range(1, 11):
        st.session_state[f"limit_min_{s}"] = int(bulk_min)
        st.session_state[f"limit_max_{s}"] = int(bulk_max)


runner_limits = {}
with st.expander("区間ごとの人数制限（任意）"):
    st.caption(
        "各区間を走る人数の下限・上限を指定できます。最大人数は99のままにすると"
        "実質「上限なし」として扱われます。指定しない区間は最少1人・上限なしの既定値になります。"
    )

    st.markdown("**一括設定**（下の値を1〜10区すべてに適用します）")
    bulk_col_min, bulk_col_max = st.columns(2)
    bulk_col_min.number_input(
        "最少人数（一括）", min_value=0, max_value=NO_LIMIT, value=1, step=1, key="bulk_limit_min",
    )
    bulk_col_max.number_input(
        "最大人数（一括）", min_value=0, max_value=NO_LIMIT, value=NO_LIMIT, step=1, key="bulk_limit_max",
    )
    st.button("↑ 1〜10区すべてに適用", on_click=_apply_bulk_limits, key="apply_bulk_limits_btn")

    st.divider()

    col_label, col_min, col_max = st.columns([1, 2, 2])
    col_min.markdown("**最少人数**")
    col_max.markdown("**最大人数**")
    for s in range(1, 11):
        st.session_state.setdefault(f"limit_min_{s}", 1)
        st.session_state.setdefault(f"limit_max_{s}", NO_LIMIT)
        col_label, col_min, col_max = st.columns([1, 2, 2])
        col_label.markdown(f"{s}区")
        min_n = col_min.number_input(
            f"{s}区の最少人数", min_value=0, max_value=NO_LIMIT, step=1,
            key=f"limit_min_{s}", label_visibility="collapsed",
        )
        max_n = col_max.number_input(
            f"{s}区の最大人数", min_value=0, max_value=NO_LIMIT, step=1,
            key=f"limit_max_{s}", label_visibility="collapsed",
        )
        runner_limits[s] = (int(min_n), None if int(max_n) >= NO_LIMIT else int(max_n))

time_limit = st.number_input(
    "1ブロックあたりの計算制限時間（秒）", min_value=10, max_value=600, value=int(DEFAULT_TIME_LIMIT), step=10,
    help="大型車が少ない構成では最適解を見つけるのに数分かかることがあります。",
)

can_write_back = use_auth and os.path.exists(CREDENTIALS_PATH)
write_back = st.checkbox(
    "結果をスプレッドシートに書き戻す（「入力データ」「配車結果」シートを作成/上書き）",
    value=can_write_back,
    disabled=not can_write_back,
    help="「サービスアカウント認証」を選び、credentials.jsonが置かれている場合のみ有効になります。" if not can_write_back else None,
)

if st.button("シミュレーション実行", type="primary", disabled=not url or not active_car_ids):
    # 計算結果はすべてsession_stateに保存する。
    # (Excelダウンロードボタンを押すとStreamlitはスクリプト全体を再実行するが、
    #  そのときst.button()はFalseに戻るため、結果をsession_state以外に置いていると
    #  再実行のたびに画面から消えてしまう。)
    log_buf = io.StringIO()

    with st.spinner("データを読み込み中..."):
        try:
            with contextlib.redirect_stdout(log_buf):
                if use_auth:
                    if input_format == "Googleフォームの回答":
                        participants = load_participants_from_form_sheet(url, CREDENTIALS_PATH)
                    else:
                        participants = load_participants_from_sheet(url, CREDENTIALS_PATH)
                else:
                    if input_format == "Googleフォームの回答":
                        participants = load_participants_from_public_form_sheet(url)
                    else:
                        participants = load_participants_from_public_sheet(url)
        except Exception as e:
            st.error(f"スプレッドシートの読み込みに失敗しました。\n\n{e}")
            st.stop()

    n_participants = len(participants)
    input_warnings = validate_participants(participants)

    with st.spinner(f"配車を計算中（大型車が少ない構成では最大{time_limit*2:.0f}秒程度かかる場合があります）..."):
        try:
            with contextlib.redirect_stdout(log_buf):
                plan = generate_full_plan_cpsat(
                    participants,
                    active_car_ids=active_car_ids,
                    time_limit=time_limit,
                    output_path=OUTPUT_XLSX_PATH,
                    runner_limits=runner_limits,
                )
        except Exception as e:
            st.error(f"計算に失敗しました。\n\n{e}")
            with st.expander("詳細ログ"):
                st.text(log_buf.getvalue())
            st.stop()

    st.session_state.result = {
        "plan": plan,
        "participants": participants,
        "n_participants": n_participants,
        "input_warnings": input_warnings,
        "log": log_buf.getvalue(),
        "write_back_requested": write_back,
        "write_back_done": False,
        "write_back_url": None,
        "write_back_error": None,
        "url": url,
    }

# --- 結果表示ブロック: session_stateに結果がある限り、再実行のたびに描画する ---
result = st.session_state.get("result")
if result:
    plan = result["plan"]
    participants = result["participants"]

    st.info(f"参加者 {result['n_participants']} 名のデータを読み込みました。")

    if result["input_warnings"]:
        with st.expander(f"⚠️ 入力データに {len(result['input_warnings'])} 件の問題があります"):
            for w in result["input_warnings"]:
                st.warning(w)

    _render_plan_result(
        plan, participants, OUTPUT_XLSX_PATH, key_prefix="fwd",
        download_label="📊 結果Excelをダウンロード（入力データ／区間別ランナー／区間別配車／個人別まとめの4シート）",
    )

    if result["write_back_requested"] and not result["write_back_done"]:
        with st.spinner("スプレッドシートに結果を書き込み中..."):
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    result["write_back_url"] = save_plan_to_sheet(plan, participants, result["url"], CREDENTIALS_PATH)
            except Exception as e:
                result["write_back_error"] = str(e)
            result["write_back_done"] = True

    if result["write_back_url"]:
        st.link_button("📄 結果のスプレッドシートを開く（入力データ／配車結果シート）", result["write_back_url"], type="primary")
    if result["write_back_error"]:
        st.error(f"スプレッドシートへの書き出しに失敗しました。\n\n{result['write_back_error']}")

    with st.expander("詳細ログ"):
        st.text(result["log"])

# --- 裏シュミへの入口(ページ最後のドット) ---
# 表シュミの結果はここに来るまでのスクリプト実行で既に確定しているので、そのまま使える。
st.divider()
_hidden_dot("dot_4", "hidden_dot_4")

if st.session_state.get("show_back_sim"):
    st.divider()
    _render_back_sim_section()
