import contextlib
import io
import os

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
from logic.car_pool import section_label, LARGE_CAR_IDS, NORMAL_CAR_IDS
from validator import (
    validate_participants,
    validate_transitions,
    count_car_changes,
    compute_runner_satisfaction,
    compute_individual_summary,
)

CREDENTIALS_PATH = "credentials.json"
OUTPUT_XLSX_PATH = "hakone_result.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

APP_VERSION = "v3-cpsat-2026-09-17"

st.set_page_config(page_title="箱根駅伝配車シミュレーター (CP-SAT版)", page_icon="🚗")
st.title("🚗 箱根駅伝配車シミュレーター")
st.caption(f"ver {APP_VERSION}　(計算エンジン: OR-Tools CP-SAT)")

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
        st.dataframe(table, hide_index=True, use_container_width=True)

    transition_errors = validate_transitions(plan, participants)
    if transition_errors:
        st.warning("⚠️ 車両引き継ぎに問題があります: " + " / ".join(transition_errors))

    changes = count_car_changes(plan)
    n_wanted, n_satisfied, total_wanted, total_ran = compute_runner_satisfaction(plan, participants)
    col1, col2 = st.columns(2)
    col1.metric("ランナー希望充足", f"{n_satisfied}/{n_wanted}人", help="希望区間数を達成できた人の割合")
    col2.metric("走行区間数", f"{total_ran}/{total_wanted}区間", help="実際に走れた区間 / 希望合計")
    st.caption(f"ℹ️ 区間をまたいで車を乗り換えた回数: {changes}回")

    if os.path.exists(OUTPUT_XLSX_PATH):
        with open(OUTPUT_XLSX_PATH, "rb") as f:
            st.download_button(
                "📊 結果Excelをダウンロード（入力データ／区間別ランナー／区間別配車／個人別まとめの4シート）",
                f,
                file_name=os.path.basename(OUTPUT_XLSX_PATH),
                mime=XLSX_MIME,
                type="primary",
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
