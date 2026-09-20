"""裏シュミのCLIエントリ。表シュミの出力xlsxと希望条件(JSON)を読み、
新しいxlsxを書き出す。

使い方:
    python run_back_sim.py [表シュミの出力xlsx] [条件JSON] [時間制限(秒)]

条件JSONの例はback_sim_config.example.jsonを参照。
"""

import json
import sys

from data_io.output_writer import write_plan_xlsx
from data_io.plan_importer import import_participants_from_output_xlsx, import_plan_from_output_xlsx
from logic.back_sim import BackSimConfig, run_back_sim

SOURCE_XLSX = sys.argv[1] if len(sys.argv) > 1 else "hakone_result.xlsx"
CONFIG_JSON = sys.argv[2] if len(sys.argv) > 2 else "back_sim_config.json"
TIME_LIMIT = float(sys.argv[3]) if len(sys.argv) > 3 else 60.0
OUTPUT_XLSX = sys.argv[4] if len(sys.argv) > 4 else "hakone_result_back.xlsx"

with open(CONFIG_JSON, encoding="utf-8") as f:
    raw = json.load(f)

config = BackSimConfig(
    together_pairs=[tuple(pair) for pair in raw.get("together_pairs", [])],
    apart_pairs=[tuple(pair) for pair in raw.get("apart_pairs", [])],
    driver_ranges={name: tuple(rng) for name, rng in raw.get("driver_ranges", {}).items()},
    together_or_groups=[(g["anchor"], list(g["group"])) for g in raw.get("together_or_groups", [])],
)

print(f"📥 表シュミの結果を読み込み中: {SOURCE_XLSX}")
participants = import_participants_from_output_xlsx(SOURCE_XLSX)
baseline_plan = import_plan_from_output_xlsx(SOURCE_XLSX, participants)
print(f"   参加者{len(participants)}人を復元しました。")

plan = run_back_sim(baseline_plan, participants, config, time_limit=TIME_LIMIT)
write_plan_xlsx(plan, participants, OUTPUT_XLSX)
