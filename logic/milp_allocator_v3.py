"""箱根駅伝配車シミュレーター 本格実装版（OR-Tools CP-SAT）。

logic/milp_allocator.py (PuLP+CBC) の全面置き換え。差分は以下の3点:

  ① 山組/非山組の分離ルールを「全員×全員のペアBig-M制約」から
     「車ごとの is_mtn_car[k,s] フラグに対する集約制約」に書き換え (O(n^2) -> O(n))
     Block A (1〜8区) が最大のボトルネックだったため、ここに適用。
     Block B (9〜10区)・ホテル組・帰路は元コードの時点で既にO(n)の集約制約
     (mtn[p] / mtn_car[k] 方式) だったため、定式化はそのまま維持しCP-SATに移植した。

  ② ソルバーを PuLP+CBC から OR-Tools CP-SAT に変更（並列探索も有効化）。

  ③ 山組ルールの意味論を変更（ユーザーとの合意に基づく）:
     - 山組 = 「9・10区を走りたい希望者」のみ（山道免許の有無では自動的に山組にしない）
     - 車に山フラグが立つ条件 = 運転手 or 同乗者の中に山組の人が一人でもいる
     - 山フラグが立った車の運転手は山道免許必須（ハード制約）
     これにより「山道免許はあるが山には行かない人」を、7・8区で普通の運転要員として
     使えるようになる。実データで検証済み(元の意味論だと非山組の運転手が不足し
     台数によらずInfeasibleになっていたが、新ルールでは解消する)。

目的関数の重み・その他の制約（希望区間・定員・2年生同乗・駐車ペナルティ等）は
元コードと同一に保っている。
"""

import os
import sys
import time
from typing import Dict, List, Optional, Tuple

from ortools.sat.python import cp_model

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Participant, CarState, SectionState
from validator import validate_section
from logic.car_pool import ALL_CAR_IDS, CAR_TYPE, CAR_CAPACITY, CAR_COST, section_label
from data_io.output_writer import write_plan_xlsx

W_FLEET = 5000
W_CONTINUITY = 5
W_RUNNER_PREF = 50
W_PRIORITY_RUNNER_PREF = 150  # 「特に走りたい区間」を実際に走れた場合の追加ボーナス(W_RUNNER_PREFに上乗せ)
W_MTN_RUNNER_DRIVE = 500
W_ADVANCE_SPREAD = 30
W_PREV_RUN_DRIVE = 200
W_PARK = 15000
W_NO_GRADE2 = 100
W_NO_PASSENGER = 50
W_NO_RUN = 500

BLOCK_A_SECTIONS = list(range(1, 9))
BLOCK_B_SECTIONS = [9, 10]
RETURN_TRIP_SECTION_ID = 11

DEFAULT_TIME_LIMIT = 300.0
DEFAULT_WORKERS = 8


def _present(participants: Dict[str, Participant], p: str, s: int) -> bool:
    leaves = participants[p].leaves_after_section
    return leaves is None or s <= leaves


def _needs_return_trip(participants: Dict[str, Participant], p: str) -> bool:
    """帰路の車が必要な人か。leaves_after_section is None(無条件で最後まで残る)だけでなく、
    離脱区間が10区(=最終区間)の人も、10区終了後は他の全員と同じく帰路の車が必要になる。
    「離脱区間=10」と「離脱区間=None」を区別して後者だけを対象にすると、10区の直前まで
    残る日帰り者が帰路の定員計算・配車から漏れてしまう(実際にこれが原因でBlock Bが
    Infeasibleになるケースを確認したため修正)。"""
    leaves = participants[p].leaves_after_section
    return leaves is None or leaves >= 10


def _solve(model: cp_model.CpModel, time_limit: float, workers: int) -> Tuple[cp_model.CpSolver, str]:
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers = workers
    solver.parameters.random_seed = 42
    status = solver.Solve(model)
    return solver, solver.StatusName(status)


# ---------------------------------------------------------------------------
# Block A (1〜8区)
# ---------------------------------------------------------------------------

def _build_block_a(participants: Dict[str, Participant], car_ids: List[str],
                    runner_limits: Optional[Dict[int, Tuple[int, Optional[int]]]] = None):
    runner_limits = runner_limits or {}
    large_ids = [k for k in car_ids if CAR_TYPE[k] == "large"]
    normal_ids = [k for k in car_ids if CAR_TYPE[k] == "normal"]
    pids = list(participants.keys())
    sections = BLOCK_A_SECTIONS

    model = cp_model.CpModel()

    rent = {k: model.NewBoolVar(f"rent_{k}") for k in car_ids}
    for i in range(len(large_ids) - 1):
        model.Add(rent[large_ids[i]] >= rent[large_ids[i + 1]])
    for i in range(len(normal_ids) - 1):
        model.Add(rent[normal_ids[i]] >= rent[normal_ids[i + 1]])

    runs = {}
    for p in pids:
        for s in sections:
            if participants[p].preferred_sections[s - 1] and _present(participants, p, s):
                runs[(p, s)] = model.NewBoolVar(f"runs_{p}_{s}")

    drive = {}
    for p in pids:
        part = participants[p]
        if not part.can_drive:
            continue
        for k in car_ids:
            if CAR_TYPE[k] == "large" and not part.can_drive_large:
                continue
            for s in sections:
                if _present(participants, p, s):
                    drive[(p, k, s)] = model.NewBoolVar(f"drive_{p}_{k}_{s}")

    ride = {
        (p, k, s): model.NewBoolVar(f"ride_{p}_{k}_{s}")
        for p in pids
        for k in car_ids
        for s in sections
        if _present(participants, p, s)
    }

    usedcar = {(k, s): model.NewBoolVar(f"used_{k}_{s}") for k in car_ids for s in sections}

    no_grade2_vars, no_passenger_vars = [], []
    for s in sections:
        for k in car_ids:
            drivers_ks = [drive[(p, k, s)] for p in pids if (p, k, s) in drive]
            model.Add(sum(drivers_ks) == usedcar[(k, s)])
            model.Add(usedcar[(k, s)] <= rent[k])

            riders_ks = [ride[(p, k, s)] for p in pids if (p, k, s) in ride]
            model.Add(sum(riders_ks) <= (CAR_CAPACITY[k] - 1) * usedcar[(k, s)])
            if s != 8:
                v_pass = model.NewBoolVar(f"no_pass_{k}_{s}")
                model.Add(v_pass >= usedcar[(k, s)] - sum(riders_ks))
                no_passenger_vars.append(v_pass)

            grade2_riders = [ride[(p, k, s)] for p in pids if participants[p].grade >= 2 and (p, k, s) in ride]
            if s != 8:
                v_g2 = model.NewBoolVar(f"no_g2_{k}_{s}")
                model.Add(v_g2 >= usedcar[(k, s)] - sum(grade2_riders))
                no_grade2_vars.append(v_g2)

        for p in pids:
            if not _present(participants, p, s):
                continue
            terms = []
            if (p, s) in runs:
                terms.append(runs[(p, s)])
            terms += [drive[(p, k, s)] for k in car_ids if (p, k, s) in drive]
            terms += [ride[(p, k, s)] for k in car_ids if (p, k, s) in ride]
            model.Add(sum(terms) == 1)

        runners_s = [runs[(p, s)] for p in pids if (p, s) in runs]
        min_n, max_n = runner_limits.get(s, (1, None))
        if runners_s:
            model.Add(sum(runners_s) >= min_n)
            if max_n is not None:
                model.Add(sum(runners_s) <= max_n)
        elif min_n > 0:
            print(f"  ⚠️ {s}区: 走れる希望者がいないため、最少人数{min_n}人の指定は満たせません。")

    park_vars = []
    for k in car_ids:
        for s in sections:
            v_park = model.NewBoolVar(f"park_{k}_{s}")
            model.Add(v_park >= rent[k] - usedcar[(k, s)])
            park_vars.append(v_park)

    for p in pids:
        run_vars = [v for (pp, s), v in runs.items() if pp == p]
        if run_vars:
            model.Add(sum(run_vars) <= participants[p].remaining_sections)

    no_run_vars = []
    for p in pids:
        if participants[p].remaining_sections <= 0:
            continue
        run_vars_p = [v for (pp, s), v in runs.items() if pp == p]
        if not run_vars_p:
            continue
        v_no_run = model.NewBoolVar(f"no_run_{p}")
        model.Add(sum(run_vars_p) + v_no_run >= 1)
        no_run_vars.append(v_no_run)

    return_trip_count = sum(1 for p in pids if _needs_return_trip(participants, p))
    model.Add(sum(CAR_CAPACITY[k] * rent[k] for k in car_ids) >= return_trip_count)

    occ = {}
    for p in pids:
        for k in car_ids:
            for s in sections:
                terms = []
                if (p, k, s) in ride:
                    terms.append(ride[(p, k, s)])
                if (p, k, s) in drive:
                    terms.append(drive[(p, k, s)])
                occ[(p, k, s)] = sum(terms) if terms else None

    adv_car_vars = []
    for i in range(len(sections) - 1):
        s, s_next = sections[i], sections[i + 1]
        next_runner_pids = [p for p in pids if (p, s_next) in runs]
        if not next_runner_pids:
            continue
        for k in car_ids:
            adv = model.NewBoolVar(f"adv_{k}_{s}")
            for p in next_runner_pids:
                occ_val = occ[(p, k, s)]
                if occ_val is None:
                    continue
                model.Add(adv >= runs[(p, s_next)] + occ_val - 1)
            adv_car_vars.append(adv)

    # 山組 = 「9・10区を走りたい希望者」のみ（免許の有無は自動的に山組へ入れない）
    mountain_hopefuls = [
        p for p in pids
        if participants[p].preferred_sections[8] or participants[p].preferred_sections[9]
    ]
    mountain_group = set(mountain_hopefuls)
    non_mountain_strict = [p for p in pids if p not in mountain_group]

    is_mtn_car = {(k, s): model.NewBoolVar(f"is_mtn_car_{k}_{s}") for k in car_ids for s in (7, 8)}

    # 6区: 「5区を走った(runs)かつ6区で車kにいる(occ)人」と山行き希望者の同乗を禁止 (AND条件)。
    # 山行き希望者自身が5区を走った本人である場合はこの制約から除外する(自分自身との同乗を
    # 禁止する形になり、5区を走ったあと6区でどの車にも乗れず、6区を走る予定もない場合に
    # Block A全体がINFEASIBLEになるバグがあったため)。
    sec5_runners = [p for p in pids if (p, 5) in runs]
    for k in car_ids:
        for p_mtn in mountain_hopefuls:
            occ_val2 = occ.get((p_mtn, k, 6))
            if occ_val2 is None:
                continue
            run5_at6_vars = []
            for p_run in sec5_runners:
                if p_run == p_mtn:
                    continue
                occ_val = occ.get((p_run, k, 6))
                if occ_val is None:
                    continue
                v = model.NewBoolVar(f"run5at6_{p_run}_{k}_{p_mtn}")
                model.Add(v <= occ_val)
                model.Add(v <= runs[(p_run, 5)])
                model.Add(v >= occ_val + runs[(p_run, 5)] - 1)
                run5_at6_vars.append(v)
            if not run5_at6_vars:
                continue
            has_other_run5 = model.NewBoolVar(f"has_other_run5_{k}_{p_mtn}")
            for v in run5_at6_vars:
                model.Add(has_other_run5 >= v)
            model.Add(occ_val2 + has_other_run5 <= 1)

    # 7〜8区: 運転手 or 同乗者に山組が一人でもいたら、その車は山フラグが立つ
    # (occで判定=運転手も含む)。山フラグが立った車には非山組は同乗できない。
    for s in [7, 8]:
        for k in car_ids:
            for p_mtn in mountain_group:
                mtn_val = occ.get((p_mtn, k, s))
                if mtn_val is not None:
                    model.Add(mtn_val <= is_mtn_car[(k, s)])
            for p_other in non_mountain_strict:
                if (p_other, k, s) in ride:
                    model.Add(ride[(p_other, k, s)] + is_mtn_car[(k, s)] <= 1)

    for k in car_ids:
        model.Add(is_mtn_car[(k, 7)] == is_mtn_car[(k, 8)])

    # 7区のランナーは、山組であっても8区で山フラグ付きの車には回収されない
    # (ride + is_mtn_car + runs(7区) <= 2 : 7区を走った かつ その車が山フラグ付き、の
    #  両方が成立するときだけ ride を0に強制する)
    for k in car_ids:
        for p in pids:
            if (p, 7) not in runs or (p, k, 8) not in ride:
                continue
            model.Add(ride[(p, k, 8)] + is_mtn_car[(k, 8)] + runs[(p, 7)] <= 2)

    # 山フラグが立った車の運転手は山道免許必須（ハード制約）
    for s in [7, 8]:
        for k in car_ids:
            for p in pids:
                if participants[p].can_drive_mountain:
                    continue
                if (p, k, s) in drive:
                    model.Add(drive[(p, k, s)] + is_mtn_car[(k, s)] <= 1)

    for p in mountain_group:
        for k in car_ids:
            occ7 = occ[(p, k, 7)]
            occ8 = occ[(p, k, 8)]
            # 7区で離脱する日帰り者は7区にはいても8区は不在(occ8=None)になり得る。
            # andだと「片方だけNone」のケースを素通りしてしまい、Noneのまま
            # model.Add(occ7 == occ8) に渡ってTypeErrorになる(実際に確認済み)。
            # 他の同種ガード(adv_car_vars, match_vars)と同じくorにする。
            if occ7 is None or occ8 is None:
                continue
            model.Add(occ7 == occ8)

    prev_run_drive_vars = []
    for i in range(len(sections) - 1):
        s_curr, s_next = sections[i], sections[i + 1]
        for p in pids:
            if (p, s_curr) not in runs:
                continue
            drive_next = [drive[(p, k, s_next)] for k in car_ids if (p, k, s_next) in drive]
            if not drive_next:
                continue
            v = model.NewBoolVar(f"prd_{p}_{s_curr}")
            model.Add(v >= runs[(p, s_curr)] + sum(drive_next) - 1)
            prev_run_drive_vars.append(v)

    mountain_runners = [p for p in mountain_hopefuls if not participants[p].can_drive_mountain]
    mtn_runner_drive_vars = []
    for s in [7, 8]:
        for k in car_ids:
            for p_run in mountain_runners:
                if (p_run, k, s) in drive:
                    mtn_runner_drive_vars.append(drive[(p_run, k, s)])

    match_vars = []
    for p in pids:
        for k in car_ids:
            for i in range(1, len(sections)):
                s_prev, s_cur = sections[i - 1], sections[i]
                occ_prev = occ[(p, k, s_prev)]
                occ_cur = occ[(p, k, s_cur)]
                if occ_prev is None or occ_cur is None:
                    continue
                m = model.NewBoolVar(f"match_{p}_{k}_{s_prev}_{s_cur}")
                model.Add(m <= occ_prev)
                model.Add(m <= occ_cur)
                match_vars.append(m)

    # 「特に走りたい区間」を実際に走れた場合、通常の希望区間ボーナスに加えて追加ボーナスを与える
    priority_run_vars = [
        v for (p, s), v in runs.items()
        if participants[p].priority_sections[s - 1]
    ]

    model.Minimize(
        W_FLEET * sum(CAR_COST[k] * rent[k] for k in car_ids)
        - W_CONTINUITY * sum(match_vars)
        - W_RUNNER_PREF * sum(runs.values())
        - W_PRIORITY_RUNNER_PREF * sum(priority_run_vars)
        + W_MTN_RUNNER_DRIVE * sum(mtn_runner_drive_vars)
        + W_ADVANCE_SPREAD * sum(adv_car_vars)
        + W_PREV_RUN_DRIVE * sum(prev_run_drive_vars)
        + W_PARK * sum(park_vars)
        + W_NO_GRADE2 * sum(no_grade2_vars)
        + W_NO_PASSENGER * sum(no_passenger_vars)
        + W_NO_RUN * sum(no_run_vars)
    )

    ctx = dict(rent=rent, runs=runs, drive=drive, ride=ride, usedcar=usedcar,
               is_mtn_car=is_mtn_car, pids=pids, sections=sections, car_ids=car_ids)
    return model, ctx


def _extract_block_a(solver: cp_model.CpSolver, ctx, participants):
    pids, sections, car_ids = ctx["pids"], ctx["sections"], ctx["car_ids"]
    runs, drive, ride, usedcar, rent = ctx["runs"], ctx["drive"], ctx["ride"], ctx["usedcar"], ctx["rent"]
    is_mtn_car = ctx["is_mtn_car"]
    val = solver.Value

    sections_state: List[SectionState] = []
    for s_idx, s in enumerate(sections):
        runner_ids = [p for p in pids if (p, s) in runs and val(runs[(p, s)]) == 1]
        s_next = sections[s_idx + 1] if s_idx + 1 < len(sections) else None
        if s_next:
            next_runners = {p for p in pids if (p, s_next) in runs and val(runs[(p, s_next)]) == 1}
        else:
            next_runners = {p for p in pids if participants[p].preferred_sections[8] and _present(participants, p, 9)}

        cars = []
        for k in car_ids:
            if val(usedcar[(k, s)]) == 0:
                continue
            driver_id = next((p for p in pids if (p, k, s) in drive and val(drive[(p, k, s)]) == 1), "NO_DRIVER")
            passenger_ids = [p for p in pids if (p, k, s) in ride and val(ride[(p, k, s)]) == 1]
            car_people = {driver_id} | set(passenger_ids)
            is_adv = bool(car_people & next_runners)
            is_mtn = s in (7, 8) and (k, s) in is_mtn_car and val(is_mtn_car[(k, s)]) == 1
            cars.append(CarState(
                car_id=k, driver_id=driver_id, passenger_ids=passenger_ids,
                is_mountain_goer=is_mtn, is_advance=is_adv, car_type=CAR_TYPE[k], group=None,
            ))
        sections_state.append(SectionState(section_id=s, runner_ids=runner_ids, cars=cars))

    rent_solution = {k: val(rent[k]) for k in car_ids}
    runs_used_in_a = {p: 0 for p in pids}
    for (p, s), v in runs.items():
        runs_used_in_a[p] += val(v)
    ran_section_8 = {p for p in pids if (p, 8) in runs and val(runs[(p, 8)]) == 1}
    ran_section_7_or_8 = {
        p for p in pids
        if ((p, 7) in runs and val(runs[(p, 7)]) == 1) or ((p, 8) in runs and val(runs[(p, 8)]) == 1)
    }

    return sections_state, rent_solution, runs_used_in_a, ran_section_8, ran_section_7_or_8


# ---------------------------------------------------------------------------
# Block B (9〜10区) -- 元コードで既にO(n)の集約制約(mtn/mtn_car)だったのでそのまま移植
# ---------------------------------------------------------------------------

def _build_block_b(participants: Dict[str, Participant], rent_solution: Dict[str, int],
                    runs_used_in_a: Dict[str, int], ran_section_8: Optional[set] = None,
                    section8_mountain_drivers: Optional[Dict[str, str]] = None,
                    runner_limits: Optional[Dict[int, Tuple[int, Optional[int]]]] = None,
                    ran_section_7_or_8: Optional[set] = None):
    """section8_mountain_drivers: Block Aの8区で山フラグが立っていた車のcar_id -> 運転手id。
    山組は「1往復のみ」という前提(_build_block_bの元コメント参照)なので、8区で山行き車を
    運転していた人が9区でも同じ車を運転し続けるよう、Block B側にも引き継ぐ。"""
    runner_limits = runner_limits or {}
    pids = [p for p in participants if _present(participants, p, 9) and _present(participants, p, 10)]
    sections = BLOCK_B_SECTIONS
    rented_cars = [k for k in ALL_CAR_IDS if rent_solution.get(k, 0) == 1]
    mountain_capable = [p for p in pids if participants[p].can_drive_mountain]
    wants_mountain = any(
        participants[p].preferred_sections[8] or participants[p].preferred_sections[9] for p in pids
    )

    if wants_mountain and not mountain_capable:
        raise RuntimeError("山道運転可の参加者が見つかりません。9・10区の山行き車を運転できる人を少なくとも1人登録してください。")

    model = cp_model.CpModel()

    mtn = {p: model.NewBoolVar(f"mtn_{p}") for p in pids}
    mtn_car = {k: model.NewBoolVar(f"mtncar_{k}") for k in rented_cars}

    remaining_budget = {p: max(participants[p].remaining_sections - runs_used_in_a.get(p, 0), 0) for p in pids}

    ran_78 = ran_section_7_or_8 or set()

    runs = {}
    for p in pids:
        if remaining_budget[p] <= 0:
            continue
        if p in ran_78:
            # 7・8区を走った人は9・10区を走行不可(逆に9・10区を走る人は7・8区を走行不可、
            # という制約の対偶: Block Aが先に確定するのでここで直接弾く)
            continue
        for s in sections:
            if participants[p].preferred_sections[s - 1]:
                runs[(p, s)] = model.NewBoolVar(f"runsB_{p}_{s}")

    # 注意: 元コードはここで大型免許チェックが漏れており、山道免許はあるが大型免許は
    # ない人が大型のmtn_carを運転できてしまうバグがあった(validator.pyでは検出されるが
    # 割り当て自体は防げていなかった)。本実装ではBlock Aと同様にcan_drive_largeも見る。
    drive = {
        (p, k, s): model.NewBoolVar(f"driveB_{p}_{k}_{s}")
        for p in mountain_capable
        for k in rented_cars
        if not (CAR_TYPE[k] == "large" and not participants[p].can_drive_large)
        for s in sections
    }
    ride = {
        (p, k, s): model.NewBoolVar(f"rideB_{p}_{k}_{s}")
        for p in pids for k in rented_cars for s in sections
    }
    usedcar = {(k, s): model.NewBoolVar(f"usedB_{k}_{s}") for k in rented_cars for s in sections}

    # 8区→9区の山行き車引き継ぎ: Block Aで山フラグが立っていた車は9区でも山行き車のまま、
    # かつ運転手も同じ人に固定する(車両・運転手ともに8区からの継続を保証する)。
    # 運転手が日帰りで9区に残っていない場合(理論上あり得るが山組は基本宿泊者)は諦めて車の継続のみ試みる。
    if section8_mountain_drivers:
        for car_id, driver_id in section8_mountain_drivers.items():
            if car_id not in mtn_car:
                continue
            model.Add(mtn_car[car_id] == 1)
            if (driver_id, car_id, 9) in drive:
                model.Add(drive[(driver_id, car_id, 9)] == 1)

    for p in pids:
        run_vars = [v for (pp, s), v in runs.items() if pp == p]
        if run_vars:
            model.Add(sum(run_vars) <= remaining_budget[p])

    no_run_vars_b = []
    for p in pids:
        if runs_used_in_a.get(p, 0) > 0:
            continue
        if remaining_budget[p] <= 0:
            continue
        run_vars_p = [v for (pp, s), v in runs.items() if pp == p]
        if not run_vars_p:
            continue
        v_no_run = model.NewBoolVar(f"no_run_b_{p}")
        model.Add(sum(run_vars_p) + v_no_run >= 1)
        no_run_vars_b.append(v_no_run)

    for p in pids:
        for s in sections:
            terms = []
            if (p, s) in runs:
                terms.append(runs[(p, s)])
                model.Add(runs[(p, s)] <= mtn[p])
            terms += [drive[(p, k, s)] for k in rented_cars if (p, k, s) in drive]
            terms += [ride[(p, k, s)] for k in rented_cars if (p, k, s) in ride]
            model.Add(sum(terms) == mtn[p])

    b_no_passenger_vars, b_no_grade2_vars = [], []
    for k in rented_cars:
        for s in sections:
            drivers_ks = [drive[(p, k, s)] for p in mountain_capable if (p, k, s) in drive]
            model.Add(sum(drivers_ks) == usedcar[(k, s)])
            model.Add(usedcar[(k, s)] <= mtn_car[k])

            riders_ks = [ride[(p, k, s)] for p in pids if (p, k, s) in ride]
            model.Add(sum(riders_ks) <= (CAR_CAPACITY[k] - 1) * usedcar[(k, s)])
            v_pass = model.NewBoolVar(f"bno_pass_{k}_{s}")
            model.Add(v_pass >= usedcar[(k, s)] - sum(riders_ks))
            b_no_passenger_vars.append(v_pass)

            grade2_riders = [ride[(p, k, s)] for p in pids if participants[p].grade >= 2 and (p, k, s) in ride]
            v_g2 = model.NewBoolVar(f"bno_g2_{k}_{s}")
            model.Add(v_g2 >= usedcar[(k, s)] - sum(grade2_riders))
            b_no_grade2_vars.append(v_g2)

        for p in mountain_capable:
            if (p, k, 9) in drive and (p, k, 10) in drive:
                model.Add(drive[(p, k, 9)] == drive[(p, k, 10)])

    for s in sections:
        runners_s = [runs[(p, s)] for p in pids if (p, s) in runs]
        min_n, max_n = runner_limits.get(s, (1, None))
        if runners_s:
            model.Add(sum(runners_s) >= min_n)
            if max_n is not None:
                model.Add(sum(runners_s) <= max_n)
        elif min_n > 0:
            print(f"  ⚠️ {s}区: 走れる希望者がいないため、最少人数{min_n}人の指定は満たせません。")

    model.Add(sum(mtn[p] for p in pids) <= sum(CAR_CAPACITY[k] * mtn_car[k] for k in rented_cars))

    total_rented_cap = sum(CAR_CAPACITY[k] for k in rented_cars)
    model.Add(
        len(pids) - sum(mtn[p] for p in pids)
        <= total_rented_cap - sum(CAR_CAPACITY[k] * mtn_car[k] for k in rented_cars)
    )

    prev_run_drive_b_vars = []
    if ran_section_8:
        for p in mountain_capable:
            if p not in ran_section_8:
                continue
            for k in rented_cars:
                if (p, k, 9) not in drive:
                    continue
                v = model.NewBoolVar(f"prd_b_{p}_{k}")
                model.Add(v >= drive[(p, k, 9)])
                prev_run_drive_b_vars.append(v)

    # 8区のランナーは、9区で山フラグ付きの車(mtn_car)には回収されない(運転は可)。
    # Block A(1〜8区)は8区で終わるため、8区ランナーの「次区間での回収」はここBlock Bの
    # 9区が最初の機会になる。
    if ran_section_8:
        for p in ran_section_8:
            for k in rented_cars:
                if (p, k, 9) in ride:
                    model.Add(ride[(p, k, 9)] + mtn_car[k] <= 1)

    b_park_vars = []
    for k in rented_cars:
        for s in sections:
            v = model.NewBoolVar(f"bpark_{k}_{s}")
            model.Add(v >= mtn_car[k] - usedcar[(k, s)])
            b_park_vars.append(v)

    # --- ホテル組（山組に入らない人）の配車をここに統合する ---
    # 元コードは山グループ確定後に別モデルでホテル組を解いていたが、それだと
    # 「山組が大型免許持ちを使い切ってホテル組の大型車に運転手がいなくなる」
    # というケースを防げない(Block Bはmtn_car側の集計容量しか見ていないため)。
    # 同じモデル内でmtn[p]/mtn_car[k]と連動させることで、両立可能な解のみを選ばせる。
    hdrive = {
        (p, k): model.NewBoolVar(f"hdrive_{p}_{k}")
        for p in pids
        for k in rented_cars
        if participants[p].can_drive
        and not (CAR_TYPE[k] == "large" and not participants[p].can_drive_large)
        and p not in (ran_section_8 or set())
    }
    hride = {(p, k): model.NewBoolVar(f"hride_{p}_{k}") for p in pids for k in rented_cars}
    husedcar = {k: model.NewBoolVar(f"husedcar_{k}") for k in rented_cars}

    for p in pids:
        h_terms = [hdrive[(p, k)] for k in rented_cars if (p, k) in hdrive]
        h_terms += [hride[(p, k)] for k in rented_cars]
        # mtn[p]=1(山組) なら 0台、mtn[p]=0(ホテル組) なら必ず1台に乗る
        model.Add(sum(h_terms) == 1 - mtn[p])

    h_no_passenger_vars, h_no_grade2_vars = [], []
    for k in rented_cars:
        # 1台の車は「山組専用」か「ホテル組(非山組)専用」のどちらか一方
        model.Add(mtn_car[k] + husedcar[k] <= 1)

        hdrivers_k = [hdrive[(p, k)] for p in pids if (p, k) in hdrive]
        model.Add(sum(hdrivers_k) == husedcar[k])

        hriders_k = [hride[(p, k)] for p in pids]
        model.Add(sum(hriders_k) <= (CAR_CAPACITY[k] - 1) * husedcar[k])

        v_pass = model.NewBoolVar(f"hno_pass_{k}")
        model.Add(v_pass >= husedcar[k] - sum(hriders_k))
        h_no_passenger_vars.append(v_pass)

        grade2 = [hride[(p, k)] for p in pids if participants[p].grade >= 2]
        v_g2 = model.NewBoolVar(f"hno_g2_{k}")
        model.Add(v_g2 >= husedcar[k] - sum(grade2))
        h_no_grade2_vars.append(v_g2)

    # 「特に走りたい区間」を実際に走れた場合、通常の希望区間ボーナスに加えて追加ボーナスを与える
    priority_run_vars_b = [
        v for (p, s), v in runs.items()
        if participants[p].priority_sections[s - 1]
    ]

    # 0.1という小さい係数はCP-SATの整数目的関数と相性が悪いため10倍して整数化
    model.Minimize(
        1 * sum(mtn_car.values())
        - 10 * W_RUNNER_PREF * sum(runs.values())
        - 10 * W_PRIORITY_RUNNER_PREF * sum(priority_run_vars_b)
        + 10 * W_PREV_RUN_DRIVE * sum(prev_run_drive_b_vars)
        + 10 * W_NO_PASSENGER * sum(b_no_passenger_vars)
        + 10 * W_NO_GRADE2 * sum(b_no_grade2_vars)
        + 10 * W_PARK * sum(b_park_vars)
        + 10 * W_NO_RUN * sum(no_run_vars_b)
        - 10 * W_PARK * sum(husedcar.values())
        + 10 * W_NO_PASSENGER * sum(h_no_passenger_vars)
        + 10 * W_NO_GRADE2 * sum(h_no_grade2_vars)
    )

    ctx = dict(mtn=mtn, mtn_car=mtn_car, runs=runs, drive=drive, ride=ride, usedcar=usedcar,
               hdrive=hdrive, hride=hride, husedcar=husedcar,
               pids=pids, sections=sections, rented_cars=rented_cars)
    return model, ctx


def _extract_block_b(solver: cp_model.CpSolver, ctx, participants):
    pids, sections, rented_cars = ctx["pids"], ctx["sections"], ctx["rented_cars"]
    runs, drive, ride, usedcar, mtn, mtn_car = (
        ctx["runs"], ctx["drive"], ctx["ride"], ctx["usedcar"], ctx["mtn"], ctx["mtn_car"]
    )
    val = solver.Value

    mountain_group = [p for p in pids if val(mtn[p]) == 1]
    mountain_cars_used = [k for k in rented_cars if val(mtn_car[k]) == 1]
    total_capacity = sum(CAR_CAPACITY[k] for k in mountain_cars_used)
    names = ", ".join(participants[p].name for p in mountain_group)
    print(f"  ⛰️ 山グループ: {len(mountain_group)}名 / 山行き車定員合計: {total_capacity}名  ({names})")

    sections_state: List[SectionState] = []
    for s in sections:
        runner_ids = [p for p in pids if (p, s) in runs and val(runs[(p, s)]) == 1]
        cars = []
        for k in rented_cars:
            if val(usedcar[(k, s)]) == 0:
                continue
            driver_id = next((p for p in pids if (p, k, s) in drive and val(drive[(p, k, s)]) == 1), "NO_DRIVER")
            passenger_ids = [p for p in pids if (p, k, s) in ride and val(ride[(p, k, s)]) == 1]
            cars.append(CarState(car_id=k, driver_id=driver_id, passenger_ids=passenger_ids,
                                  is_mountain_goer=True, car_type=CAR_TYPE[k], group="mountain"))
        sections_state.append(SectionState(section_id=s, runner_ids=runner_ids, cars=cars))
    return sections_state


def _extract_hotel_group(solver: cp_model.CpSolver, ctx, participants) -> List[CarState]:
    """Block Bと同じモデルで同時に解いたホテル組(非山組)の配車を取り出す。"""
    pids, rented_cars = ctx["pids"], ctx["rented_cars"]
    hdrive, hride, husedcar = ctx["hdrive"], ctx["hride"], ctx["husedcar"]
    val = solver.Value

    cars = []
    for k in rented_cars:
        if val(husedcar[k]) == 0:
            continue
        driver_id = next((p for p in pids if (p, k) in hdrive and val(hdrive[(p, k)]) == 1), "NO_DRIVER")
        passenger_ids = [p for p in pids if val(hride[(p, k)]) == 1]
        cars.append(CarState(car_id=k, driver_id=driver_id, passenger_ids=passenger_ids,
                              is_mountain_goer=False, car_type=CAR_TYPE[k], group="hotel"))
    return cars


# ホテル組(非山組)は _build_block_b に統合済み (hdrive/hride/husedcar)。
# 単独の別モデルとしては解かない -- 詳細は _build_block_b 内のコメント参照。


# ---------------------------------------------------------------------------
# 帰路
# ---------------------------------------------------------------------------

def _build_return_trip(participants: Dict[str, Participant], rent_solution: Dict[str, int],
                        hotel_cars: Optional[List[CarState]] = None):
    pids = [p for p in participants if _needs_return_trip(participants, p)]
    rented_cars = [k for k in ALL_CAR_IDS if rent_solution.get(k, 0) == 1]

    model = cp_model.CpModel()
    drive = {
        (p, k): model.NewBoolVar(f"driveR_{p}_{k}")
        for p in pids if participants[p].can_drive
        for k in rented_cars
        if not (CAR_TYPE[k] == "large" and not participants[p].can_drive_large)
    }
    ride = {(p, k): model.NewBoolVar(f"rideR_{p}_{k}") for p in pids for k in rented_cars}
    usedcar = {k: model.NewBoolVar(f"usedR_{k}") for k in rented_cars}

    for p in pids:
        terms = [drive[(p, k)] for k in rented_cars if (p, k) in drive]
        terms += [ride[(p, k)] for k in rented_cars]
        model.Add(sum(terms) == 1)

    # ホテル組は10区の車構成(運転手・同乗者)をそのまま帰路にも引き継ぐ(ハード制約)。
    # 運転手が帰路の対象外(稀なケース)の場合は同乗者の継続だけ強制し、運転手は他の人から選ばれる。
    # さらに9→10区と同じ「車の専属」制約(mtn_car[k]+husedcar[k]<=1 相当)も帰路に適用し、
    # 元のホテル組メンバー以外がその車に混ざらないようにする(運転・同乗ともに禁止)。
    for car in (hotel_cars or []):
        k = car.car_id
        members = {car.driver_id} | set(car.passenger_ids)

        if (car.driver_id, k) in drive:
            model.Add(drive[(car.driver_id, k)] == 1)
        for p in car.passenger_ids:
            if (p, k) in ride:
                model.Add(ride[(p, k)] == 1)

        for p in pids:
            if p in members:
                continue
            if (p, k) in ride:
                model.Add(ride[(p, k)] == 0)
            if (p, k) in drive:
                model.Add(drive[(p, k)] == 0)

    r_no_passenger_vars, r_no_grade2_vars = [], []
    for k in rented_cars:
        drivers_k = [drive[(p, k)] for p in pids if (p, k) in drive]
        model.Add(sum(drivers_k) == usedcar[k])

        riders_k = [ride[(p, k)] for p in pids]
        model.Add(sum(riders_k) <= (CAR_CAPACITY[k] - 1) * usedcar[k])
        v_pass = model.NewBoolVar(f"rno_pass_{k}")
        model.Add(v_pass >= usedcar[k] - sum(riders_k))
        r_no_passenger_vars.append(v_pass)

        grade2_riders = [ride[(p, k)] for p in pids if participants[p].grade >= 2]
        v_g2 = model.NewBoolVar(f"rno_g2_{k}")
        model.Add(v_g2 >= usedcar[k] - sum(grade2_riders))
        r_no_grade2_vars.append(v_g2)

    model.Minimize(
        -W_PARK * sum(usedcar.values())
        + W_NO_PASSENGER * sum(r_no_passenger_vars)
        + W_NO_GRADE2 * sum(r_no_grade2_vars)
    )

    ctx = dict(drive=drive, ride=ride, usedcar=usedcar, pids=pids, rented_cars=rented_cars)
    return model, ctx


def _extract_return_trip(solver: cp_model.CpSolver, ctx, participants):
    pids, rented_cars = ctx["pids"], ctx["rented_cars"]
    drive, ride, usedcar = ctx["drive"], ctx["ride"], ctx["usedcar"]
    val = solver.Value

    cars = []
    for k in rented_cars:
        if val(usedcar[k]) == 0:
            continue
        driver_id = next((p for p in pids if (p, k) in drive and val(drive[(p, k)]) == 1), "NO_DRIVER")
        passenger_ids = [p for p in pids if val(ride[(p, k)]) == 1]
        cars.append(CarState(car_id=k, driver_id=driver_id, passenger_ids=passenger_ids,
                              is_mountain_goer=False, car_type=CAR_TYPE[k], group="return"))
    return SectionState(section_id=RETURN_TRIP_SECTION_ID, runner_ids=[], cars=cars)


def _renumber_cars(plan: List[SectionState], rent_solution: Dict[str, int]) -> None:
    used = [k for k in ALL_CAR_IDS if rent_solution.get(k, 0) == 1]
    large_used = [k for k in used if CAR_TYPE[k] == "large"]
    normal_used = [k for k in used if CAR_TYPE[k] == "normal"]
    mapping = {k: f"L{i}" for i, k in enumerate(large_used, start=1)}
    mapping.update({k: f"N{i}" for i, k in enumerate(normal_used, start=1)})
    for section in plan:
        for car in section.cars:
            car.car_id = mapping.get(car.car_id, car.car_id)


# ---------------------------------------------------------------------------
# 全体オーケストレーション
# ---------------------------------------------------------------------------

def generate_full_plan_cpsat(
    participants: Dict[str, Participant],
    active_car_ids: Optional[List[str]] = None,
    time_limit: float = DEFAULT_TIME_LIMIT,
    workers: int = DEFAULT_WORKERS,
    output_path: Optional[str] = "hakone_result.xlsx",
    runner_limits: Optional[Dict[int, Tuple[int, Optional[int]]]] = None,
) -> List[SectionState]:
    """runner_limits: 区間番号(1〜10) -> (最少人数, 最大人数)。最大人数はNoneで無制限。
    指定のない区間は既定値(最少1人・無制限)になる。"""
    if active_car_ids is None:
        active_car_ids = ALL_CAR_IDS
    n_large_slots = sum(1 for k in active_car_ids if CAR_TYPE[k] == "large")
    n_normal_slots = sum(1 for k in active_car_ids if CAR_TYPE[k] == "normal")
    mode_label = f"大型{n_large_slots}台スロット + 普通{n_normal_slots}台スロット"

    n_drive = sum(1 for p in participants.values() if p.can_drive)
    n_large = sum(1 for p in participants.values() if p.can_drive_large)
    n_mtn = sum(1 for p in participants.values() if p.can_drive_mountain)
    n_stay = sum(1 for p in participants if _needs_return_trip(participants, p))
    total_pref = sum(sum(p.preferred_sections) for p in participants.values())
    print(f"🚗 車種モード: {mode_label}")
    print(f"📋 参加者データ: 計{len(participants)}人 / 運転可={n_drive} / 大型可={n_large} / 山道可={n_mtn} / 宿泊(帰路対象)={n_stay} / 希望延べ区間数={total_pref}")

    t0 = time.time()
    model_a, ctx_a = _build_block_a(participants, car_ids=active_car_ids, runner_limits=runner_limits)
    solver_a, status_a = _solve(model_a, time_limit, workers)
    print(f"Block A (1〜8区) 最適化ステータス: {status_a}  ({time.time()-t0:.1f}秒)")
    if status_a not in ("OPTIMAL", "FEASIBLE"):
        raise RuntimeError(f"Block A(1〜8区)が解けませんでした: {status_a}")
    if status_a == "FEASIBLE":
        print("  ⚠️ 制限時間内に最適性は証明できませんでしたが、見つかった解を使用します。")

    sections_a, rent_solution, runs_used_in_a, ran_section_8, ran_section_7_or_8 = _extract_block_a(
        solver_a, ctx_a, participants)
    if ran_section_8:
        names_s8 = ", ".join(participants[p].name for p in ran_section_8)
        print(f"  🏃 8区走者（次区間の運転除外）: {names_s8}")
    if ran_section_7_or_8:
        names_s78 = ", ".join(participants[p].name for p in ran_section_7_or_8)
        print(f"  🚫 7・8区走者（9・10区の走行不可）: {names_s78}")

    section8 = next((s for s in sections_a if s.section_id == 8), None)
    section8_mountain_drivers = {
        car.car_id: car.driver_id
        for car in (section8.cars if section8 else [])
        if car.is_mountain_goer and car.driver_id != "NO_DRIVER"
    }
    if section8_mountain_drivers:
        names = ", ".join(f"{k}:{participants[d].name}" for k, d in section8_mountain_drivers.items())
        print(f"  🔗 8区時点の山行き車と運転手（9区に引き継ぐ）: {names}")

    t0 = time.time()
    model_b, ctx_b = _build_block_b(participants, rent_solution, runs_used_in_a, ran_section_8,
                                     section8_mountain_drivers, runner_limits=runner_limits,
                                     ran_section_7_or_8=ran_section_7_or_8)
    solver_b, status_b = _solve(model_b, time_limit, workers)
    print(f"Block B (9〜10区) 最適化ステータス: {status_b}  ({time.time()-t0:.1f}秒)")
    if status_b not in ("OPTIMAL", "FEASIBLE"):
        raise RuntimeError(f"Block B(9〜10区)が解けませんでした: {status_b}")
    if status_b == "FEASIBLE":
        print("  ⚠️ 制限時間内に最適性は証明できませんでしたが、見つかった解を使用します。")

    sections_b = _extract_block_b(solver_b, ctx_b, participants)

    mountain_group_b = {p for p in ctx_b["pids"] if solver_b.Value(ctx_b["mtn"][p]) == 1}
    names_b = ", ".join(participants[p].name for p in mountain_group_b)
    print(f"  🏔️ Block B 確定山グループ: {names_b if names_b else '（なし）'}")
    # 7・8区の表示上の山フラグは、実際に座席制約(7区/8区ランナーは山フラグ車に回収されない等)
    # で使われたBlock A自身の is_mtn_car の判定をそのまま使う(_extract_block_a で既に
    # is_mountain_goer にセット済みなので、ここでは上書きしない)。
    # 以前はBlock B確定の山グループで事後的に上書きしていたが、それだと実際には
    # is_mtn_car=0(=山フラグの立っていない車)に7区/8区ランナーが同乗しているのに、
    # 表示上だけ山フラグが立って見える不整合が生じていたため。

    hotel_cars = _extract_hotel_group(solver_b, ctx_b, participants)
    if hotel_cars:
        hotel_total = sum(c.total_people for c in hotel_cars)
        names = ", ".join(
            participants[p].name for c in hotel_cars for p in [c.driver_id] + c.passenger_ids if p in participants
        )
        print(f"  🏨 ホテルグループ: {hotel_total}名 (車{len(hotel_cars)}台)  ({names})")
        for sec in sections_b:
            sec.cars.extend(hotel_cars)

    t0 = time.time()
    model_c, ctx_c = _build_return_trip(participants, rent_solution, hotel_cars=hotel_cars)
    solver_c, status_c = _solve(model_c, time_limit, workers)
    print(f"帰路 最適化ステータス: {status_c}  ({time.time()-t0:.1f}秒)")
    if status_c not in ("OPTIMAL", "FEASIBLE"):
        raise RuntimeError(f"帰路の割り当てが解けませんでした: {status_c}")
    if status_c == "FEASIBLE":
        print("  ⚠️ 制限時間内に最適性は証明できませんでしたが、見つかった解を使用します。")
    section_c = _extract_return_trip(solver_c, ctx_c, participants)

    plan = sections_a + sections_b + [section_c]
    _renumber_cars(plan, rent_solution)

    for section in plan:
        errors = validate_section(section, participants)
        label = section_label(section.section_id)
        if errors:
            print(f"❌ {label}でエラー: " + " / ".join(errors))
        else:
            print(f"✅ {label}は問題なく割り当てられました！")

    used = [k for k, v in rent_solution.items() if v == 1]
    large_n = sum(1 for k in used if CAR_TYPE[k] == "large")
    normal_n = sum(1 for k in used if CAR_TYPE[k] == "normal")
    print(f"\n🚗 レンタルした車: 大型{large_n}台 + 普通{normal_n}台 = 合計{len(used)}台")

    if output_path:
        write_plan_xlsx(plan, participants, output_path)
    return plan
