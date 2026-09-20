"""裏シュミ: 表シュミ(generate_full_plan_cpsat)の結果を土台に、個人の希望
(「この人とこの人はできるだけ一緒にしたい」「離したい」「運転区間数を範囲に
収める」)を追加のソフト/ハード制約として加え、既存の割り当てからの変更を
最小限に抑えながら再最適化する。

表シュミ本体は変えず、Block A/Block B/帰路の3フェーズ構成をそのまま踏襲する。
各フェーズで以下を行う:
  1) 表シュミの結果をヒントとして与える(AddHint) — 探索の出発点にする
  2) 「一緒にしたい/離したい」を目的関数へのソフト項として追加
     (milp_allocator_v3側のextra_objective_fnフック経由。同じCpModelインスタンス上で
     補助変数を作る必要があるため、「先に項を作ってから後で目的関数に足す」ことは
     できない — フックはmodel.Minimize()を呼ぶ直前に、そのモデル自身を渡して呼ばれる)
  3) 「運転区間数の範囲」をハード制約として追加
     (上限は各フェーズで独立に、累積で超えないようチェックする。
      下限は運転の機会が残っている最後のフェーズでのみチェックする
      — フェーズごとに独立モデルなので「後のフェーズで帳尻を合わせる」判断は
      前のフェーズの時点ではできないため)
  4) 表シュミの結果から変えたことへのペナルティを目的関数に追加
     (新条件を満たすために必要な変更だけに絞られるように誘導する)

表シュミ自体が「最適性を証明できないまま打ち切ることがある」性質をそのまま
引き継ぐため、裏シュミの結果も同様に「完全な最適」は保証しない(ユーザーの
想定通り: 元が最適に近い解であれば、そこからの小さな変更で条件を足しても
十分近い解になるはず、という前提に立っている)。
"""

import os
import sys
from typing import Dict, List, Optional, Tuple

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Participant, SectionState
from logic.car_pool import section_label
from logic.milp_allocator_v3 import (
    _build_block_a, _build_block_b, _build_return_trip,
    _extract_block_a, _extract_block_b, _extract_hotel_group, _extract_return_trip,
    _solve, _renumber_cars, _present, _needs_return_trip,
    DEFAULT_TIME_LIMIT, DEFAULT_WORKERS, W_PRIORITY_RUNNER_PREF,
)
from validator import validate_section

# 裏シュミ専用の重み。表シュミの結果はあくまで土台であり、裏シュミの新条件は
# 「それに加えてこうだったら良い」という副次的な立ち位置(表シュミ側の重要な
# ソフト制約、特に「特に走りたい区間」(W_PRIORITY_RUNNER_PREF)は裏シュミの
# 新条件より優先されるべき、というユーザーとの合意に基づく)。そのため
# W_TOGETHER/W_APARTはW_PRIORITY_RUNNER_PREFより明確に小さい値にしている。
W_TOGETHER = W_PRIORITY_RUNNER_PREF - 50  # 「一緒にしたい」ペアが同じ車に乗れたときのボーナス(区間ごとに加算)
W_APART = W_PRIORITY_RUNNER_PREF - 50     # 「離したい」ペアが同じ車に乗ってしまったときの罰(区間ごとに加算)
W_STAY_CLOSE = 20  # 表シュミの結果から変えないことへの誘導(小さな変更で条件を満たす方向に働く)
# 運転区間数の下限はハード制約として最後の機会(_last_present_block)でチェックするが、
# それより前のフェーズには「下限に届いていない」という情報が伝わらないため、Block A/Bが
# 均等に運転を割り振ってしまい、最後のフェーズで帳尻を合わせようとしてもINFEASIBLEに
# なることがあった(実際に確認済み)。下限未達成の人にはどのフェーズでも運転を積極的に
# 割り振るよう、ハード制約とは別にソフトな誘導も加える。
W_DRIVER_MIN_PUSH = 1000
# 「特に走りたい区間」は表シュミ側の目的関数(W_PRIORITY_RUNNER_PREF)に既に組み込まれて
# いるが、それだけだと「表シュミ以降に入力データが編集されて新たに追加された優先区間」
# のような、土台の時点で未充足だったものを裏シュミが確実に拾って直しにいく保証がない
# (元の割り当てへのヒントが探索を古い解に引っ張るため、時間制限内では変化が起きない
# ことがある)。そこで、ユーザーが指定した条件の有無に関わらず常に「土台の時点で未充足の
# 優先区間」を検出し、それをこの重みで強く後押しする。
# ユーザーとの合意により、この修正は「裏シュミの新条件より優先」どころか「表シュミ側の
# 他のソフト制約(台数最小化・駐車ペナルティ等)よりも最優先」で行う。そのため、表シュミ・
# 裏シュミを通じて最大のソフト制約重み(W_MTN_FLEET=50000)よりも明確に大きい値にしている
# (ハード制約である「希望区間数」の上限や、そもそも「走りたい区間」に入っていない区間は
# ソフト項がどれだけ大きくても破れない/効かないので、これらは自動的に対象外になる)。
W_PRIORITY_FIX_PUSH = 200000


class BackSimConfig:
    def __init__(
        self,
        together_pairs: Optional[List[Tuple[str, str]]] = None,
        apart_pairs: Optional[List[Tuple[str, str]]] = None,
        driver_ranges: Optional[Dict[str, Tuple[int, Optional[int]]]] = None,
        together_or_groups: Optional[List[Tuple[str, List[str]]]] = None,
    ):
        """together_pairs/apart_pairs: 参加者名(表記ゆれ吸収済みで突き合わせる)のペアのリスト。
        together_or_groups: (基準となる人, 候補者名のリスト) のリスト。「候補者のうち誰か一人とでも
        同じ車になれれば満足」というOR条件(together_pairsは各ペアが独立に評価される、いわば
        「それぞれと一緒になれるほど良い」というAND寄りの条件であるのに対し、こちらは
        「一人と一緒になれた時点でその区間は満足」という条件)。
        driver_ranges: 参加者名 -> (最少運転区間数, 最大運転区間数)。最大はNoneで無制限。"""
        self.together_pairs = together_pairs or []
        self.apart_pairs = apart_pairs or []
        self.driver_ranges = driver_ranges or {}
        self.together_or_groups = together_or_groups or []


def _resolve_pairs_and_ranges(config: BackSimConfig, participants: Dict[str, Participant]):
    """名前(表記ゆれ吸収)をparticipant_idに変換する。見つからない名前は警告して無視する。"""
    from data_io.plan_importer import _name_to_id_map, _normalize_name

    name_to_id = _name_to_id_map(participants)

    def resolve(name):
        pid = name_to_id.get(_normalize_name(name))
        if pid is None:
            print(f"  ⚠️ 裏シュミ設定: 参加者に見つからない名前をスキップします: {name}")
        return pid

    together_ids = []
    for a, b in config.together_pairs:
        pa, pb = resolve(a), resolve(b)
        if pa and pb:
            together_ids.append((pa, pb))

    apart_ids = []
    for a, b in config.apart_pairs:
        pa, pb = resolve(a), resolve(b)
        if pa and pb:
            apart_ids.append((pa, pb))

    driver_ranges_ids = {}
    for name, rng in config.driver_ranges.items():
        pid = resolve(name)
        if pid:
            driver_ranges_ids[pid] = rng

    together_or_group_ids = []
    for anchor, group in config.together_or_groups:
        pa = resolve(anchor)
        group_ids = [pid for pid in (resolve(name) for name in group) if pid]
        if pa and group_ids:
            together_or_group_ids.append((pa, group_ids))

    return together_ids, apart_ids, driver_ranges_ids, together_or_group_ids


def _baseline_occupancy(plan: List[SectionState]):
    """(p, car_id, section_id) の集合(運転手・同乗者どちらも含む)と、
    区間別のランナー集合を表シュミの結果から作る。"""
    occ = set()
    runner_by_section: Dict[int, set] = {}
    for section in plan:
        runner_by_section[section.section_id] = set(section.runner_ids)
        for car in section.cars:
            if car.driver_id and car.driver_id != "NO_DRIVER":
                occ.add((car.driver_id, car.car_id, section.section_id))
            for pid in car.passenger_ids:
                occ.add((pid, car.car_id, section.section_id))
    return occ, runner_by_section


def check_priority_sections(plan: List[SectionState], participants: Dict[str, Participant]):
    """「特に走りたい区間」ごとに、planで実際にその人がその区間を走れているかを調べる。
    表シュミ・裏シュミどちらの結果に対しても使える(裏シュミでは、実行前の土台に対して
    「直すべき未充足が無いか」を調べるのにも、実行後の結果に対して「直せたか」を
    確認するのにも使う)。
    戻り値: (詳細のリスト[{"name","section_id","satisfied"}], 未充足の(participant_id, section_id)集合)
    """
    _, runner_by_section = _baseline_occupancy(plan)
    details = []
    unmet = set()
    for pid, p in participants.items():
        for s in range(1, 11):
            if not p.priority_sections[s - 1]:
                continue
            satisfied = pid in runner_by_section.get(s, set())
            details.append({"name": p.name, "section_id": s, "satisfied": satisfied})
            if not satisfied:
                unmet.add((pid, s))
    return details, unmet


def _apply_hints(model, var_dict: dict, baseline_keys: set) -> None:
    for key, var in var_dict.items():
        model.AddHint(var, 1 if key in baseline_keys else 0)


def _together_apart_terms(model, pair_ids, occ_lookup, car_ids, sections, weight, want_together: bool):
    """occ_lookup(p, k, s) -> BoolVar/LinearExpr/None。
    want_together=Trueなら「同じ車になれた」ことへのボーナス(目的関数にマイナス)、
    Falseなら「同じ車になってしまった」ことへの罰(目的関数にプラス)を返す。"""
    terms = []
    for pa, pb in pair_ids:
        for s in sections:
            same_k_vars = []
            for k in car_ids:
                occ_a = occ_lookup(pa, k, s)
                occ_b = occ_lookup(pb, k, s)
                if occ_a is None or occ_b is None:
                    continue
                same_k = model.NewBoolVar(f"same_{pa}_{pb}_{k}_{s}")
                model.Add(same_k <= occ_a)
                model.Add(same_k <= occ_b)
                model.Add(same_k >= occ_a + occ_b - 1)
                same_k_vars.append(same_k)
            if not same_k_vars:
                continue
            together_s = model.NewBoolVar(f"together_{pa}_{pb}_{s}")
            for v in same_k_vars:
                model.Add(together_s >= v)
            model.Add(together_s <= sum(same_k_vars))
            terms.append(-weight * together_s if want_together else weight * together_s)
    return terms


def _or_together_terms(model, or_group_ids, occ_lookup, car_ids, sections, weight):
    """or_group_ids: (基準となる人, 候補者idのリスト) のリスト。
    区間ごとに「基準の人が候補者のうち誰か一人とでも同じ車になれたか」を表す
    単一のBoolVarを作り、それに対してのみボーナスを与える(候補者を何人も同時に
    連れて行っても、一人の時と同じ1回分のボーナスにしかならない = OR条件)。"""
    terms = []
    for anchor, group in or_group_ids:
        for s in sections:
            same_vars = []
            for k in car_ids:
                occ_a = occ_lookup(anchor, k, s)
                if occ_a is None:
                    continue
                for g in group:
                    occ_g = occ_lookup(g, k, s)
                    if occ_g is None:
                        continue
                    same_k_g = model.NewBoolVar(f"same_or_{anchor}_{g}_{k}_{s}")
                    model.Add(same_k_g <= occ_a)
                    model.Add(same_k_g <= occ_g)
                    model.Add(same_k_g >= occ_a + occ_g - 1)
                    same_vars.append(same_k_g)
            if not same_vars:
                continue
            satisfied_s = model.NewBoolVar(f"or_satisfied_{anchor}_{s}")
            for v in same_vars:
                model.Add(satisfied_s >= v)
            model.Add(satisfied_s <= sum(same_vars))
            terms.append(-weight * satisfied_s)
    return terms


def _stay_close_terms(var_dict: dict, baseline_keys: set, weight) -> list:
    terms = []
    for key, var in var_dict.items():
        terms.append(-weight * var if key in baseline_keys else weight * var)
    return terms


def _last_present_block(participants: Dict[str, Participant], p: str) -> str:
    """運転区間数の下限を最終チェックすべきフェーズを、参加区間の状況からおおまかに判定する
    ("A"/"B"/"return")。この人がその区間で実際に運転可能か(免許等)までは見ない
    — 満たせない場合はCP-SATがそのフェーズでINFEASIBLEを返すので、それをそのまま
    ユーザーに伝える(下限が構造的に満たせないことを意味するので、これ自体は正しい)。"""
    if _needs_return_trip(participants, p):
        return "return"
    if _present(participants, p, 9) and _present(participants, p, 10):
        return "B"
    return "A"


def run_back_sim(
    baseline_plan: List[SectionState],
    participants: Dict[str, Participant],
    config: BackSimConfig,
    time_limit: float = DEFAULT_TIME_LIMIT,
    workers: int = DEFAULT_WORKERS,
) -> List[SectionState]:
    together_ids, apart_ids, driver_ranges, together_or_group_ids = _resolve_pairs_and_ranges(config, participants)

    active_car_ids = sorted({car.car_id for section in baseline_plan for car in section.cars})

    baseline_occ, baseline_runners = _baseline_occupancy(baseline_plan)
    baseline_occ_no_section = {(p, k) for (p, k, s) in baseline_occ}
    baseline_run_keys = {(p, s) for s, names in baseline_runners.items() for p in names}
    last_block_for = {p: _last_present_block(participants, p) for p in driver_ranges}
    driven_so_far: Dict[str, int] = {p: 0 for p in driver_ranges}

    _, unmet_priority_keys = check_priority_sections(baseline_plan, participants)
    if unmet_priority_keys:
        names = ", ".join(f"{participants[p].name}:{s}区" for p, s in sorted(unmet_priority_keys, key=lambda x: x[1]))
        print(f"[裏シュミ] 土台の時点で未充足の「特に走りたい区間」を検出、優先的に直します: {names}")

    def priority_fix_terms(runs: Dict[Tuple[str, int], object]) -> list:
        return [-W_PRIORITY_FIX_PUSH * v for (p, s), v in runs.items() if (p, s) in unmet_priority_keys]

    # 9・10区は「7・8区を走った人は9・10区を走行不可」というBlock A→Block Bをまたぐ
    # ハード制約があるため(milp_allocator_v3._build_block_b参照)、Block B側だけで
    # 9・10区の優先度を後押ししても、その人がBlock Aの時点で既に7・8区の担当に
    # 決まっていれば手遅れになる。裏シュミではこれも直したいので、9・10区に未充足の
    # 優先区間がある人については、Block A側でも7・8区の担当を避けるよう誘導する
    # (もちろんソフトな誘導であり、他のハード制約は一切変えない)。
    priority_910_people = {p for (p, s) in unmet_priority_keys if s in (9, 10)}

    def priority_avoid_78_terms(runs_a: Dict[Tuple[str, int], object]) -> list:
        return [
            W_PRIORITY_FIX_PUSH * v
            for (p, s), v in runs_a.items()
            if p in priority_910_people and s in (7, 8)
        ]

    def apply_driver_range_constraints(model, drive_terms_by_person: Dict[str, list], block_name: str):
        for p, terms_p in drive_terms_by_person.items():
            if p not in driver_ranges or not terms_p:
                continue
            min_r, max_r = driver_ranges[p]
            already = driven_so_far.get(p, 0)
            if max_r is not None:
                model.Add(sum(terms_p) <= max(max_r - already, 0))
            if last_block_for[p] == block_name:
                model.Add(already + sum(terms_p) >= min_r)

    def driver_min_push_terms(drive_terms_by_person: Dict[str, list]) -> list:
        """下限に届いていない人には、このフェーズでも運転を割り振るよう誘導するボーナス。"""
        terms = []
        for p, terms_p in drive_terms_by_person.items():
            if p not in driver_ranges or not terms_p:
                continue
            min_r, _max_r = driver_ranges[p]
            if driven_so_far.get(p, 0) >= min_r:
                continue
            terms.append(-W_DRIVER_MIN_PUSH * sum(terms_p))
        return terms

    # ------------------------------------------------------------------
    # Block A (1〜8区)
    # ------------------------------------------------------------------
    def extra_fn_a(model, ctx):
        occ = ctx["occ"]
        terms = []
        terms += _together_apart_terms(model, together_ids, lambda p, k, s: occ.get((p, k, s)),
                                        ctx["car_ids"], ctx["sections"], W_TOGETHER, True)
        terms += _together_apart_terms(model, apart_ids, lambda p, k, s: occ.get((p, k, s)),
                                        ctx["car_ids"], ctx["sections"], W_APART, False)
        terms += _or_together_terms(model, together_or_group_ids, lambda p, k, s: occ.get((p, k, s)),
                                     ctx["car_ids"], ctx["sections"], W_TOGETHER)
        terms += _stay_close_terms(ctx["drive"], baseline_occ, W_STAY_CLOSE)
        terms += _stay_close_terms(ctx["ride"], baseline_occ, W_STAY_CLOSE)
        terms += priority_fix_terms(ctx["runs"])
        terms += priority_avoid_78_terms(ctx["runs"])

        drive_terms_by_person: Dict[str, list] = {}
        for (p, k, s), v in ctx["drive"].items():
            drive_terms_by_person.setdefault(p, []).append(v)
        apply_driver_range_constraints(model, drive_terms_by_person, "A")
        terms += driver_min_push_terms(drive_terms_by_person)
        return terms

    model_a, ctx_a = _build_block_a(participants, car_ids=active_car_ids, extra_objective_fn=extra_fn_a)
    _apply_hints(model_a, ctx_a["drive"], baseline_occ)
    _apply_hints(model_a, ctx_a["ride"], baseline_occ)
    _apply_hints(model_a, ctx_a["runs"], baseline_run_keys)

    solver_a, status_a = _solve(model_a, time_limit, workers)
    print(f"[裏シュミ] Block A (1〜8区) 最適化ステータス: {status_a}")
    if status_a not in ("OPTIMAL", "FEASIBLE"):
        raise RuntimeError(
            f"Block A(1〜8区)が指定された条件で解けませんでした: {status_a}\n"
            "運転区間数の下限が厳しすぎる、または「離したい」が多すぎる可能性があります。"
        )

    sections_a, rent_solution_a, runs_used_in_a, ran_section_8, ran_section_7_or_8 = _extract_block_a(
        solver_a, ctx_a, participants)

    for p in driver_ranges:
        driven_so_far[p] += sum(solver_a.Value(v) for (pp, k, s), v in ctx_a["drive"].items() if pp == p)

    section8 = next((s for s in sections_a if s.section_id == 8), None)
    section8_mountain_drivers = {
        car.car_id: car.driver_id
        for car in (section8.cars if section8 else [])
        if car.is_mountain_goer and car.driver_id != "NO_DRIVER"
    }

    # ------------------------------------------------------------------
    # Block B (9〜10区、ホテル組込み)
    # ------------------------------------------------------------------
    def extra_fn_b(model, ctx):
        drive_b, ride_b, hdrive_b, hride_b = ctx["drive"], ctx["ride"], ctx["hdrive"], ctx["hride"]

        def occ_b(p, k, s):
            m_d, m_r = drive_b.get((p, k, s)), ride_b.get((p, k, s))
            h_d, h_r = hdrive_b.get((p, k)), hride_b.get((p, k))
            parts = [v for v in (m_d, m_r, h_d, h_r) if v is not None]
            return sum(parts) if parts else None

        terms = []
        terms += _together_apart_terms(model, together_ids, occ_b, ctx["rented_cars"], ctx["sections"],
                                        W_TOGETHER, True)
        terms += _together_apart_terms(model, apart_ids, occ_b, ctx["rented_cars"], ctx["sections"],
                                        W_APART, False)
        terms += _or_together_terms(model, together_or_group_ids, occ_b, ctx["rented_cars"], ctx["sections"],
                                     W_TOGETHER)
        terms += _stay_close_terms(drive_b, baseline_occ, W_STAY_CLOSE)
        terms += _stay_close_terms(ride_b, baseline_occ, W_STAY_CLOSE)
        terms += _stay_close_terms(hdrive_b, baseline_occ_no_section, W_STAY_CLOSE)
        terms += _stay_close_terms(hride_b, baseline_occ_no_section, W_STAY_CLOSE)
        terms += priority_fix_terms(ctx["runs"])

        drive_terms_by_person: Dict[str, list] = {}
        for (p, k, s), v in drive_b.items():
            drive_terms_by_person.setdefault(p, []).append(v)
        for (p, k), v in hdrive_b.items():
            drive_terms_by_person.setdefault(p, []).append(v)
        apply_driver_range_constraints(model, drive_terms_by_person, "B")
        terms += driver_min_push_terms(drive_terms_by_person)
        return terms

    model_b, ctx_b = _build_block_b(
        participants, rent_solution_a, runs_used_in_a, ran_section_8,
        section8_mountain_drivers, ran_section_7_or_8=ran_section_7_or_8,
        extra_objective_fn=extra_fn_b,
    )
    _apply_hints(model_b, ctx_b["drive"], baseline_occ)
    _apply_hints(model_b, ctx_b["ride"], baseline_occ)
    _apply_hints(model_b, ctx_b["hdrive"], baseline_occ_no_section)
    _apply_hints(model_b, ctx_b["hride"], baseline_occ_no_section)
    _apply_hints(model_b, ctx_b["runs"], baseline_run_keys)

    solver_b, status_b = _solve(model_b, time_limit, workers)
    print(f"[裏シュミ] Block B (9〜10区) 最適化ステータス: {status_b}")
    if status_b not in ("OPTIMAL", "FEASIBLE"):
        raise RuntimeError(
            f"Block B(9〜10区)が指定された条件で解けませんでした: {status_b}\n"
            "運転区間数の下限が厳しすぎる、または「離したい」が多すぎる可能性があります。"
        )

    sections_b = _extract_block_b(solver_b, ctx_b, participants)
    hotel_cars = _extract_hotel_group(solver_b, ctx_b, participants)
    for sec in sections_b:
        sec.cars.extend(hotel_cars)

    for p in driver_ranges:
        driven_so_far[p] += sum(solver_b.Value(v) for (pp, k, s), v in ctx_b["drive"].items() if pp == p)
        driven_so_far[p] += sum(solver_b.Value(v) for (pp, k), v in ctx_b["hdrive"].items() if pp == p)

    # ------------------------------------------------------------------
    # 帰路
    # ------------------------------------------------------------------
    def extra_fn_c(model, ctx):
        drive_c, ride_c = ctx["drive"], ctx["ride"]

        def occ_c(p, k, _s):
            d, r = drive_c.get((p, k)), ride_c.get((p, k))
            parts = [v for v in (d, r) if v is not None]
            return sum(parts) if parts else None

        terms = []
        terms += _together_apart_terms(model, together_ids, occ_c, ctx["rented_cars"], [None],
                                        W_TOGETHER, True)
        terms += _together_apart_terms(model, apart_ids, occ_c, ctx["rented_cars"], [None],
                                        W_APART, False)
        terms += _or_together_terms(model, together_or_group_ids, occ_c, ctx["rented_cars"], [None],
                                     W_TOGETHER)
        terms += _stay_close_terms(drive_c, baseline_occ_no_section, W_STAY_CLOSE)
        terms += _stay_close_terms(ride_c, baseline_occ_no_section, W_STAY_CLOSE)

        drive_terms_by_person: Dict[str, list] = {}
        for (p, k), v in drive_c.items():
            drive_terms_by_person.setdefault(p, []).append(v)
        for p in driver_ranges:
            if last_block_for[p] != "return":
                continue
            min_r, _max_r = driver_ranges[p]
            if p not in drive_terms_by_person and driven_so_far.get(p, 0) < min_r:
                raise RuntimeError(
                    f"{participants[p].name}さんは帰路で運転する機会がなく、"
                    f"指定された運転区間数の下限({min_r})に届きません。"
                )
        apply_driver_range_constraints(model, drive_terms_by_person, "return")
        return terms

    model_c, ctx_c = _build_return_trip(participants, rent_solution_a, hotel_cars=hotel_cars,
                                         extra_objective_fn=extra_fn_c)
    _apply_hints(model_c, ctx_c["drive"], baseline_occ_no_section)
    _apply_hints(model_c, ctx_c["ride"], baseline_occ_no_section)

    solver_c, status_c = _solve(model_c, time_limit, workers)
    print(f"[裏シュミ] 帰路 最適化ステータス: {status_c}")
    if status_c not in ("OPTIMAL", "FEASIBLE"):
        raise RuntimeError(
            f"帰路が指定された条件で解けませんでした: {status_c}\n"
            "運転区間数の下限が厳しすぎる、または「離したい」が多すぎる可能性があります。"
        )
    section_c = _extract_return_trip(solver_c, ctx_c, participants)

    plan = sections_a + sections_b + [section_c]
    _renumber_cars(plan, rent_solution_a)

    for section in plan:
        errors = validate_section(section, participants)
        label = section_label(section.section_id)
        if errors:
            print(f"❌ {label}でエラー: " + " / ".join(errors))

    return plan


def summarize_conditions(plan: List[SectionState], participants: Dict[str, Participant],
                          config: BackSimConfig) -> dict:
    """裏シュミの結果に対して、指定した条件がどれだけ満たせたかをまとめる
    (UI表示・CLI表示のどちらからも使える形の辞書を返す)。"""
    together_ids, apart_ids, driver_ranges, together_or_group_ids = _resolve_pairs_and_ranges(config, participants)
    occ, _ = _baseline_occupancy(plan)

    by_person: Dict[str, set] = {}
    for (pid, k, s) in occ:
        by_person.setdefault(pid, set()).add((k, s))

    def shared(a, b):
        return sum(1 for (k, s) in by_person.get(a, set()) if (b, k, s) in occ)

    together_summary = [
        {"a": participants[a].name, "b": participants[b].name, "shared_sections": shared(a, b)}
        for a, b in together_ids
    ]
    apart_summary = [
        {"a": participants[a].name, "b": participants[b].name, "shared_sections": shared(a, b)}
        for a, b in apart_ids
    ]

    def shared_with_any(anchor, group):
        """anchorが候補者(group)のうち誰か一人とでも同じ車になれた区間数と、
        その区間ごとに実際に一緒だった相手の名前を返す。"""
        count = 0
        partners_by_section = []
        for (k, s) in sorted(by_person.get(anchor, set()), key=lambda ks: (ks[1] is None, ks[1])):
            partners = [participants[g].name for g in group if (g, k, s) in occ]
            if partners:
                count += 1
                partners_by_section.append({"section_id": s, "with": partners})
        return count, partners_by_section

    together_or_summary = []
    for anchor, group in together_or_group_ids:
        count, detail = shared_with_any(anchor, group)
        together_or_summary.append({
            "anchor": participants[anchor].name,
            "group": [participants[g].name for g in group],
            "satisfied_sections": count,
            "detail": detail,
        })

    driver_count: Dict[str, int] = {}
    for section in plan:
        for car in section.cars:
            if car.driver_id in participants:
                driver_count[car.driver_id] = driver_count.get(car.driver_id, 0) + 1

    driver_summary = [
        {
            "name": participants[p].name,
            "min": min_r,
            "max": max_r,
            "actual": driver_count.get(p, 0),
            "satisfied": driver_count.get(p, 0) >= min_r and (max_r is None or driver_count.get(p, 0) <= max_r),
        }
        for p, (min_r, max_r) in driver_ranges.items()
    ]

    priority_details, _ = check_priority_sections(plan, participants)

    return {
        "together": together_summary,
        "apart": apart_summary,
        "driver_ranges": driver_summary,
        "together_or": together_or_summary,
        "priority_sections": priority_details,
    }
