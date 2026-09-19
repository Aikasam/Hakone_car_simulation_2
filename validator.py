from typing import Dict, List, Tuple
from models import Participant, SectionState
from logic.car_pool import MOUNTAIN_SECTIONS, LARGE_CAPACITY, NORMAL_CAPACITY, section_label


def _name(participants: Dict[str, Participant], pid: str) -> str:
    """participants に存在しないid(NO_DRIVER等)でも安全に名前文字列を返す。"""
    p = participants.get(pid)
    return p.name if p else "unknown"


def validate_participants(participants: Dict[str, Participant]) -> List[str]:
    """参加者データの入力チェック。問題があれば警告メッセージのリストを返す。"""
    warnings = []
    for pid, p in participants.items():
        label = f"{p.name or pid}"
        if not p.name.strip():
            warnings.append(f"{pid}: 名前が空欄です")
        if p.grade < 1 or p.grade > 9:
            warnings.append(f"{label}: 学年が不正な値です ({p.grade})")
        if p.remaining_sections < 0:
            warnings.append(f"{label}: 希望区間数が負です ({p.remaining_sections})")
        if p.remaining_sections > 10:
            warnings.append(f"{label}: 希望区間数が10を超えています ({p.remaining_sections})")
        if p.remaining_sections > 0 and not any(p.preferred_sections):
            warnings.append(f"{label}: 希望区間数>{0}ですが走りたい区間が未選択です")
        if p.leaves_after_section is not None and not (1 <= p.leaves_after_section <= 10):
            warnings.append(f"{label}: 離脱区間が不正です ({p.leaves_after_section})")
    return warnings


def compute_runner_satisfaction(
    plan: List[SectionState], participants: Dict[str, Participant]
) -> Tuple[int, int, int, int]:
    """ランナー希望充足率を計算する。
    Returns: (希望者数, 希望通り走れた人数, 希望総区間数, 実際の走行総区間数)
    """
    actual_runs: Dict[str, int] = {}
    for section in plan:
        for pid in section.runner_ids:
            actual_runs[pid] = actual_runs.get(pid, 0) + 1

    wanted = [(pid, p) for pid, p in participants.items() if p.remaining_sections > 0]
    satisfied = sum(
        1 for pid, p in wanted if actual_runs.get(pid, 0) >= p.remaining_sections
    )
    total_wanted = sum(p.remaining_sections for _, p in wanted)
    total_ran = sum(actual_runs.values())
    return len(wanted), satisfied, total_wanted, total_ran

def compute_individual_summary(
    plan: List[SectionState], participants: Dict[str, Participant]
) -> Dict[str, Dict[str, str]]:
    """各参加者が各区間(1〜10区・帰路)でランナー/運転手/同乗者のどれだったかをまとめる。
    Returns: {participant_id: {区間ラベル: 役割の表示文字列}} (該当なしの区間はキー自体が無い)
    """
    summary: Dict[str, Dict[str, str]] = {pid: {} for pid in participants}
    for section in plan:
        label = section_label(section.section_id)
        for pid in section.runner_ids:
            if pid in summary:
                summary[pid][label] = "🏃 ランナー"
        for car in section.cars:
            if car.driver_id in summary:
                summary[car.driver_id][label] = f"🚘 運転手({car.car_id})"
            for pid in car.passenger_ids:
                if pid in summary:
                    summary[pid][label] = f"👥 同乗者({car.car_id})"
    return summary


def validate_section(state: SectionState, participants: Dict[str, Participant]) -> List[str]:
    errors = []
    all_runners = set(state.runner_ids)
    all_drivers = set()

    for car in state.cars:
        all_drivers.add(car.driver_id)
        driver = participants.get(car.driver_id)

        if not driver:
            errors.append(f"車{car.car_id}のドライバーが不在です")
            continue

        if car.car_type == "large":
            if car.total_people > LARGE_CAPACITY: errors.append(f"車{car.car_id}(大型)定員オーバー")
            if not driver.can_drive_large: errors.append(f"車{car.car_id}ですが{driver.name}は大型免許なし")
        else:
            if car.total_people > NORMAL_CAPACITY: errors.append(f"車{car.car_id}(普通)定員オーバー")

        if not driver.can_drive: errors.append(f"車{car.car_id}の{driver.name}は免許なし")

        if state.section_id in MOUNTAIN_SECTIONS and car.is_mountain_goer:
            if not driver.can_drive_mountain:
                errors.append(f"{state.section_id}区の山行き車両({car.car_id})ですが、{driver.name}は山道免許なし")

        if car.driver_id in all_runners:
            errors.append(f"車{car.car_id}の運転手{driver.name}は同時にランナーです")

        for pid in car.passenger_ids:
            if pid in all_runners:
                name = participants[pid].name if pid in participants else pid
                errors.append(f"車{car.car_id}の同乗者{name}は同時にランナーです")

        # 8区はソルバー側も「同乗者なし」「2年生以上なし」のペナルティを課していない
        # (次区間へ向かう先行車が単独になり得るため)。ここも合わせて対象外にする。
        if state.section_id != 8:
            if len(car.passenger_ids) == 0: errors.append(f"車{car.car_id}に同乗者なし")
            else:
                if not any(participants[pid].grade >= 2 for pid in car.passenger_ids if pid in participants):
                    errors.append(f"車{car.car_id}に2年生以上なし")

    driver_list = [car.driver_id for car in state.cars]
    for d_id in set(driver_list):
        if driver_list.count(d_id) > 1:
            name = participants[d_id].name if d_id in participants else d_id
            errors.append(f"{name}が複数の車で同時にドライバーになっています")

    return errors

def _car_of(state: SectionState, person_id: str) -> str:
    """ある区間でperson_idがどの車にいるか(運転/同乗どちらでも)を返す。乗っていなければ空文字。"""
    for car in state.cars:
        if person_id == car.driver_id or person_id in car.passenger_ids:
            return car.car_id
    return ""

def validate_transitions(sections: List[SectionState], participants: Dict[str, Participant]) -> List[str]:
    """区間をまたいだ配車の継続性をチェックする。

    - 一般の区間間: ランナーでない限り、同じ車に乗り続けているかを確認する（情報用、エラーにはしない）。
    - 8区→9区（山行き開始）: 山行き車両(is_mountain_goer)のドライバーが
      9区→10区で変わっていないかをチェックする（山組は1往復のみという前提のため、ここはエラーとする）。
    """
    errors = []
    by_section = {s.section_id: s for s in sections}

    for s_id in range(9, 11):
        prev = by_section.get(s_id - 1)
        cur = by_section.get(s_id)
        if not prev or not cur:
            continue
        prev_mt_drivers = {c.car_id: c.driver_id for c in prev.cars if c.is_mountain_goer}
        cur_mt_drivers = {c.car_id: c.driver_id for c in cur.cars if c.is_mountain_goer}
        for car_id, driver_id in cur_mt_drivers.items():
            if car_id in prev_mt_drivers and prev_mt_drivers[car_id] != driver_id:
                errors.append(
                    f"{s_id - 1}区→{s_id}区: 山行き車両{car_id}のドライバーが"
                    f"{_name(participants, prev_mt_drivers[car_id])}から"
                    f"{_name(participants, driver_id)}に変わっています"
                )

    return errors

def count_car_changes(sections: List[SectionState]) -> int:
    """区間をまたいでいて車を乗り換えた(ランナーでもなく同じ車でもない)人数の合計。連続性の指標として使う。"""
    changes = 0
    sections_sorted = sorted(sections, key=lambda s: s.section_id)
    for prev, cur in zip(sections_sorted, sections_sorted[1:]):
        all_people = set()
        for car in prev.cars:
            all_people.add(car.driver_id)
            all_people.update(car.passenger_ids)
        for car in cur.cars:
            all_people.add(car.driver_id)
            all_people.update(car.passenger_ids)
        for pid in all_people:
            if pid in prev.runner_ids or pid in cur.runner_ids:
                continue
            prev_car = _car_of(prev, pid)
            cur_car = _car_of(cur, pid)
            if prev_car and cur_car and prev_car != cur_car:
                changes += 1
    return changes