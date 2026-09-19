from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class Participant:
    id: str
    name: str
    preferred_sections: List[bool]
    can_drive: bool
    can_drive_large: bool
    can_drive_mountain: bool
    staying_overnight: bool
    grade: int
    remaining_sections: int = 0
    leaves_after_section: Optional[int] = None  # 日帰りの人がこの区間まで参加して離脱する(Noneなら最後まで参加)
    priority_sections: List[bool] = field(default_factory=lambda: [False] * 10)  # 「特に走りたい区間」(重み付けのみ、ハード制約ではない)

@dataclass
class CarState:
    car_id: str
    driver_id: str
    passenger_ids: List[str] = field(default_factory=list)
    is_mountain_goer: bool = False
    is_advance: bool = False  # 先行車（次走者を乗せて早めに次の中継所へ向かう）
    car_type: str = "normal"  # "large" or "normal"
    group: Optional[str] = None  # "hotel" or "mountain"（9・10区のみ使用）

    @property
    def total_people(self) -> int:
        return 1 + len(self.passenger_ids)

@dataclass
class SectionState:
    section_id: int
    runner_ids: List[str]
    cars: List[CarState]