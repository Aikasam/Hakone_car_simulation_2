"""レンタカー候補車両プールの定義。

大型/普通車の容量・レンタルコスト・候補スロットIDをここに集約する。
allocator.py / milp_allocator.py / validator.py はここから定数を import し、
car_id の文字列パターンで車種を判定するハードコードをしない。
"""

MOUNTAIN_SECTIONS = [9, 10]

LARGE_CAPACITY = 8
NORMAL_CAPACITY = 4

# レンタル1台あたりの相対コスト（実際の金額ではなく、大型が普通より高いことを表す比率）
LARGE_COST = 2
NORMAL_COST = 1

# 最適化が選べる候補車スロット（実際に使うのはこの中の一部だけでよい）
LARGE_CAR_IDS = ["L1", "L2", "L3", "L4"]
NORMAL_CAR_IDS = ["N1", "N2", "N3", "N4", "N5", "N6", "N7", "N8", "N9", "N10"]
ALL_CAR_IDS = LARGE_CAR_IDS + NORMAL_CAR_IDS

CAR_TYPE = {car_id: "large" for car_id in LARGE_CAR_IDS}
CAR_TYPE.update({car_id: "normal" for car_id in NORMAL_CAR_IDS})

CAR_CAPACITY = {car_id: LARGE_CAPACITY for car_id in LARGE_CAR_IDS}
CAR_CAPACITY.update({car_id: NORMAL_CAPACITY for car_id in NORMAL_CAR_IDS})

CAR_COST = {car_id: LARGE_COST for car_id in LARGE_CAR_IDS}
CAR_COST.update({car_id: NORMAL_COST for car_id in NORMAL_CAR_IDS})


def section_label(section_id: int) -> str:
    """区間IDの表示ラベル。10区より後は「帰路」として表示する。"""
    return f"{section_id}区" if section_id <= 10 else "帰路"
