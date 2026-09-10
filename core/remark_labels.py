"""비고 라벨 축약표 (v0.11).

적요를 그대로 쓰면 비고가 너무 장황해서, 짧은 라벨로 바꿔 씁니다.
라벨은 DB(remark_labels 테이블)로 관리하고, '6 비고 라벨 관리' 화면에서 바로 고칠 수 있습니다.
"""
from __future__ import annotations

from core.data_store import load_remark_labels, save_remark_label
from core.engine import ROW_LABEL

# 사용자가 준 기본 예시
DEFAULTS = {
    ("인건비", "임직원 6월 급여(6/25)"): "급여",
    ("지급수수료", "카탈리스터 전산자산 매각에 따른 위탁매매 수수료"): "전산자산 매각 수수료",
}


def get_label(tag: str, memo: str, row_key: str = "") -> str:
    """라벨 우선순위: DB 저장값 → 기본 예시 → 유동자금세부 행 이름 → 적요 원문"""
    labels = load_remark_labels()
    if (tag, memo) in labels:
        return labels[(tag, memo)]
    if (tag, memo) in DEFAULTS:
        return DEFAULTS[(tag, memo)]
    if row_key and row_key in ROW_LABEL:
        return ROW_LABEL[row_key]
    return memo or tag


def seed_defaults() -> None:
    for (tag, memo), label in DEFAULTS.items():
        save_remark_label(tag, memo, label)
