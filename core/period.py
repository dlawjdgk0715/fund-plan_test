"""월(period) 다루기 — 'YYYY-MM' 문자열 하나로 실적/계획 월을 모두 정합니다.

규칙: 실적 월 = period,  계획 월 = period의 다음 달
"""
from __future__ import annotations


def parse(period: str) -> tuple[int, int]:
    y, m = period.strip().split("-")
    y, m = int(y), int(m)
    if not 1 <= m <= 12:
        raise ValueError(f"월이 잘못됐습니다: {period}")
    return y, m


def next_period(period: str) -> str:
    y, m = parse(period)
    return f"{y+1}-01" if m == 12 else f"{y}-{m+1:02d}"


def prev_period(period: str) -> str:
    y, m = parse(period)
    return f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"


def actual_month_of(period: str) -> int:
    """실적 월 (숫자)"""
    return parse(period)[1]


def plan_month_of(period: str) -> int:
    """계획 월 (숫자) — 실적 다음 달"""
    return parse(next_period(period))[1]


def fmt(period: str) -> str:
    """화면 표시용: '2026-07' → '2026년 7월'"""
    y, m = parse(period)
    return f"{y}년 {m}월"
