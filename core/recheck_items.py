"""매달 손으로 챙겨야 하는 항목 알림 (M6).

취합 파일(일일자금계획·부서 계획서)에 **매번 나오지는 않는** 항목입니다.
이번 달 거래에 하나도 안 보이면 '3 실적/계획 관리' 화면에 안내창을 띄웁니다.
입력을 강제하지는 않습니다 — 해당 없는 달이면 그냥 넘어가면 됩니다.

찾는 조건 (규칙은 DB `recheck_items` 에 있습니다)
    pattern     적요 또는 거래처에 이 낱말이 들어 있어야 함
    counterpart 거래처 조건 (빈칸이면 안 봄)
    tag         계정과목 조건 (빈칸이면 안 봄)

※ 계정과목(tag)까지 맞아야 찾은 것으로 치던 예전 방식은, 계정과목이 아직
   안 붙은 거래를 못 알아봐서 **있는데도 없다고** 하는 일이 있었습니다.
   지금은 기본 규칙 두 개 모두 계정과목 조건을 비워 두고 낱말로만 찾습니다.
"""
from __future__ import annotations

from core.data_store import Txn, load_recheck_items


def matches(t: Txn, item: dict) -> bool:
    """거래 한 건이 그 규칙에 해당하는지."""
    if t.excluded:
        return False
    tag = (item.get("tag") or "").strip()
    if tag and (t.tag or "").strip() != tag:
        return False
    cp = (item.get("counterpart") or "").strip()
    if cp and cp not in (t.counterpart or ""):
        return False
    pat = (item.get("pattern") or "").strip()
    if pat and pat not in f"{t.memo or ''} {t.counterpart or ''}":
        return False
    return True


def check_all(txns: list[Txn]) -> list[tuple[dict, list[Txn]]]:
    """규칙별로 (규칙, 해당하는 거래들)을 돌려줍니다. 거래가 비었으면 '없음'."""
    return [(item, [t for t in txns if matches(t, item)])
            for item in load_recheck_items()]


def check_missing(txns: list[Txn]) -> list[dict]:
    """이번 달에 안 보이는 규칙들을 돌려줍니다 (규칙 dict 그대로)."""
    return [item for item, hits in check_all(txns) if not hits]


def condition_text(item: dict) -> str:
    """사람이 읽는 조건 설명 — '거래처 삼성카드 + 적요 법인카드' 처럼."""
    bits = []
    if (item.get("counterpart") or "").strip():
        bits.append(f"거래처 **{item['counterpart']}**")
    if (item.get("pattern") or "").strip():
        bits.append(f"적요에 **{item['pattern']}**")
    if (item.get("tag") or "").strip():
        bits.append(f"계정과목 **{item['tag']}**")
    return " + ".join(bits) if bits else "(조건 없음)"
