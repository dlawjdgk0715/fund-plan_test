"""③ 계정 분류 확인 — 적요/계정과목 글자를 태그 사전과 비교해 태그를 제안합니다.

점수 기준(설계서 M2):
  90점 이상 → 자동 확정
  70~89점  → 추천(사람이 확인)
  70점 미만 → 제안 없음(직접 선택)
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass

from core.data_store import Txn, load_tags

AUTO, SUGGEST = 90, 70

# 예전 표기 → 표준 태그 별칭 교정
ALIASES = {
    "임금": "급여", "월급": "급여", "인건비": "급여", "상여금": "상여",
    "퇴직금": "퇴직연금", "국민연금": "4대보험", "건강보험": "4대보험",
    "고용보험": "4대보험", "산재보험": "4대보험",
    "소프트웨어": "S/W", "하드웨어": "H/W", "서버": "H/W",
    "광고선전비": "광고판촉비", "판촉비": "광고판촉비", "마케팅": "광고판촉비",
    "이자비용": "지급이자", "부가세": "부가가치세", "세금과공과": "기타제세공과금",
    "주민세": "기타제세공과금", "전산유지": "전산유지비", "클라우드": "전산유지비",
    "외주": "외주용역비", "경영지원": "경영지원용역비", "업무지원": "경영지원용역비",
    "복지포인트": "직원비용", "법인카드": "직원비용", "직원경비": "직원비용",
    "임차료": "지급임차료", "관리비": "지급임차료", "수수료": "지급수수료",
    "정기예금": "금융상품", "예금": "금융상품", "통신비": "기타지출",
    "차량": "차량", "인테리어": "시설장치",
}


def _clean(s: str) -> str:
    return re.sub(r"[\s\(\)\[\]/,.·\-_]", "", str(s or "")).lower()


def _score(a: str, b: str) -> int:
    a, b = _clean(a), _clean(b)
    if not a or not b:
        return 0
    if a == b:
        return 100
    if a in b or b in a:
        return 92
    return int(difflib.SequenceMatcher(None, a, b).ratio() * 100)


@dataclass
class Suggestion:
    tag: str
    score: int
    level: str   # 'auto' | 'suggest' | 'manual'
    reason: str


def suggest_tag(txn: Txn) -> Suggestion:
    tags = load_tags()

    # 1) 파일에 계정과목이 이미 적혀 있으면 그것부터 (별칭 교정 포함)
    if txn.tag:
        raw = _clean(txn.tag)
        if txn.tag in tags:
            return Suggestion(txn.tag, 100, "auto", "파일의 계정과목과 정확히 일치")
        for alias, target in ALIASES.items():
            if _clean(alias) == raw or _clean(alias) in raw:
                return Suggestion(target, 95, "auto", f"예전 표기 '{txn.tag}' → '{target}' 교정")
        best = max(tags, key=lambda t: _score(txn.tag, t))
        s = _score(txn.tag, best)
        if s >= AUTO:
            return Suggestion(best, s, "auto", f"계정과목 글자 유사({s}점)")

    # 2) 적요를 태그 이름·적요 예시와 비교
    text = f"{txn.memo} {txn.counterpart}".strip()
    best_tag, best_score, why = "", 0, ""
    for tag, info in tags.items():
        s = _score(text, tag)
        w = f"태그 이름 유사({s}점)"
        for ex in info.example_list:
            se = _score(text, ex)
            if se > s:
                s, w = se, f"적요 예시 '{ex}' 유사({se}점)"
        if s > best_score:
            best_tag, best_score, why = tag, s, w
    for alias, target in ALIASES.items():
        if _clean(alias) and _clean(alias) in _clean(text):
            if 93 > best_score:
                best_tag, best_score, why = target, 93, f"키워드 '{alias}' 발견"

    if best_score >= AUTO:
        return Suggestion(best_tag, best_score, "auto", why)
    if best_score >= SUGGEST:
        return Suggestion(best_tag, best_score, "suggest", why)
    return Suggestion("", best_score, "manual", "비슷한 항목을 못 찾았습니다 — 직접 선택해주세요")


def classify_all(txns: list[Txn]) -> list[tuple[Txn, Suggestion]]:
    return [(t, suggest_tag(t)) for t in txns]


# ── 원본 파일의 계정과목 자동 교정 ────────────────────────────────
# 일일자금계획의 계정과목 열이 **실제 내용과 다르게** 적혀 오는 경우가 있습니다.
# (예: 삼성카드 '공용법인카드대금' 이 계정과목은 '기타지출' 로 적혀 옴)
# 합계에는 영향이 없지만(둘 다 3p '기타 > 기타'), 분류가 정확해야 나중에 쓰기 좋습니다.
# 매달 손으로 고치지 않도록 아래 규칙으로 자동 교정하고, 무엇을 고쳤는지 화면에 보여줍니다.
# word        적요에 있어야 하는 낱말
# cp          거래처에 있어야 하는 낱말 (빈칸이면 거래처는 안 봄)
# src → dst   원본 계정과목이 src 일 때만 dst 로 바꿉니다
RETAG_RULES: list[dict] = [
    {"word": "법인카드", "cp": "카드", "src": "기타지출", "dst": "직원비용",
     "why": "카드사 법인카드 대금은 직원비용으로 분류"},
]


def retag(t: Txn) -> tuple[str, str] | None:
    """계정과목을 고쳐야 하면 (새 계정과목, 이유), 아니면 None.

    ⚠️ 조건을 넉넉하게 잡으면 엉뚱한 거래까지 바뀝니다.
       (실제로 '법인카드' 만 보고 잡았더니 '법인카드데이타 연동서비스'(소프트웨어)
        까지 직원비용으로 바뀌었습니다 → 거래처가 **카드사**인 건만 잡도록 좁힘)
    """
    cur = (t.tag or "").strip()
    if not cur:
        return None
    memo = t.memo or ""
    cp = t.counterpart or ""
    for r in RETAG_RULES:
        if cur != r["src"] or r["dst"] == cur:
            continue
        if r["word"] not in memo:
            continue
        if r["cp"] and r["cp"] not in cp:
            continue
        return r["dst"], r["why"]
    return None


def apply_retag(txns: list[Txn]) -> list[dict]:
    """목록 전체에 자동 교정을 적용하고, 무엇을 고쳤는지 돌려줍니다."""
    changed = []
    for t in txns:
        got = retag(t)
        if got:
            new_tag, why = got
            changed.append({"거래처": t.counterpart, "적요": t.memo,
                            "원본 계정과목": t.tag, "바꾼 계정과목": new_tag,
                            "이유": why})
            t.tag = new_tag
    return changed
