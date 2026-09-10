"""비고의 '쓰는 방식'(서식)만 학습합니다.

무엇을 배우나 — **틀만** 배웁니다.
    · 금액 단위      : 백만 / 억 / 천원 …
    · 소수 자릿수    : 12.3 인지 12 인지
    · 항목 개수      : 최대 몇 건까지 쓰는지
    · 구분 기호      : 줄바꿈인지 ", " 인지 " / " 인지
    · 줄 앞머리      : "∙7월 : " 처럼 매달 달라지는 숫자는 {m} 으로 기억
    · 꼬리말         : 생략한 항목이 있을 때 " 등" 을 붙이는지

무엇을 배우지 않나 — **문장 내용**입니다.
    어떤 거래가 컸는지는 매달 새로 계산합니다. 그래서 금액이 바뀌어도 틀리지 않습니다.
"""
from __future__ import annotations

import re

# 기본 서식 (지금까지 앱이 쓰던 방식)
SUMMARY_STYLE = {          # 요약1 서술형 · 자금수지 표 — 한 줄에 쉼표로
    "prefix": "{m}월 : ",
    "sep": ", ",
    "max_items": 3,
    "unit": "백만",
    "decimals": 1,
    "tail": " 등",
}
LEAF_STYLE = {             # 유동자금세부 · PPT 6장 — 건별로 줄바꿈
    "prefix": "∙{m}월 : ",
    "sep": "\n",
    "max_items": 3,
    "unit": "백만",
    "decimals": 1,
    "tail": "",
}
DEFAULT_STYLE = {"cash_flow": SUMMARY_STYLE, "summary1": SUMMARY_STYLE,
                 "detail": LEAF_STYLE}

# 단위 이름 → 백만원 기준 나눗셈 값
UNIT_DIV = {"억원": 100.0, "억": 100.0, "백만원": 1.0, "백만": 1.0,
            "천원": 0.001, "원": 0.000001}

_SEP_CANDIDATES = (" / ", " · ", "; ", ", ", " ,", "/")
_NUM_UNIT = re.compile(r"(-?\d[\d,]*(?:\.(\d+))?)\s*(억원|억|백만원|백만|천원|원)")
_PREFIX = re.compile(r"^\s*([^:：]{0,12}[:：]\s*)")


def style_for(table_key: str, learned: dict | None = None) -> dict:
    """기본 서식에 학습된 서식을 덮어씌워 돌려줍니다."""
    base = dict(DEFAULT_STYLE.get(table_key, SUMMARY_STYLE))
    if learned:
        for k in base:
            if k in learned and learned[k] is not None:
                base[k] = learned[k]
    return base


def _fmt_amount(v_million: float, style: dict) -> str:
    div = UNIT_DIV.get(style.get("unit", "백만"), 1.0)
    d = int(style.get("decimals", 1))
    s = f"{v_million / div:,.{d}f}"
    if d and s.endswith("." + "0" * d):        # 12.0 → 12
        s = s[: -(d + 1)]
    return s + style.get("unit", "백만")


def render_items(pairs: list[tuple[str, float]], month: int, style: dict) -> str:
    """(라벨, 금액[백만원]) 목록을 학습된 서식대로 한 문장으로 만듭니다."""
    pairs = [p for p in pairs if abs(p[1]) >= 0.05]
    if not pairs:
        return ""
    n = max(1, int(style.get("max_items", 3)))
    head, rest = pairs[:n], pairs[n:]
    prefix = str(style.get("prefix", "")).replace("{m}", str(month))
    sep = style.get("sep", ", ")
    body = [f"{lab} {_fmt_amount(amt, style)}" for lab, amt in head]
    if sep == "\n":                     # 줄마다 앞머리를 붙이는 서식
        out = "\n".join(prefix + b for b in body)
        return out + (style.get("tail", "") if rest else "")
    return prefix + sep.join(body) + (style.get("tail", "") if rest else "")


def infer_style(text: str, month: int, base: dict) -> dict | None:
    """사용자가 고친 문장에서 '틀'만 뽑아냅니다. 못 알아보면 None."""
    t = (text or "").strip()
    if not t:
        return None
    st = dict(base)

    lines = [ln.strip() for ln in t.split("\n") if ln.strip()]
    if len(lines) > 1:
        st["sep"] = "\n"
        items, first = lines, lines[0]
    else:
        first = lines[0]
        m = _PREFIX.match(first)
        body = first[m.end():] if m else first
        for cand in _SEP_CANDIDATES:
            if cand in body:
                st["sep"] = cand
                break
        else:
            st["sep"] = base.get("sep", ", ")
        items = [x.strip() for x in body.split(st["sep"]) if x.strip()]

    # 줄 앞머리 — 이번 달 숫자를 {m} 으로 되돌려 기억
    #
    # ⚠️ **월 표기는 절대 잃어버리면 안 됩니다.** 비고는 '언제 생긴 거래인지'가
    #    핵심이라, 앞머리에 {m}(월)이 안 남는 형태는 배우지 않고 기본값을 지킵니다.
    #    (예전에는 사용자가 콜론 없는 문장을 한 번 쓰면 그 뒤 모든 초안에서
    #     '∙7월 : ' 가 통째로 사라졌습니다)
    m = _PREFIX.match(first)
    cand = m.group(1).replace(f"{month}월", "{m}월") if m else ""
    st["prefix"] = cand if "{m}" in cand else base.get("prefix", "")

    # 꼬리말
    st["tail"] = " 등" if re.search(r"등\s*$", t) else ""

    # 항목 개수 — '등'으로 끝나면 그 조각은 개수에서 뺍니다
    n = len(items)
    if st["tail"] and items and re.fullmatch(r"등", items[-1]):
        n -= 1
    if n >= 1:
        st["max_items"] = min(n, 9)

    # 금액 단위 · 소수 자릿수
    hit = _NUM_UNIT.search(t)
    if hit:
        st["unit"] = hit.group(3)
        st["decimals"] = len(hit.group(2) or "")
    return st


def learn(table_key: str, drafts: list[str], edited: list[str], month: int) -> dict | None:
    """초안과 사용자가 고친 글을 비교해, 실제로 바뀐 줄에서만 서식을 배웁니다.

    반환값이 None 이면 배울 게 없었다는 뜻입니다(저장하지 않습니다).
    """
    from collections import Counter

    base = style_for(table_key)
    votes: list[tuple] = []
    for d, e in zip(drafts, edited):
        if (e or "").strip() and (e or "").strip() != (d or "").strip():
            s = infer_style(e, month, base)
            if s:
                votes.append((s["prefix"], s["sep"], s["max_items"],
                              s["unit"], s["decimals"], s["tail"]))
    if not votes:
        return None
    best = Counter(votes).most_common(1)[0][0]
    return {"prefix": best[0], "sep": best[1], "max_items": best[2],
            "unit": best[3], "decimals": best[4], "tail": best[5]}
