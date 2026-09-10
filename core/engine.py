"""계산 엔진 — 실제 원본 워크북(3p 시트)의 행 구조를 그대로 따릅니다.

※ 업로드해주신 `카탈리스터_7월_자금운용계획_260714_.xlsm` 원본을 직접 읽어
   행 번호·계정과목 이름·수식을 전부 대조해서 맞췄습니다.

금액 단위: 입력은 '원'(외화는 해당 통화 단위), 집계 결과는 '백만원'(외화는 천$).
"""
from __future__ import annotations

from dataclasses import dataclass, field

from core.config import MILLION
from core.data_store import Txn

# ── 3p(유동자금세부) 행 정의 ──────────────────────────────────────
# (row_key, 3p 행번호, 표시이름, 방향, [원본 계정과목 이름들])
LEAF: list[tuple[str, int, str, str, list[str]]] = [
    # 수입
    ("income.sales",       9,  "매출",        "in",  ["외상매출금"]),
    ("income.taxrefund",   10, "세금환급",    "in",  ["세금환급"]),
    ("income.interest",    11, "이자수익",    "in",  ["이자수익"]),
    ("income.etc",         12, "기타",        "in",  ["기타입금", "수입임대료"]),
    # 지출 — 인건비 (14 = 15~18 합)
    ("expense.salary",     15, "급여",        "out", ["급여"]),
    ("expense.bonus",      16, "상여",        "out", ["상여"]),
    ("expense.pension",    17, "퇴직연금",    "out", ["퇴직연금"]),
    ("expense.insurance",  18, "4대보험료",   "out", ["4대보험"]),
    # 지출 — 단독 행
    ("expense.marketing",  19, "마케팅비",    "out", ["광고판촉비"]),
    # 지출 — 제세공과금 (20 = 21~24 합)
    ("expense.tax.corp",   21, "법인세",         "out", ["법인세"]),
    ("expense.tax.vat",    22, "부가가치세",     "out", ["부가가치세"]),
    ("expense.tax.wht",    23, "원천세",         "out", ["원천세"]),
    ("expense.tax.etc",    24, "기타제세공과금", "out", ["기타제세공과금"]),
    # 지출 — 단독 행
    ("expense.interest",   25, "지급이자",    "out", ["지급이자"]),
    ("expense.capex",      26, "자산매입",    "out", ["S/W", "H/W", "시설장치", "차량"]),
    ("expense.invest",     27, "투자금",      "out", ["투자금", "지분투자"]),
    # 지출 — 기타 (28 = 29,30,31,32 합)
    ("expense.it",         29, "전산관련지출",             "out", ["전산유지비", "전산관련지출"]),
    ("expense.outsource",  30, "개발 외주용역비",          "out", ["외주용역비"]),
    ("expense.support",    31, "경영, 업무지원, 운영용역비", "out", ["경영지원용역비"]),
    # 32 기타 = 통신비등(38) + 33~37
    ("expense.welfare",    33, "복리후생비",            "out", ["복리후생비"]),
    ("expense.cafeteria",  34, "식당, 카페 등 운영비",  "out", ["식당운영비"]),
    ("expense.staffcost",  35, "직원비용, 복지포인트",  "out", ["직원비용"]),
    ("expense.rent",       36, "임차료, 관리비",        "out", ["지급임차료"]),
    ("expense.fee",        37, "지급수수료",            "out", ["지급수수료"]),
    ("expense.telecom",    38, "통신비등",              "out", ["기타지출"]),
    # 운용자금 / FX / 조달
    ("fund.release",   40, "해지(입금)",  "in",  ["금융상품"]),
    ("fund.deposit",   41, "예치(지출)",  "out", ["금융상품"]),
    ("fx.in",          43, "FX 입금",     "in",  ["외환매도"]),
    ("fx.out",         44, "FX 출금",     "out", ["외환매입"]),
    ("finance.in",     46, "조달 입금",   "in",  ["차입금"]),
    ("finance.out",    47, "조달 출금",   "out", ["차입금상환"]),
]

ROW_LABEL = {k: label for k, _, label, _, _ in LEAF}
ROW_NO = {k: no for k, no, _, _, _ in LEAF}

# 계정과목 → row_key (방향별)
ACCOUNT_MAP: dict[tuple[str, str], str] = {}
for _k, _no, _lb, _dir, _accts in LEAF:
    for _a in _accts:
        ACCOUNT_MAP[(_a, _dir)] = _k
ACCOUNT_MAP[("투자금", "in")] = "finance.in"     # 투자금이 들어오면 조달(46행)

ALL_ACCOUNTS = sorted({a for (a, _) in ACCOUNT_MAP})

# 집계에서 빼는 계정과목 (계좌간 대체 등)
EXCLUDED_ACCOUNTS = {"보통예금", "자금대체", "제외"}

SUB_PAYROLL = ["expense.salary", "expense.bonus", "expense.pension", "expense.insurance"]
SUB_TAX = ["expense.tax.corp", "expense.tax.vat", "expense.tax.wht", "expense.tax.etc"]
SUB_ETC32 = ["expense.telecom", "expense.welfare", "expense.cafeteria",
             "expense.staffcost", "expense.rent", "expense.fee"]
SUB_ETC28 = ["expense.it", "expense.outsource", "expense.support"]
INCOME_LEAVES = ["income.sales", "income.taxrefund", "income.interest", "income.etc"]

# 미리보기 화면 표시 순서
DISPLAY: list[tuple[str, str]] = [
    ("opening", "A. 기초유동잔고"),
    ("income.total", "B-1. 수입"),
    *[(k, "  · " + ROW_LABEL[k]) for k in INCOME_LEAVES],
    ("expense.total", "B-2. 지출"),
    ("expense.payroll", "  · 인건비"),
    *[(k, "      - " + ROW_LABEL[k]) for k in SUB_PAYROLL],
    ("expense.marketing", "  · 마케팅비"),
    ("expense.tax", "  · 제세공과금"),
    *[(k, "      - " + ROW_LABEL[k]) for k in SUB_TAX],
    ("expense.interest", "  · 지급이자"),
    ("expense.capex", "  · 자산매입"),
    ("expense.invest", "  · 투자금"),
    ("expense.etc", "  · 기타"),
    *[(k, "      - " + ROW_LABEL[k]) for k in SUB_ETC28],
    ("expense.etc32", "      - 기타(소계)"),
    *[(k, "          · " + ROW_LABEL[k]) for k in SUB_ETC32],
    ("balance", "B. 수지차(=수입-지출)"),
    ("fund.net", "C. 운용자금 입지출(=해지-예치)"),
    ("fund.release", "  · 해지(입금)"),
    ("fund.deposit", "  · 예치(지출)"),
    ("fx.net", "D. FX"),
    ("finance.net", "E. 조달"),
    ("liquid.change", "F. 유동자금 증감(=B+C+D+E)"),
    ("closing", "G. 기말유동잔고(=A+F)"),
]


@dataclass
class Result:
    values: dict[str, float] = field(default_factory=dict)   # 환산계(백만원)
    krw: dict[str, float] = field(default_factory=dict)      # 원화(백만원)
    usd: dict[str, float] = field(default_factory=dict)      # 외화(천$)
    opening: float = 0.0
    closing: float = 0.0
    unmapped: list[Txn] = field(default_factory=list)        # V1
    unregistered: list[Txn] = field(default_factory=list)    # V4
    rows_by_key: dict[str, list[Txn]] = field(default_factory=dict)


def aggregate(txns: list[Txn], opening_krw: float = 0.0, opening_usd: float = 0.0,
              fx_rate: float | None = None) -> Result:
    res = Result(opening=opening_krw)
    for k, _, _, _, _ in LEAF:
        res.krw[k] = 0.0
        res.usd[k] = 0.0

    for t in txns:
        if t.excluded:
            continue
        acct = (t.tag or "").strip()
        if not acct or acct in EXCLUDED_ACCOUNTS:
            continue
        key = ACCOUNT_MAP.get((acct, t.direction))
        if key is None:
            (res.unregistered if acct in ALL_ACCOUNTS else res.unmapped).append(t)
            continue
        if (t.currency or "KRW").upper() == "KRW":
            res.krw[key] += t.amount / MILLION
        else:
            res.usd[key] += t.amount / 1000
        res.rows_by_key.setdefault(key, []).append(t)

    rate = fx_rate or 0.0
    v = res.values
    for k in res.krw:
        v[k] = res.krw[k] + res.usd[k] * rate / 1000

    v["expense.payroll"] = sum(v[k] for k in SUB_PAYROLL)
    v["expense.tax"] = sum(v[k] for k in SUB_TAX)
    v["expense.etc32"] = sum(v[k] for k in SUB_ETC32)
    v["expense.etc"] = sum(v[k] for k in SUB_ETC28) + v["expense.etc32"]
    v["income.total"] = sum(v[k] for k in INCOME_LEAVES)
    v["expense.total"] = (v["expense.payroll"] + v["expense.marketing"] + v["expense.tax"]
                          + v["expense.interest"] + v["expense.capex"] + v["expense.invest"]
                          + v["expense.etc"])
    v["balance"] = v["income.total"] - v["expense.total"]
    v["fund.net"] = v["fund.release"] - v["fund.deposit"]
    v["fx.net"] = v["fx.in"] - v["fx.out"]
    v["finance.net"] = v["finance.in"] - v["finance.out"]
    v["liquid.change"] = v["balance"] + v["fund.net"] + v["fx.net"] + v["finance.net"]
    v["opening"] = opening_krw + opening_usd * rate / 1000
    res.closing = v["opening"] + v["liquid.change"]
    v["closing"] = res.closing
    return res


# ── 운용자금(정기예금 등) 잔고 (§11) ──────────────────────────────
def compute_fund_balance_change(txns: list[Txn], opening_krw: float,
                                opening_usd: float) -> dict[str, float]:
    """기말 운용자금 = 기초 - (해지 - 예치). 해지되면 운용자금이 줄어듭니다.
    원본 2-1 시트: D(증감) = 예치 - 해지, E(기말) = 기초 + D."""
    rel = {"KRW": 0.0, "USD": 0.0}
    dep = {"KRW": 0.0, "USD": 0.0}
    for t in txns:
        if t.excluded or (t.tag or "").strip() != "금융상품":
            continue
        cur = "USD" if (t.currency or "KRW").upper() == "USD" else "KRW"
        (rel if t.direction == "in" else dep)[cur] += t.amount
    return {
        "release_krw": rel["KRW"] / MILLION, "deposit_krw": dep["KRW"] / MILLION,
        "release_usd": rel["USD"] / 1000, "deposit_usd": dep["USD"] / 1000,
        "end_krw": opening_krw - (rel["KRW"] - dep["KRW"]) / MILLION,
        "end_usd": opening_usd - (rel["USD"] - dep["USD"]) / 1000,
    }


# ── 자금실적(당해) 세부그룹 ───────────────────────────────────────
ANNUAL_GROUP: dict[str, tuple[str, str]] = {
    "expense.salary": ("인건비", "경상"), "expense.bonus": ("인건비", "경상"),
    "expense.pension": ("인건비", "경상"), "expense.insurance": ("인건비", "경상"),
    "expense.capex": ("S/W, H/W 매입", "경상"),
    "expense.tax.corp": ("제세공과금", "경상"), "expense.tax.vat": ("제세공과금", "경상"),
    "expense.tax.wht": ("제세공과금", "경상"), "expense.tax.etc": ("제세공과금", "경상"),
    "expense.it": ("전산관련지출", "경상"),
    "expense.support": ("경영, 업무지원, 운영용역비", "경상"),
    "expense.welfare": ("복리후생비", "경상"),
    "expense.cafeteria": ("식당, 카페, 매점 등 운영비", "경상"),
    "expense.staffcost": ("직원비용, 복지포인트 사용비", "경상"),
    "expense.rent": ("임차료, 관리비", "경상"),
    "expense.fee": ("지급수수료", "경상"),
    "expense.telecom": ("통신비등", "경상"),
    "expense.invest": ("투자금", "비경상"),
    "expense.marketing": ("마케팅,지급이자,외주용역비", "경상"),
    "expense.interest": ("마케팅,지급이자,외주용역비", "경상"),
    "expense.outsource": ("마케팅,지급이자,외주용역비", "경상"),
}

# 수입 경상/비경상 (v0.12 사용자 결정: 자산매각대금·투자분배금·투자수익만 비경상)
EXTRAORDINARY_INCOME_ACCOUNTS = {"자산매각대금", "투자분배금", "투자수익"}


def income_class(row_key: str, txns: list[Txn]) -> str:
    for t in txns:
        if (t.tag or "").strip() in EXTRAORDINARY_INCOME_ACCOUNTS:
            return "비경상"
    return "경상"


# ── 검증 게이트 6종 ───────────────────────────────────────────────
@dataclass
class Gate:
    id: str
    name: str
    passed: bool
    detail: str = ""
    items: list = field(default_factory=list)


ANNUAL_ROWS_EXIST = {
    "인건비", "S/W, H/W 매입", "제세공과금", "전산관련지출",
    "경영, 업무지원, 운영용역비", "복리후생비", "식당, 카페, 매점 등 운영비",
    "직원비용, 복지포인트 사용비", "임차료, 관리비", "지급수수료", "통신비등",
    "마케팅,지급이자,외주용역비",   # tools/add_marketing_row.py 로 추가함
    "지식재산권 양수", "임대보증금", "시설장치", "투자금", "기타",
}


def validate(txns: list[Txn], result: Result,
             external_balance: float | None = None,
             fx_rate: float | None = None) -> list[Gate]:
    gates: list[Gate] = []

    gates.append(Gate("V1", "전액 대사 (계정과목 부재)", not result.unmapped,
                      "계정과목 부재, 오타 확인 필요" if result.unmapped else "이상없음",
                      [f"'{t.tag}' — {t.memo or t.counterpart}" for t in result.unmapped]))

    if external_balance is None:
        gates.append(Gate("V2", "외부 정합성 (수동 비교)", True, "비교값 미입력 — 건너뜀"))
    else:
        diff = round(result.closing - external_balance, 6)
        # 화면 표시는 소수점 둘째 자리까지만 (합격/불합격 판정은 6자리 그대로)
        gates.append(Gate("V2", "외부 정합성", abs(diff) < 0.001,
                          f"계산 기말 {result.closing:,.2f} vs 입력 {external_balance:,.2f}"
                          f" (차이 {diff:,.2f})"))

    v = result.values
    ok3 = (abs(sum(v[k] for k in INCOME_LEAVES) - v["income.total"]) < 1e-9
           and abs(sum(v[k] for k in SUB_PAYROLL) - v["expense.payroll"]) < 1e-9
           and abs(sum(v[k] for k in SUB_TAX) - v["expense.tax"]) < 1e-9
           and abs(sum(v[k] for k in SUB_ETC32) - v["expense.etc32"]) < 1e-9)
    gates.append(Gate("V3", "소계 정합", ok3, "이상없음" if ok3 else "소계와 하위합이 불일치"))

    hints = [f"'{t.tag}' — {t.memo or t.counterpart}"
             f" (이 거래는 {'입금' if t.direction == 'in' else '출금'})"
             for t in result.unregistered]
    gates.append(Gate("V4", "태그 등록 (방향 조합)", not hints,
                      "거래 내역 확인 필요" if hints else "이상없음", hints))

    has_fx = any((t.currency or "KRW").upper() != "KRW" for t in txns if not t.excluded)
    ok5 = (not has_fx) or bool(fx_rate)
    gates.append(Gate("V5", "환율", ok5,
                      "환율 불필요 — 자동 통과" if not has_fx
                      else (f"적용 환율 {fx_rate}" if fx_rate else "외화 거래가 있으나 환율 없음")))

    missing = sorted({
        ANNUAL_GROUP[k][0] for k in result.rows_by_key
        if k in ANNUAL_GROUP and ANNUAL_GROUP[k][0] not in ANNUAL_ROWS_EXIST
        and abs(result.values.get(k, 0.0)) > 1e-9
    })
    gates.append(Gate("V6", "자금실적 세부그룹 매핑", not missing,
                      "이상없음" if not missing
                      else "자금실적(당해) 시트에 대응 행이 없는 그룹이 있습니다", missing))
    return gates
