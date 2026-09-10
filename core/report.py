"""보고서 표 만들기 — 미리보기 화면과 PPT가 **같은 계산 결과**를 씁니다.

PPT 6장 구조 (2026-08 확인, `카탈리스터_7월_자금운용계획_260716__v3.pptx`)
  2장  표15  1. X월 자금수지 (전체 자금 = 유동 + 운용)
        표10  2. 유동자금 계획
  3장  표9   3. 운용자금 계획
        표13  4. X월말 자금(stock)
  5장  표1   연간 자금실적 (당해)
  6장  표2   유동자금 계획 세부내역

금액 단위: 원화 백만원, 외화 천$
"""
from __future__ import annotations

from dataclasses import dataclass, field

from core.engine import DISPLAY, Result, aggregate, compute_fund_balance_change


def conv(krw: float, usd: float, fx: float) -> float:
    """환산 계 = 원화 + 외화(천$) × 환율 ÷ 1000"""
    return krw + usd * (fx or 0) / 1000


@dataclass
class Report:
    actual_month: int
    plan_month: int
    fx: float = 0.0
    actual: Result | None = None
    plan: Result | None = None
    tables: dict[str, list[list]] = field(default_factory=dict)
    meta: dict = field(default_factory=dict)


def _num(v: float) -> float:
    return round(v, 6)


# ── 유동자금 계획 세부내역 줄 정의 (PPT 6장 순서 그대로) ─────────
# (들여쓰기 단계, 표시 이름, 값 키)
DETAIL_ROWS: list[tuple[int, str, str]] = [
    (0, "A. 기초유동잔고", "opening"),
    (0, "B. 수지차(=수입-지출)", "balance"),
    (1, "수입", "income.total"),
    (2, "매출", "income.sales"),
    (2, "세금환급", "income.taxrefund"),
    (2, "이자수익", "income.interest"),
    (2, "기타", "income.etc"),
    (1, "지출", "expense.total"),
    (2, "인건비", "expense.payroll"),
    (3, "급여", "expense.salary"),
    (3, "상여", "expense.bonus"),
    (3, "퇴직연금", "expense.pension"),
    (3, "4대보험료", "expense.insurance"),
    (2, "마케팅비", "expense.marketing"),
    (2, "제세공과금", "expense.tax"),
    (3, "법인세", "expense.tax.corp"),
    (3, "부가가치세", "expense.tax.vat"),
    (3, "원천세", "expense.tax.wht"),
    (3, "기타제세공과금", "expense.tax.etc"),
    (2, "자산매입", "expense.capex"),
    (2, "투자금", "expense.invest"),
    (2, "기타", "expense.etc"),
    (3, "전산관련지출", "expense.it"),
    (3, "경영, 운영용역비", "expense.support"),
    (3, "기타", "expense.etc.rest"),
    (0, "C. 운용자금 입지출(=해지-예치)", "fund.net"),
    (1, "해지(입금)", "fund.release"),
    (1, "예치(지출)", "fund.deposit"),
    (0, "D. FX", "fx.net"),
    (1, "입금", "fx.in"),
    (1, "출금", "fx.out"),
    (0, "E. 조달", "finance.net"),
    (1, "입금", "finance.in"),
    (1, "출금", "finance.out"),
    (0, "F. 유동자금 증감(=B+C+D+E)", "liquid.change"),
    (0, "G. 기말유동잔고(=A+F)", "closing"),
]


def _components(res: Result, open_krw: float, open_usd: float) -> dict[str, tuple[float, float]]:
    """세부내역 각 줄의 (원화, 외화) 값. 소계는 하위 항목을 더해서 만듭니다."""
    from core.engine import (INCOME_LEAVES, SUB_ETC28, SUB_ETC32, SUB_PAYROLL, SUB_TAX)

    out: dict[str, tuple[float, float]] = {}
    for k in res.krw:
        out[k] = (res.krw[k], res.usd[k])

    def total(keys):
        return (sum(res.krw.get(k, 0.0) for k in keys),
                sum(res.usd.get(k, 0.0) for k in keys))

    def sub(a, b):
        return (a[0] - b[0], a[1] - b[1])

    def add(a, b):
        return (a[0] + b[0], a[1] + b[1])

    out["income.total"] = total(INCOME_LEAVES)
    out["expense.payroll"] = total(SUB_PAYROLL)
    out["expense.tax"] = total(SUB_TAX)
    out["expense.etc32"] = total(SUB_ETC32)
    # PPT 6장 세부내역에는 **지급이자 행이 없습니다**(원본 템플릿 36줄 확인, 2026-08-24).
    # 3p 시트에는 25행에 따로 있지만, 이 표에서는 '기타 > 기타'에 포함시켜야
    # 지출 소계와 눈에 보이는 하위 항목의 합이 맞습니다.
    out["expense.etc"] = add(add(total(SUB_ETC28), out["expense.etc32"]),
                             total(["expense.interest"]))
    # PPT의 '기타 > 기타' = 외주용역비 + 지급이자 + 통신비등·복리후생·식당·직원비용·임차료·수수료
    out["expense.etc.rest"] = add(add(total(["expense.outsource"]), out["expense.etc32"]),
                                  total(["expense.interest"]))
    out["expense.total"] = add(add(out["expense.payroll"], total(["expense.marketing"])),
                               add(add(out["expense.tax"],
                                       total(["expense.capex", "expense.invest"])),
                                   out["expense.etc"]))
    out["balance"] = sub(out["income.total"], out["expense.total"])
    out["fund.net"] = sub(total(["fund.release"]), total(["fund.deposit"]))
    out["fx.net"] = sub(total(["fx.in"]), total(["fx.out"]))
    out["finance.net"] = sub(total(["finance.in"]), total(["finance.out"]))
    out["liquid.change"] = add(add(out["balance"], out["fund.net"]),
                               add(out["fx.net"], out["finance.net"]))
    out["opening"] = (open_krw, open_usd)
    out["closing"] = add(out["opening"], out["liquid.change"])
    return out


def build_report(
    actual_txns, plan_txns,
    actual_month: int, plan_month: int,
    open_krw: float, open_usd: float,
    plan_open_krw: float, plan_open_usd: float,
    invest_open_krw: float, invest_open_usd: float,
    fx: float = 0.0,
) -> Report:
    """실적·계획 거래와 기초잔고로 보고서 표 5종을 만듭니다."""
    ra = aggregate(actual_txns, opening_krw=open_krw, opening_usd=open_usd, fx_rate=fx or None)
    rp = aggregate(plan_txns, opening_krw=plan_open_krw, opening_usd=plan_open_usd,
                   fx_rate=fx or None)
    rep = Report(actual_month, plan_month, fx, ra, rp)

    # ── 운용자금 (표9) ────────────────────────────────────────────
    fa = compute_fund_balance_change(actual_txns, invest_open_krw, invest_open_usd)
    fp = compute_fund_balance_change(plan_txns, fa["end_krw"], fa["end_usd"])

    def fund_row(label, sub, a_krw, a_usd, p_krw, p_usd):
        return [label, sub,
                _num(a_krw), _num(a_usd), _num(conv(a_krw, a_usd, fx)),
                _num(p_krw), _num(p_usd), _num(conv(p_krw, p_usd, fx))]

    rep.tables["fund_plan"] = [
        fund_row("A. 기초자금", "", invest_open_krw, invest_open_usd,
                 fa["end_krw"], fa["end_usd"]),
        fund_row("B. 해지(감소)", "", fa["release_krw"], fa["release_usd"],
                 fp["release_krw"], fp["release_usd"]),
        fund_row("C. 금융기관예치", "", fa["deposit_krw"], fa["deposit_usd"],
                 fp["deposit_krw"], fp["deposit_usd"]),
        fund_row("", "재예치", 0, 0, 0, 0),
        fund_row("", "신규예치", fa["deposit_krw"], fa["deposit_usd"],
                 fp["deposit_krw"], fp["deposit_usd"]),
        fund_row("D. 운용자금 증감 (=C-B)", "",
                 fa["deposit_krw"] - fa["release_krw"], fa["deposit_usd"] - fa["release_usd"],
                 fp["deposit_krw"] - fp["release_krw"], fp["deposit_usd"] - fp["release_usd"]),
        fund_row("E. 기말자금 (=A+D)", "", fa["end_krw"], fa["end_usd"],
                 fp["end_krw"], fp["end_usd"]),
    ]

    # ── 1. 자금수지 (표15) — 전체 자금 = 유동 + 운용 ──────────────
    a_open_total = conv(open_krw, open_usd, fx) + conv(invest_open_krw, invest_open_usd, fx)
    a_in, a_out = ra.values["income.total"], ra.values["expense.total"]
    a_net = a_in - a_out
    a_fin = ra.values["finance.net"]
    a_close_total = a_open_total + a_net + a_fin

    p_open_total = conv(plan_open_krw, plan_open_usd, fx) + conv(fa["end_krw"], fa["end_usd"], fx)
    p_in, p_out = rp.values["income.total"], rp.values["expense.total"]
    p_net = p_in - p_out
    p_fin = rp.values["finance.net"]
    p_close_total = p_open_total + p_net + p_fin

    def flow(label, a, b):
        return [label, _num(a), _num(b), _num(b - a)]

    rep.tables["cash_flow"] = [
        flow("A. 기초잔고", a_open_total, p_open_total),
        flow("B. 수입", a_in, p_in),
        flow("C. 지출", a_out, p_out),
        flow("D. 자금수지차 (=B-C)", a_net, p_net),
        flow("E. 자금조달", a_fin, p_fin),
        flow("F. 기말잔고 (=A+D+E)", a_close_total, p_close_total),
    ]

    # ── 세부 성분(원화/외화 분리) — 아래 표들이 공유합니다 ────────
    # ※ ra.closing / rp.closing 은 '원화 + 외화×환율'로 **이미 환산된 값**입니다.
    #    이걸 '원화' 칸에 넣고 다시 환산하면 외화가 두 번 더해집니다(과거 버그).
    ka = _components(ra, open_krw, open_usd)
    kp = _components(rp, plan_open_krw, plan_open_usd)

    # ── 2. 유동자금 계획 (표10) — 계획 월 ─────────────────────────
    def liq(label, krw, usd):
        return [label, _num(krw), _num(usd), _num(conv(krw, usd, fx))]

    def liqk(label, key):
        krw, usd = kp.get(key, (0.0, 0.0))
        return liq(label, krw, usd)

    p_bal_krw, p_bal_usd = kp["balance"]
    p_close_krw, p_close_usd = kp["closing"]
    rep.tables["liquid_plan"] = [
        liq("A. 기초 유동잔고", plan_open_krw, plan_open_usd),
        liqk("B. 수입", "income.total"),
        liqk("C. 지출", "expense.total"),
        liq("D. 자금수지차 (=B-C)", p_bal_krw, p_bal_usd),
        liq("E. 자금과부족 (=A+D)", plan_open_krw + p_bal_krw, plan_open_usd + p_bal_usd),
        liq("F. 운용자금 활용(해지)", fp["release_krw"], fp["release_usd"]),
        # G는 'I + 예치'로 계산합니다 — 그래야 G-H = I 가 항상 맞고,
        # I(기말 유동잔고)가 세부내역 G행·stock A행과 정확히 같은 값이 됩니다.
        liq("G. 예치 전 기말 유동잔고", p_close_krw + fp["deposit_krw"],
            p_close_usd + fp["deposit_usd"]),
        liq("H. 예치 자금", fp["deposit_krw"], fp["deposit_usd"]),
        liq("I. 기말 유동잔고", p_close_krw, p_close_usd),
    ]

    # ── 4. 월말 자금 stock (표13) ─────────────────────────────────
    def stock(label, sub, a_v, b_v):
        return [label, sub, _num(a_v), _num(b_v), _num(b_v - a_v)]

    a_liq_krw, a_liq_usd = ka["closing"]
    b_liq_krw, b_liq_usd = kp["closing"]
    rep.tables["stock"] = [
        stock("A. 유동자금", "원화", a_liq_krw, b_liq_krw),
        stock("", "외화($)", a_liq_usd, b_liq_usd),
        stock("환산 계", "", conv(a_liq_krw, a_liq_usd, fx), conv(b_liq_krw, b_liq_usd, fx)),
        stock("B. 운용자금", "원화", fa["end_krw"], fp["end_krw"]),
        stock("", "외화($)", fa["end_usd"], fp["end_usd"]),
        stock("환산 계", "", conv(fa["end_krw"], fa["end_usd"], fx),
              conv(fp["end_krw"], fp["end_usd"], fx)),
        stock("C. 기말자금 (=A+B)", "원화", a_liq_krw + fa["end_krw"], b_liq_krw + fp["end_krw"]),
        stock("", "외화($)", a_liq_usd + fa["end_usd"], b_liq_usd + fp["end_usd"]),
        stock("환산 계", "",
              conv(a_liq_krw + fa["end_krw"], a_liq_usd + fa["end_usd"], fx),
              conv(b_liq_krw + fp["end_krw"], b_liq_usd + fp["end_usd"], fx)),
    ]

    # ── 6. 유동자금 계획 세부내역 (PPT 6장과 같은 36줄) ───────────
    detail = []
    for depth, label, key in DETAIL_ROWS:
        a_krw, a_usd = ka.get(key, (0.0, 0.0))
        p_krw, p_usd = kp.get(key, (0.0, 0.0))
        a_c, p_c = conv(a_krw, a_usd, fx), conv(p_krw, p_usd, fx)
        detail.append(["　" * depth + label,
                       _num(a_krw), _num(a_usd), _num(a_c),
                       _num(p_krw), _num(p_usd), _num(p_c), _num(p_c - a_c)])
    rep.tables["detail"] = detail

    # 표1(자금수지)·표2(유동자금 계획)에는 FX 행이 없습니다. FX가 0이 아니면
    # 그 표의 기말잔고가 세부내역 G행과 달라지므로 화면에 알려줍니다.
    structure_warnings = []
    for who, r in (("실적", ra), ("계획", rp)):
        if abs(r.values.get("fx.net", 0.0)) > 1e-9:
            structure_warnings.append(
                f"{who}에 FX(외환매도·매입) 거래가 있습니다. "
                "'1. 자금수지'와 '2. 유동자금 계획' 표에는 FX 행이 없어 "
                "기말잔고가 세부내역과 다르게 보일 수 있습니다.")

    rep.meta = {
        "structure_warnings": structure_warnings,
        "actual_close_total": a_close_total,
        "plan_close_total": p_close_total,
        "actual_liquid_close": ra.closing,
        "plan_liquid_close": rp.closing,
        "invest_actual_end": fa["end_krw"],
        "invest_plan_end": fp["end_krw"],
        "diff_total": p_close_total - a_close_total,
    }
    return rep


# ── 화면·PPT 공용 헤더 ────────────────────────────────────────────
def headers(rep: Report) -> dict[str, list[str]]:
    """표 머리글. 판다스가 같은 이름을 싫어해서 전부 다르게 붙입니다."""
    a, p = rep.actual_month, rep.plan_month
    return {
        "cash_flow": ["구분", f"{a}월(a)", f"{p}월(b)", "b - a"],
        "liquid_plan": ["구분", "원화", "외화($)", "환산 계"],
        "fund_plan": ["구분", "세부", f"{a}월 원화", f"{a}월 외화($)", f"{a}월 환산 계",
                      f"{p}월 원화", f"{p}월 외화($)", f"{p}월 환산 계"],
        "stock": ["구분", "세부", f"{a}월(a)", f"{p}월(b)", "b - a"],
        "detail": ["구분", f"{a}월 원화", f"{a}월 외화($)", f"{a}월 환산계(a)",
                   f"{p}월 원화", f"{p}월 외화($)", f"{p}월 환산계(b)", "b-a"],
    }


# 표별 비고 줄 수 (사용자가 직접 쓰는 칸)
REMARK_ROWS = {"cash_flow": 6, "liquid_plan": 9, "fund_plan": 7,
               "stock": 9, "detail": 36}


# 표 이름은 **보고 대상 월(계획월)** 기준입니다.
# 원본 PPT도 '7월 자금운용계획(안)' 안에 '1. 7월 자금수지' 로 적혀 있습니다
# (표 안에는 6월(a)·7월(b) 두 칸이 다 있지만, 제목은 계획월을 씁니다).
TABLE_TITLE = {
    "cash_flow": "1. {p}월 자금수지",
    "liquid_plan": "2. {p}월 유동자금 계획",
    "fund_plan": "3. 운용자금 계획",
    "stock": "4. {p}월말 자금(stock)",
    "detail": "5. 유동자금 계획 세부내역",
}
