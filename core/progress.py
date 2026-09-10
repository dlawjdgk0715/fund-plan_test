"""지금 어디까지 했는지 — 4단계 진행 현황.

**따로 기록하지 않습니다.** 사용자가 체크할 필요도 없습니다.
이미 저장돼 있는 것들에서 매번 계산합니다.

    1 실적 업로드     ← transactions (실적)
    2 계획 업로드     ← transactions (계획)
    3 실적/계획 관리  ← 계정과목이 비었거나 사전에 없는 거래 수
    4 최종 점검·산출  ← 검증 게이트 + 만들어 둔 산출물

따로 로그를 남기면 실제 자료와 어긋날 수 있어서, **계산으로만** 알아냅니다.
"""
from __future__ import annotations

from dataclasses import dataclass

from core.data_store import (load_gate_acks, load_month_balance,
                             load_transactions, resolve_open_invest,
                             txn_fingerprint)
from core.engine import aggregate, validate


@dataclass
class Step:
    no: int
    label: str          # 화면 이름
    page: str           # 이동할 파일
    done: bool
    detail: str         # 끝났으면 요약, 아니면 '무엇이 남았는지'


def auto_vals(period: str) -> dict:
    """4번 화면을 안 거쳤을 때 쓰는 기본 잔고·환율."""
    mb = load_month_balance(period)
    ik, iu, _src = resolve_open_invest(period)
    return {
        "open_krw": float(mb["open_krw"]) if mb else 0.0,
        "open_usd": float(mb["open_usd"]) if mb else 0.0,
        "plan_open_krw": float(mb["close_krw"]) if mb else 0.0,
        "plan_open_usd": float(mb["close_usd"]) if mb else 0.0,
        "invest_open_krw": ik, "invest_open_usd": iu, "fx": 0.0,
    }


def _need_check(txns) -> int:
    """사람이 봐야 할 거래 수 — 셋을 더합니다.

      ① 계정과목이 **비어 있는** 거래 (집계에서 조용히 빠집니다)
      ② 사전에 없는 이름       (unmapped)
      ③ 이름은 맞는데 방향 조합이 없는 것 (unregistered)

    ※ ①은 `aggregate` 가 그냥 건너뛰기 때문에 여기서 따로 셉니다.
    """
    if not txns:
        return 0
    blank = sum(1 for t in txns if not t.excluded and not (t.tag or "").strip())
    res = aggregate(txns)
    return blank + len(res.unmapped) + len(res.unregistered)


def _red_gates(period: str, vals: dict) -> list[str]:
    """아직 빨간 검증 게이트 이름들 ([확인함] 누른 것은 뺍니다)."""
    out = []
    for scope, kind in (("actual", "actual"), ("plan", "plan")):
        txns = load_transactions(period, kind)
        if not txns:
            continue
        fp = txn_fingerprint(txns)
        acks = load_gate_acks(period, scope)
        opening = (vals["open_krw"] if kind == "actual" else vals["plan_open_krw"])
        res = aggregate(txns, opening_krw=float(opening or 0.0),
                        fx_rate=float(vals.get("fx") or 0) or None)
        ext = None
        if kind == "actual":
            mb = load_month_balance(period)
            ext = float(mb["close_krw"]) if mb else None
        for g in validate(txns, res, external_balance=ext,
                          fx_rate=float(vals.get("fx") or 0) or None):
            if g.passed:
                continue
            if acks.get(g.id, {}).get("fingerprint") == fp:
                continue
            out.append(f"{g.id}")
    return out


def steps(period: str, vals: dict | None = None,
          output_ready: bool | None = None) -> list[Step]:
    """4단계 진행 현황. output_ready = 최신 산출물이 있는지(시작 화면에서 넘겨줌)."""
    vals = vals or auto_vals(period)
    actual = load_transactions(period, "actual")
    plan = load_transactions(period, "plan")

    sheets = len({t.txn_date for t in actual if t.txn_date})
    depts = sorted({t.dept for t in plan if t.dept})

    s1 = Step(1, "1  실적 업로드", "pages/1_Upload_Actual.py", bool(actual),
              f"{len(actual):,}건 · 시트 {sheets}개" if actual
              else "일일자금계획 파일을 올려주세요")

    if plan:
        who = ", ".join(depts[:4]) + (f" 외 {len(depts) - 4}곳" if len(depts) > 4 else "")
        d2 = f"{len(plan):,}건 · 부서 {len(depts)}곳" + (f" ({who})" if who else "")
    else:
        d2 = "부서 자금집행계획서를 올려주세요"
    s2 = Step(2, "2  계획 업로드", "pages/2_Upload_Plan.py", bool(plan), d2)

    n_chk = _need_check(actual) + _need_check(plan)
    s3 = Step(3, "3  실적/계획 관리", "pages/3_Manage_Data.py", n_chk == 0,
              "계정과목 정리 끝" if n_chk == 0 else f"계정과목 확인 필요 {n_chk}건")

    reds = _red_gates(period, vals) if (actual or plan) else []
    if not (actual or plan):
        d4, ok4 = "1·2번을 먼저 해주세요", False
    elif reds:
        d4, ok4 = f"확인할 검증 {len(reds)}개 ({', '.join(reds)})", False
    elif output_ready is False:
        d4, ok4 = "수치·비고 확인 완료 — 엑셀·PPT를 만들어주세요", False
    else:
        d4, ok4 = "검증 통과 · 산출물 준비됨", True
    s4 = Step(4, "4  최종 점검 · 산출", "pages/4_Preview_Report.py", ok4, d4)

    return [s1, s2, s3, s4]


def first_todo(items: list[Step]) -> Step | None:
    for s in items:
        if not s.done:
            return s
    return None
