"""비고 자동 작성 — 방식 A(규칙 기반). 정해진 문장 틀에 실제 거래 값만 끼워 넣습니다.

v0.11 개선 반영:
  1) 계정과목당 최대 3건까지만(금액 큰 순), 나머지는 조용히 생략
  2) 적요를 그대로 쓰지 않고 짧은 라벨로 축약(DB 매핑표)
  3) 같은 라벨끼리 금액 합산해서 한 줄로

v0.14(2026-08-25):
  4) '쓰는 방식'(금액 단위·소수 자릿수·항목 개수·구분 기호)은 core/remark_style 이
     사용자의 수정에서 학습한 값을 씁니다. 여기서는 **내용만** 만듭니다.
"""
from __future__ import annotations

import re

from core.config import MILLION
from core.data_store import Txn, load_remark_style
from core.engine import EXCLUDED_ACCOUNTS, Result
from core.remark_labels import get_label
from core.remark_style import render_items, style_for

MAX_LINES = 3


def _fmt(v: float) -> str:
    """백만원 단위 표시 — 소수점 1자리, 끝의 .0은 생략."""
    s = f"{v:,.1f}"
    return s[:-2] if s.endswith(".0") else s


def _style(table_key: str, style: dict | None):
    if style is not None:
        return style
    try:
        return style_for(table_key, load_remark_style(table_key))
    except Exception:          # DB가 아직 없을 때(테스트 등)
        return style_for(table_key)


def _pairs(items: dict[str, float]) -> list[tuple[str, float]]:
    return sorted(items.items(), key=lambda kv: -abs(kv[1]))


def generate_leaf_remarks(result: Result, month: int,
                          style: dict | None = None) -> dict[str, str]:
    """유동자금세부 N열 — row_key별 비고 문자열.

    기본 형식:  ∙{월} : {라벨} {금액}백만   (건별로 줄바꿈)
    """
    st = _style("detail", style)
    out: dict[str, str] = {}
    for row_key, txns in result.rows_by_key.items():
        merged: dict[str, float] = {}
        for t in txns:
            label = get_label(t.tag, t.memo, row_key)
            merged[label] = merged.get(label, 0.0) + t.amount / MILLION
        text = render_items(_pairs(merged), month, st)
        if text:
            out[row_key] = text
    return out


def generate_summary_draft(txns: list[Txn], month: int, direction: str,
                           table_key: str = "summary1",
                           style: dict | None = None) -> str:
    """요약1의 서술형 비고 초안 — 금액 큰 상위 3건을 한 줄로.

    기본 형식:  {월} : {라벨1} {금액1}백만, {라벨2} {금액2}백만 등
    해당 거래가 없으면 빈 문자열(사람이 직접 채움).
    """
    st = _style(table_key, style)
    picked = [t for t in txns
              if (not t.excluded) and t.direction == direction and t.amount > 0
              and (t.tag or "").strip() not in EXCLUDED_ACCOUNTS]
    if not picked:
        return ""
    merged: dict[str, float] = {}
    for t in picked:
        label = get_label(t.tag, t.memo)
        merged[label] = merged.get(label, 0.0) + t.amount / MILLION
    return render_items([kv for kv in _pairs(merged) if kv[1] >= 0.05], month, st)


def join_lines(*parts: str) -> str:
    return "\n".join(p for p in parts if p)


_join = join_lines          # 기존 이름 유지

# 요약1(엑셀 1-1)·자금수지 표에 들어가는 문장들. 화면·엑셀·PPT가 이 순서를 공유합니다.
# 원본 PPT 서식 그대로 ● = 실적월 / ○ = 계획월 기호를 붙입니다.
#   ● 6월 : 정기예금 만기해지 이자 6.4백만 등
#   ○ 7월 : 급여 15백만, 부가세 납부 3.7백만 등
#   ○ 기초잔고: 유동자금 362.5백만, 운용자금 2,100백만
# ※ 뒤의 두 개(open_bal·close_bal)는 나중에 추가한 것이라 **맨 뒤**에 붙였습니다.
#    (앞 네 개의 저장 위치가 바뀌면 이미 저장된 글이 뒤섞입니다)
SUMMARY1_KEYS = ("actual_in", "plan_in", "actual_out", "plan_out",
                 "open_bal", "close_bal")
SUMMARY1_ORDER = ("open_bal", "actual_in", "plan_in", "actual_out", "plan_out",
                  "close_bal")          # 화면에 보여줄 순서 (표 순서와 같게)
SUMMARY1_LABEL = {
    "open_bal": "○ A. 기초잔고",
    "actual_in": "● {a}월 실적 수입", "plan_in": "○ {p}월 계획 수입",
    "actual_out": "● {a}월 실적 지출", "plan_out": "○ {p}월 계획 지출",
    "close_bal": "○ F. 기말잔고",
}
SUMMARY1_MARK = {"actual_in": "●", "actual_out": "●",
                 "plan_in": "○", "plan_out": "○",
                 "open_bal": "○", "close_bal": "○"}
_HAS_MARK = re.compile(r"^\s*[●○]")


def mark_summary1(key: str, body: str) -> str:
    """원본 PPT 서식대로 ● / ○ 기호를 붙입니다 (이미 있으면 그대로)."""
    body = (body or "").strip()
    if not body or _HAS_MARK.match(body):
        return body
    return f"{SUMMARY1_MARK.get(key, '○')} {body}"


def _mil(v: float) -> str:
    s = f"{v:,.1f}"
    return s[:-2] if s.endswith(".0") else s


def balance_notes(rep) -> dict[str, str]:
    """기초·기말잔고 문장 — 원본 PPT 서식 그대로.

    '○ 기초잔고: 유동자금 362.5백만, 운용자금 2,100백만'
    기준은 **계획월**입니다(기초 = 실적월 말, 기말 = 계획월 말).
    """
    try:
        stk = rep.tables["stock"]        # 2행 = 유동 환산계, 5행 = 운용 환산계
        return {
            "open_bal": f"○ 기초잔고: 유동자금 {_mil(stk[2][2])}백만, "
                        f"운용자금 {_mil(stk[5][2])}백만",
            "close_bal": f"○ 기말잔고: 유동자금 {_mil(stk[2][3])}백만, "
                         f"운용자금 {_mil(stk[5][3])}백만",
        }
    except Exception:
        return {"open_bal": "", "close_bal": ""}


def summary1_drafts(actual_txns, plan_txns, actual_month: int, plan_month: int,
                    style: dict | None = None, rep=None) -> dict[str, str]:
    """요약1 서술형 비고 초안 (계정과목별 큰 금액 순 최대 3건 + 잔고 두 줄)."""
    st = _style("summary1", style)
    out = {
        "actual_in": generate_summary_draft(actual_txns, actual_month, "in", "summary1", st),
        "plan_in": generate_summary_draft(plan_txns, plan_month, "in", "summary1", st),
        "actual_out": generate_summary_draft(actual_txns, actual_month, "out", "summary1", st),
        "plan_out": generate_summary_draft(plan_txns, plan_month, "out", "summary1", st),
        "open_bal": "", "close_bal": "",
    }
    if rep is not None:
        out.update(balance_notes(rep))
    return {k: mark_summary1(k, v) for k, v in out.items()}


def cashflow_rows_from_summary1(notes: dict[str, str]) -> list[str]:
    """문장들 → '1. 자금수지' 표(6줄)의 비고.

    요약1·PPT 2장 표가 **같은 문장**을 쓰도록, 사용자는 여기 문장만 손보면 됩니다.
    줄 순서: A.기초잔고 / B.수입 / C.지출 / D.수지차 / E.조달 / F.기말잔고
    """
    def g(k):
        return mark_summary1(k, notes.get(k, ""))
    return [g("open_bal"),
            join_lines(g("actual_in"), g("plan_in")),
            join_lines(g("actual_out"), g("plan_out")),
            "", "", g("close_bal")]


def summary1_from_cashflow(rows: list[str]) -> dict[str, str]:
    """예전에 '자금수지' 표에 직접 써 두신 비고를 문장들로 되돌립니다.

    (7차 이전 버전에서 고치신 글이 요약1로 안 넘어가던 문제를 메웁니다)
    """
    def split(text: str) -> tuple[str, str]:
        parts = [x.strip() for x in re.split(r"\n| / ", text or "") if x.strip()]
        act = next((x for x in parts if x.startswith("●")), "")
        pln = next((x for x in parts if x.startswith("○")), "")
        if not act and not pln:                    # 기호가 없으면 첫 줄=실적, 둘째 줄=계획
            act = parts[0] if parts else ""
            pln = parts[1] if len(parts) > 1 else ""
        return act, pln

    rows = list(rows) + [""] * (6 - len(rows))
    a_in, p_in = split(rows[1])
    a_out, p_out = split(rows[2])
    return {"actual_in": a_in, "plan_in": p_in, "actual_out": a_out, "plan_out": p_out,
            "open_bal": (rows[0] or "").strip(), "close_bal": (rows[5] or "").strip()}


def cashflow_drafts(actual_txns, plan_txns, actual_month: int, plan_month: int,
                    style: dict | None = None, rep=None) -> list[str]:
    """'1. 자금수지' 표의 비고 초안 6줄."""
    return cashflow_rows_from_summary1(
        summary1_drafts(actual_txns, plan_txns, actual_month, plan_month, style, rep))


def detail_drafts(actual_result: Result, plan_result: Result,
                  actual_month: int, plan_month: int,
                  style: dict | None = None) -> list[str]:
    """'5. 유동자금 계획 세부내역' 표의 비고 초안 (표의 줄 순서 그대로)."""
    from core.report import DETAIL_ROWS

    st = _style("detail", style)
    la = generate_leaf_remarks(actual_result, actual_month, st)
    lp = generate_leaf_remarks(plan_result, plan_month, st)
    return [_join(la.get(key, ""), lp.get(key, "")) for _, _, key in DETAIL_ROWS]
