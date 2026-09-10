"""원본 워크북 실측값으로 계산 결과를 검증합니다.

실행:  python tests/test_engine.py
      (templates/자금운용계획_template.xlsm 이 있어야 전체 검증이 돌아갑니다)
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl

from core.config import TEMPLATE_PATH
from core.data_store import Txn, init_db
from core.engine import aggregate, compute_fund_balance_change, validate
from core.remarks_rule import generate_leaf_remarks
from core.tag_classifier import suggest_tag

init_db()

# ── 계정 자동분류 (템플릿 없이도 확인 가능) ──
for memo, expect in [("사무실 임차료 6월분", "지급임차료"), ("임직원 급여", "급여"),
                     ("정기예금 300백만 만기해지", "금융상품"),
                     ("1기 확정 부가가치세 납부", "부가가치세")]:
    s = suggest_tag(Txn("2026-06", "plan", memo=memo, amount_out=1))
    assert s.tag == expect, (memo, s.tag)
print("✓ 계정 자동분류 정상")

if not TEMPLATE_PATH.exists():
    print(f"\n(templates 에 원본 워크북이 없어 나머지 검증은 건너뜁니다: {TEMPLATE_PATH})")
    sys.exit(0)

wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)


def load(sheet, kind):
    return [Txn(period="2026-06", kind=kind, tag=str(r[0]).strip(),
                counterpart=str(r[1] or ""), memo=str(r[2] or ""),
                currency=str(r[3] or "KRW"),
                amount_in=float(r[4] or 0), amount_out=float(r[5] or 0))
            for r in wb[sheet].iter_rows(min_row=4, values_only=True) if r[0]]


actual, plan = load("6월데이터", "actual"), load("7월데이터", "plan")
res = aggregate(actual, opening_krw=88.757327)
assert abs(res.closing - 362.495882) < 1e-6, res.closing
print(f"✓ 6월 실적 기말잔고 재현 — {res.closing:.6f} (원본 362.495882)")

resp = aggregate(plan, opening_krw=res.closing)
assert abs(resp.closing - 318.8615828) < 1e-6, resp.closing
print(f"✓ 7월 계획 기말잔고 재현 — {resp.closing:.6f} (원본 318.8615828)")

assert abs(res.values["income.total"] - 6.535533) < 1e-6
assert abs(res.values["expense.total"] - 32.796978) < 1e-6
print("✓ 수입·지출 소계 일치")

for g in validate(actual, res):
    assert g.passed, (g.id, g.detail, g.items)
print("✓ 검증 게이트 6종 통과")

ch = compute_fund_balance_change(actual, 2400.0, 0.0)
assert abs(ch["release_krw"] - 300.0) < 1e-9 and abs(ch["end_krw"] - 2100.0) < 1e-9
print("✓ 운용자금 이월 정상 (기초 2400 - 해지 300 = 기말 2100)")

r = generate_leaf_remarks(res, 6)
assert "급여" in r["expense.salary"]
print("✓ 비고 자동 작성 정상 —", r["expense.salary"])

print("\n전부 통과했습니다.")
