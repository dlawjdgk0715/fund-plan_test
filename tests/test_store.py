"""영구 저장소(깃허브) + 연간 시트 월별 보관 검증 (2026-11).

실행:  python tests/test_store.py

여기서 보는 것
  1) 연간 월별 값 저장·조회·덮어쓰기
  2) 빈 DB 를 만들면 `db/annual_seed.json` 에서 8·9월이 되살아나는지
  3) 연간 시트 복원 — 지난 달은 채우고, 이번 달은 건드리지 않는지
  4) 빈 달 경고 — 같은 사고(8·9월이 조용히 비는 것)가 다시 나면 잡히는지
  5) 깃허브 저장소 — 설정이 없을 때 조용히 꺼지는지(내 PC 사용에 지장 없음)
  6) 저장하면 '올릴 것 있음' 표시가 켜지는지
"""
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
ROOT = Path(__file__).resolve().parents[1]

# 실제 업무 DB 를 절대 건드리지 않습니다. core.config import 전에 정해야 합니다.
_TMP = Path(tempfile.mkdtemp(prefix="fundplan_store_"))
os.environ["FUNDPLAN_DB"] = str(_TMP / "tags.db")
for _k in ("FUNDPLAN_GH_TOKEN", "FUNDPLAN_GH_REPO",
           "FUNDPLAN_GH_BRANCH", "FUNDPLAN_GH_PATH"):
    os.environ.pop(_k, None)

n_ok = 0


def ok(msg):
    global n_ok
    n_ok += 1
    print("✓", msg)


from core import gitstore                                          # noqa: E402
from core.data_store import (annual_saved_months, init_db,         # noqa: E402
                             load_annual_year, save_annual_month)

init_db()

# ── 1. 저장·조회·덮어쓰기 ─────────────────────────────────────────
save_annual_month(2099, 3, {"income.sales": 1.5, "expense.salary": 7.0},
                  {"income.sales": 0.25}, kind="actual")
h = load_annual_year(2099)
assert h[3]["krw"]["income.sales"] == 1.5
assert h[3]["usd"]["income.sales"] == 0.25
assert h[3]["kind"] == "actual"
ok("연간 월별 값 저장·조회")

save_annual_month(2099, 3, {"income.sales": 9.0}, {}, kind="plan")
h = load_annual_year(2099)
assert h[3]["krw"]["income.sales"] == 9.0, "같은 달은 덮어써야 합니다"
assert h[3]["kind"] == "plan"
ok("같은 달 다시 저장하면 덮어쓰기 (중복 행 안 생김)")

assert annual_saved_months(2099) == [3]
ok("보관된 달 목록")

# ── 2. 씨앗 복구 — 8·9월이 되살아나는지 ──────────────────────────
saved = annual_saved_months(2026)
assert 8 in saved and 9 in saved, f"8·9월이 씨앗에서 안 들어옴: {saved}"
y26 = load_annual_year(2026)
assert y26[8]["krw"]["expense.salary"] > 0, "8월 급여가 0입니다"
assert y26[9]["krw"]["income.sales"] > 0, "9월 매출이 0입니다"
ok(f"빈 DB 에서 씨앗 복구 — {len(saved)}개 달 (8·9월 포함)")

_before = len(load_annual_year(2026)[8]["krw"])
init_db()                                    # 두 번 불러도 덮어쓰지 않아야 합니다
assert len(load_annual_year(2026)[8]["krw"]) == _before
ok("재초기화해도 씨앗이 기존 값을 덮지 않음")

# ── 3·4. 연간 시트 복원 + 빈 달 경고 ─────────────────────────────
import openpyxl                                                    # noqa: E402
from openpyxl.utils import column_index_from_string as CI          # noqa: E402

from core.excel_writer import (ANNUAL_LEAF_ROW, ANNUAL_SHEET_COL,  # noqa: E402
                               SRC_ANNUAL, _restore_annual_history,
                               _warn_empty_annual_months, annual_values)

wb = openpyxl.Workbook()
ws = wb.create_sheet(SRC_ANNUAL)
rows = list(ANNUAL_LEAF_ROW.values())


def filled(month):
    c = CI(ANNUAL_SHEET_COL[month])
    return sum(1 for r in rows if ws.cell(row=r, column=c).value is not None)


hist = {4: {"krw": {k: 1.0 for k in ANNUAL_LEAF_ROW},
            "usd": {k: 0.0 for k in ANNUAL_LEAF_ROW}, "kind": "actual"},
        5: {"krw": {k: 2.0 for k in ANNUAL_LEAF_ROW},
            "usd": {k: 0.0 for k in ANNUAL_LEAF_ROW}, "kind": "actual"}}
warns = []
_restore_annual_history(wb, hist, warns, skip={5})
assert filled(4) == len(rows), "복원 대상 달이 안 채워졌습니다"
assert filled(5) == 0, "이번에 채울 달(skip)은 건드리면 안 됩니다"
ok("지난 달 복원 — 대상 달만 채우고 이번 달은 그대로 둠")

warns = []
_warn_empty_annual_months(wb, 6, 7, warns)
assert warns and "비어 있습니다" in warns[0], "빈 달 경고가 안 나왔습니다"
assert "5월" in warns[0], f"빈 달 목록이 틀립니다: {warns[0]}"
assert "4월" not in warns[0], "채워진 달까지 경고에 들어갔습니다"
ok("빈 달 경고 — 비어 있는 달만 집어냄")

full = {m: {"krw": {k: 1.0 for k in ANNUAL_LEAF_ROW},
            "usd": {k: 0.0 for k in ANNUAL_LEAF_ROW}, "kind": "actual"}
        for m in range(1, 6)}
_restore_annual_history(wb, full, [], skip=set())
warns = []
_warn_empty_annual_months(wb, 5, 5, warns)
assert not warns, f"다 채웠는데 경고가 났습니다: {warns}"
ok("빈 달이 없으면 경고 없음")


# ── 4-b. 합계·평균 열의 빠진 수식 보정 ───────────────────────────
from core.excel_writer import _fix_annual_total_formulas            # noqa: E402

wb2 = openpyxl.Workbook()
ws2 = wb2.create_sheet(SRC_ANNUAL)
ws2.cell(row=4, column=CI("AQ")).value = "합계"
ws2.cell(row=4, column=CI("AT")).value = "평균"
# 원본 워크북에 **이미 수식이 들어 있는** 행들 (마케팅비·지급이자·자산매입·투자금·
# 운용자금·FX·조달) — 이 칸은 건드리면 안 됩니다.
_HAVE = (19, 25, 26, 27, 34, 35, 37, 38, 40, 41)
for _r in _HAVE:
    ws2.cell(row=_r, column=CI("AQ")).value = f"=SUM(G{_r},J{_r})"
warns = []
_fix_annual_total_formulas(wb2, warns)
f9 = ws2.cell(row=9, column=CI("AQ")).value
assert f9 and f9.startswith("=SUM(G9,J9,"), f"합계 수식이 안 들어갔습니다: {f9}"
assert f9.endswith("AN9)"), f"12월까지 더해야 합니다: {f9}"
for _r in _HAVE:
    assert ws2.cell(row=_r, column=CI("AQ")).value == f"=SUM(G{_r},J{_r})", \
        f"원래 있던 수식을 덮어썼습니다 ({_r}행)"
assert warns and "16칸" in warns[0], warns
ok("연간 합계열 — 비어 있던 16칸만 채우고 기존 수식은 그대로")

wb3 = openpyxl.Workbook()
wb3.create_sheet(SRC_ANNUAL)                 # '합계' 머리글이 없는 경우
warns = []
_fix_annual_total_formulas(wb3, warns)
assert warns and "못 찾아" in warns[0]
ok("합계 열을 못 찾으면 건너뛰고 알려줌")


class _R:
    krw = {"income.sales": 3.0, "expense.salary": 4.0}
    usd = {"income.sales": 0.5}


k, u = annual_values(_R())
assert k["income.sales"] == 3.0 and u["income.sales"] == 0.5
assert set(k) == set(ANNUAL_LEAF_ROW), "모든 항목이 나와야 합니다(없으면 0)"
assert k["expense.capex"] == 0.0
ok("연간 한 달치 값 뽑기 — 빠진 항목은 0")

# ── 5. 깃허브 저장소: 설정이 없으면 조용히 꺼짐 ──────────────────
assert gitstore.enabled() is False
assert gitstore.pull()["ok"] is None
assert gitstore.push()["ok"] is None
gitstore.mark_dirty("테스트")          # 아무 일도 없어야 합니다
ok("저장소 설정이 없으면 아무 것도 안 함 (내 PC 사용에 지장 없음)")

cfg = gitstore.config()
assert cfg["branch"] == "data" and cfg["path"].endswith(".db")
assert "/" in cfg["repo"]
ok(f"기본 설정 — 브랜치 `{cfg['branch']}`, 경로 `{cfg['path']}`")

stt = gitstore.status()
assert stt["enabled"] is False and "repo" in stt
ok("상태 조회 (시작 화면 표시용)")

# ── 6. 저장하면 '올릴 것 있음' 이 켜지는지 ───────────────────────
os.environ["FUNDPLAN_GH_TOKEN"] = "test-token-not-used"
try:
    assert gitstore.enabled() is True
    save_annual_month(2099, 7, {"income.sales": 1.0}, {})
    assert gitstore.status()["dirty"] is True, "저장했는데 표시가 안 켜졌습니다"
    ok("저장하면 '저장소에 올릴 것 있음' 표시가 켜짐")
finally:
    os.environ.pop("FUNDPLAN_GH_TOKEN", None)
    gitstore.flush  # noqa: B018  (타이머는 데몬이라 그냥 둬도 됩니다)

seed = ROOT / "db" / "annual_seed.json"
assert seed.exists() and seed.stat().st_size > 1000
ok("복구용 씨앗 파일이 리포에 들어 있음")


# ── 7. 서식기준 파일 — .xlsm 과 .xlsx 둘 다 받는지 ───────────────
from core.excel_writer import style_ref_path                       # noqa: E402

_tpl = style_ref_path().parent
_xlsm = _tpl / "서식기준.xlsm"
_xlsx = _tpl / "서식기준.xlsx"
# 실제 쓰시는 파일이 있을 수 있으니 잠시 치워 두고, 끝나면 되돌립니다.
_baks = {}
for _p in (_xlsm, _xlsx):
    if _p.exists():
        _b = _p.with_suffix(_p.suffix + ".bak")
        _p.rename(_b)
        _baks[_p] = _b
try:
    assert style_ref_path().name == "서식기준.xlsm", "없을 땐 .xlsm 이름으로 안내해야 합니다"
    _xlsx.write_bytes(b"x")
    assert style_ref_path().name == "서식기준.xlsx", \
        f".xlsx 를 못 찾았습니다: {style_ref_path().name}"
    ok(".xlsx 로 저장한 서식기준 파일도 찾음 (예전에는 조용히 무시됐습니다)")
    _xlsm.write_bytes(b"x")
    assert style_ref_path().name == "서식기준.xlsm", "둘 다 있으면 .xlsm 이 먼저"
    ok("둘 다 있으면 .xlsm 우선")
finally:
    for _p in (_xlsx, _xlsm):
        if _p.exists() and _p.stat().st_size == 1:
            _p.unlink()
    for _p, _b in _baks.items():
        if _b.exists():
            _b.rename(_p)

# ── 8. 빈 달 보호 — 자료 없이 산출해도 기록을 덮지 않는지 (정적 검사) ──
_src = (ROOT / "core" / "outputs.py").read_text(encoding="utf-8")
_i = _src.index("save_annual_month(year, actual_month")
_head = _src[:_i]
assert _head.rstrip().endswith("a_krw, a_usd = annual_values(res)"), "가드 구조가 바뀌었습니다"
assert "if actual:" in _head[-300:], \
    "거래가 없을 때 연간 기록을 덮지 않도록 막는 `if actual:` 가 없습니다"
assert "if plan:" in _src[_i:_i + 400], "계획월에도 같은 가드가 필요합니다"
ok("자료 없이 산출해도 그 달 연간 기록을 덮지 않음 (0 으로 지워지는 사고 방지)")


print(f"\n총 {n_ok}개 항목 통과")
