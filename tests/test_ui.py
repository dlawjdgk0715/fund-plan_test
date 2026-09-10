"""4·5번 화면 개선분 검증 (2026-08-25).

실행:  python tests/test_ui.py

여기서 보는 것
  1) 비고 '쓰는 방식' 학습 — 내용이 아니라 틀만 배우는지
  2) 원본 소계 대조표 — 시트가 직접 적어 둔 소계를 그대로 읽는지
  3) 거래 지문 — 한 건만 바뀌어도 값이 달라지는지(확인이 풀려야 함)
  4) 템플릿 자동 탐색 — 파일 이름이 달라도 찾는지
  5) 화면 파일이 없는 이름을 import 하고 있지 않은지(정적 검사)
"""
import ast
import re
import os
import sys
import tempfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ROOT = Path(__file__).resolve().parents[1]

# ※ 실제 업무 DB(db/tags.db)를 절대 건드리지 않도록, 검사용 임시 DB를 씁니다.
#    core.config 를 import 하기 **전에** 정해야 합니다.
_TMPDB = Path(tempfile.mkdtemp(prefix="fundplan_test_")) / "tags.db"
os.environ["FUNDPLAN_DB"] = str(_TMPDB)
n_ok = 0


def ok(msg):
    global n_ok
    n_ok += 1
    print("✓", msg)


# ── 1. 비고 서식 학습 ─────────────────────────────────────────────
from core.remark_style import infer_style, learn, render_items, style_for  # noqa: E402

pairs = [("급여", 149.5), ("퇴직연금", 30.0), ("4대보험", 12.34), ("기타", 1.0)]

base = style_for("summary1")
assert render_items(pairs, 7, base) == "7월 : 급여 149.5백만, 퇴직연금 30백만, 4대보험 12.3백만 등"
ok("기본 서식 — 상위 3건 + '등'")

leaf = style_for("detail")
assert render_items(pairs[:2], 7, leaf) == "∙7월 : 급여 149.5백만\n∙7월 : 퇴직연금 30백만"
ok("세부내역 기본 서식 — 건별 줄바꿈")

# 사용자가 '억 / 소수 2자리 / 2건만 / 슬래시 구분 / 등 없음' 으로 고친 경우
edited = "7월 : 급여 1.50억 / 퇴직연금 0.30억"
st = infer_style(edited, 7, base)
assert st["unit"] == "억", st
assert st["decimals"] == 2, st
assert st["max_items"] == 2, st
assert st["sep"] == " / ", st
assert st["tail"] == "", st
assert st["prefix"] == "{m}월 : ", st
ok("사용자 수정에서 틀만 뽑아냄 (억·소수2·2건·슬래시)")

# 배운 틀을 **다음 달·다른 금액**에 적용해도 형태가 유지되는지
nxt = [("상여", 220.0), ("급여", 150.0), ("기타", 3.0)]
out = render_items(nxt, 8, st)
assert out == "8월 : 상여 2.20억 / 급여 1.50억", out
ok("배운 틀을 다음 달 새 금액에 적용 — 내용은 새로, 형태는 그대로")

# learn(): 안 고친 줄은 무시하고, 고친 줄에서만 배웁니다
drafts = ["", "7월 : 급여 149.5백만", "7월 : 이자수익 3.2백만"]
edits = ["", "7월 : 급여 149.5백만", "7월 : 이자수익 0.03억"]
got = learn("cash_flow", drafts, edits, 7)
assert got and got["unit"] == "억", got
assert learn("cash_flow", drafts, drafts, 7) is None
ok("learn() — 고친 줄에서만 배우고, 안 고쳤으면 아무것도 저장하지 않음")


# ── 2. 원본 소계 대조 ─────────────────────────────────────────────
from core.parsers import daily_sheet_subtotals, parse_daily_sheet  # noqa: E402

rows = [
    ["", "", "1. 일일자금수지 실적"],
    ["통화구분", "거래처명", "내 역", "계정과목", "수입", "지출"],
    ["1)원화", "우리은행", "이자", "지급이자", 0, 1_000_000],
    ["", "카카오", "매출대금", "외상매출금", 5_000_000, 0],
    ["", "", "소계", "", 5_000_000, 1_000_000],
    ["2) 외화(USD)", "AWS", "클라우드", "전산유지비", 0, 2_000],
    ["", "", "소계", "", 0, 2_000],
    ["", "", "2. 일일자금수지 계획"],
    ["통화구분", "거래처명", "내 역", "계정과목", "수입", "지출"],
    ["1)원화", "국세청", "부가세", "부가가치세", 0, 7_000_000],
    ["", "", "소계", "", 0, 7_000_000],
]
subs = daily_sheet_subtotals(rows)
assert len(subs) == 3, subs
assert subs[0] == {"구간": "실적", "통화": "KRW", "수입": 5_000_000, "지출": 1_000_000}, subs[0]
assert subs[1]["통화"] == "USD" and subs[1]["지출"] == 2_000, subs[1]
assert subs[2]["구간"] == "계획" and subs[2]["지출"] == 7_000_000, subs[2]
ok("원본 소계 대조 — 실적/계획·원화/외화 구간별 소계를 그대로 읽음")

acts = parse_daily_sheet(rows, "2026-07", "0731", "실적")
krw = [t for t in acts if t.currency == "KRW"]
assert abs(sum(t.amount_in for t in krw) - subs[0]["수입"]) < 1
assert abs(sum(t.amount_out for t in krw) - subs[0]["지출"]) < 1
ok("앱이 읽은 합계 = 원본 소계 (누락 없음)")


# ── 3. 거래 지문 ──────────────────────────────────────────────────
from core.data_store import Txn, txn_fingerprint  # noqa: E402

a = [Txn("2026-07", "actual", tag="급여", amount_out=1_000, id=1),
     Txn("2026-07", "actual", tag="상여", amount_out=2_000, id=2)]
b = [Txn("2026-07", "actual", tag="급여", amount_out=1_000, id=1),
     Txn("2026-07", "actual", tag="퇴직연금", amount_out=2_000, id=2)]
assert txn_fingerprint(a) == txn_fingerprint(list(reversed(a)))
assert txn_fingerprint(a) != txn_fingerprint(b)
ok("거래 지문 — 순서는 무시, 계정과목 한 건만 바뀌어도 달라짐")


# ── 4. 템플릿 이름 무관 탐색 ──────────────────────────────────────
from core.config import TEMPLATE_DIR, find_template  # noqa: E402

if TEMPLATE_DIR.exists():
    got_x = find_template(TEMPLATE_DIR / "이런이름없음.xlsm", (".xlsm",))
    got_p = find_template(TEMPLATE_DIR / "이런이름없음.pptx", (".pptx", ".potx"))
    if any(p.suffix == ".xlsm" for p in TEMPLATE_DIR.iterdir()):
        assert got_x.exists() and got_x.suffix == ".xlsm", got_x
        assert not got_x.name.startswith("_fast_") and got_x.name != "서식기준.xlsm", got_x
    if any(p.suffix == ".pptx" for p in TEMPLATE_DIR.iterdir()):
        assert got_p.exists() and got_p.suffix == ".pptx", got_p
    ok("템플릿 자동 탐색 — 이름이 달라도 찾고, 속도용 사본·서식기준은 제외")
else:
    print("- templates 폴더가 없어 템플릿 탐색 검사는 건너뜀")


# ── 5. 화면 파일 정적 검사 ────────────────────────────────────────
import importlib.util  # noqa: E402

for page in sorted((ROOT / "app" / "pages").glob("*.py")):
    tree = ast.parse(page.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and (node.module or "").startswith("core."):
            spec = importlib.util.find_spec(node.module)
            assert spec, f"{page.name}: {node.module} 모듈이 없습니다"
            mod = importlib.import_module(node.module)
            for al in node.names:
                assert hasattr(mod, al.name), \
                    f"{page.name}: {node.module}.{al.name} 가 없습니다"
ok("화면 파일이 없는 이름을 import 하고 있지 않음")


# ── 6. 화면 실제 실행 (Streamlit AppTest) ─────────────────────────
#    임시 DB에 가짜 거래를 넣고 화면 4개를 실제로 그려봅니다.
#    '검증 게이트 확인' → '거래를 고치면 확인이 풀린다' 까지 눌러서 확인합니다.
try:
    from streamlit.testing.v1 import AppTest
except Exception:                                    # 스트림릿이 없으면 건너뜀
    AppTest = None

if AppTest is None:
    print("- 스트림릿이 없어 화면 실행 검사는 건너뜀")
else:
    from core.data_store import (clear_period, init_db, load_transactions,  # noqa: E402
                                 save_month_balance, save_transactions,
                                 update_transaction)

    init_db()
    P = "2026-07"
    clear_period(P, "actual")
    clear_period(P, "plan")
    save_transactions([
        Txn(P, "actual", tag="급여", memo="7월 급여", amount_out=149_500_000, txn_date="0731"),
        Txn(P, "actual", tag="외상매출금", memo="매출대금", amount_in=500_000_000, txn_date="0731"),
        # 텐센트 건처럼 '출금 전용 계정과목인데 입금으로 들어온' 거래 → V4 가 빨간불
        Txn(P, "actual", tag="전산유지비", memo="클라우드 분담금", amount_in=56_328,
            txn_date="0731"),
        Txn(P, "plan", tag="부가가치세", memo="8월 부가세", amount_out=70_000_000),
    ])
    save_month_balance(P, open_krw=362.495882, open_usd=0.0, close_krw=321.357313,
                       close_usd=0.0, invest_krw=2100.0, invest_usd=0.0,
                       open_sheet="0630", close_sheet="0731")

    def _run(page: str):
        at = AppTest.from_file(str(ROOT / page), default_timeout=180)
        at.session_state["period"] = P
        at.run()
        assert not at.exception, f"{page}: {at.exception[0]}"
        return at

    for page in ("app/Home.py", "app/pages/3_Manage_Data.py",
                 "app/pages/4_Preview_Report.py"):
        _run(page)
    ok("화면 3개(Home·3·4)가 오류 없이 그려짐")

    # ── 올린 파일이 화면을 옮겨도 남는지 ──────────────────────────
    import openpyxl  # noqa: E402

    from core.uploads import KeptFile, keep  # noqa: E402

    _daily = [["", "", "1. 일일자금수지 실적"],
              ["통화구분", "거래처명", "내 역", "계정과목", "수입", "지출"],
              ["1)원화", "우리은행", "이자", "지급이자", 0, 1_000_000],
              ["", "카카오", "매출대금", "외상매출금", 5_000_000, 0],
              ["", "", "소계", "", 5_000_000, 1_000_000],
              ["", "", "3. 자금현황"],
              ["", "", "원화", "당일잔액", 362_495_882, ""]]
    _wb = openpyxl.Workbook()
    for _i, _nm in enumerate(("0731", "0630")):
        _ws = _wb.active if _i == 0 else _wb.create_sheet()
        _ws.title = _nm
        for _r in _daily:
            _ws.append(_r)
    _tmpx = Path(tempfile.mkdtemp()) / "daily.xlsx"
    _wb.save(_tmpx)

    kf = keep("actual", "일일자금계획.xlsx", _tmpx.read_bytes())
    at = AppTest.from_file(str(ROOT / "app/pages/1_Upload_Actual.py"), default_timeout=180)
    at.session_state["period"] = P
    at.session_state["actual_file"] = kf.to_dict()      # ← 2번 화면 갔다 돌아온 상태
    at.run()
    assert not at.exception, at.exception[0]
    heads = [h.value for h in at.subheader]
    assert "시트별 읽기 결과" in heads, heads
    assert any("그대로 씁니다" in s.value for s in at.success), [s.value for s in at.success]
    ok("업로드한 파일이 화면을 옮겼다 돌아와도 남아 있음 (다시 안 올려도 됨)")

    # 보관 파일이 지워졌으면 조용히 처음 상태로 (엉뚱한 파일을 읽지 않도록)
    Path(kf.path).unlink()
    assert KeptFile.from_dict(kf.to_dict()) is None
    ok("보관 파일이 사라지면 처음 상태로 되돌아감")

    # A → B → 다시 A 순서로 올려도 내용이 섞이지 않는지 (예전 고정 이름 방식의 함정)
    ka = keep("t", "a.xlsx", b"AAAA")
    kb = keep("t", "b.xlsx", b"BBBB")
    ka2 = keep("t", "a.xlsx", b"AAAA")
    assert ka.path == ka2.path != kb.path
    assert Path(ka2.path).read_bytes() == b"AAAA" and Path(kb.path).read_bytes() == b"BBBB"
    ok("A→B→A 순서로 올려도 파일 내용이 섞이지 않음 (이름이 곧 내용 지문)")

    at = _run("app/pages/4_Preview_Report.py")
    reds = [e.value for e in at.error]
    assert any("V4" in r for r in reds), reds
    # V2(외부 정합성) 금액은 소수점 둘째 자리까지만
    v2 = [r for r in reds if "V2" in r]
    assert v2 and "321.36" in v2[0] and "321.357" not in v2[0], v2
    ok("검증 게이트 — 문제 있으면 빨간불, 외부 정합성 금액은 소수점 2자리")

    acks = [b for b in at.button if "확인함" in b.label]
    assert acks, [b.label for b in at.button]
    acks[0].click().run()
    at2 = _run("app/pages/4_Preview_Report.py")
    assert not any("V4" in e.value for e in at2.error), [e.value for e in at2.error]
    ok("[확인함]을 누르면 초록색 '특이사항 없음' 으로 바뀜")

    # 확인 이후 거래를 고치면(계정과목은 그대로, 금액만) 확인이 풀려야 합니다
    t = [x for x in load_transactions(P, "actual") if x.tag == "전산유지비"][0]
    update_transaction(t.id, amount_in=99_999.0)
    at3 = _run("app/pages/4_Preview_Report.py")
    assert any("V4" in e.value for e in at3.error), [e.value for e in at3.error]
    ok("확인 뒤 거래가 바뀌면 확인이 자동으로 풀려 다시 빨간불")


# ── 7. 기초 운용자금을 전월 마지막 시트에서 가져오는지 ─────────────
from core.data_store import (init_db, load_report_remarks,  # noqa: E402
                             resolve_open_invest, save_fund_balance,
                             save_month_balance, save_report_remarks)

init_db()
Q = "2026-11"                                   # 다른 검사와 안 겹치는 달
save_month_balance(Q, open_krw=100.0, open_usd=0.0, close_krw=120.0, close_usd=0.0,
                   invest_krw=2050.0, invest_usd=0.0,
                   open_invest_krw=2100.0, open_invest_usd=0.0,
                   open_sheet="1031-1", close_sheet="1130-1")
krw, usd, src = resolve_open_invest(Q)
assert krw == 2100.0 and "1031-1" in src, (krw, src)
ok("기초 운용자금을 전월 마지막 시트(1031-1)에서 읽어옴")

# 시트값이 없으면 지난달 산출 이력으로 물러섬
Z = "2026-12"
save_fund_balance("2026-11", 1234.0, 0.0)
save_month_balance(Z, open_krw=0.0, open_usd=0.0, close_krw=0.0, close_usd=0.0,
                   invest_krw=0.0, invest_usd=0.0,
                   open_invest_krw=0.0, open_invest_usd=0.0,
                   open_sheet="", close_sheet="")
krw2, _u2, src2 = resolve_open_invest(Z)
assert krw2 == 1234.0 and "이력" in src2, (krw2, src2)
ok("시트에서 못 읽었으면 지난달 산출 이력으로 물러섬")


# ── 8. 4번에서 확정한 비고가 엑셀에 실제로 들어가는지 ──────────────
from core.config import TEMPLATE_PATH  # noqa: E402

if not TEMPLATE_PATH.exists():
    print("- 원본 워크북이 없어 비고 반영 검사는 건너뜀")
else:
    import warnings as _w  # noqa: E402
    _w.filterwarnings("ignore")
    import openpyxl  # noqa: E402

    from core.data_store import Txn as _T  # noqa: E402
    from core.engine import ROW_NO, aggregate  # noqa: E402
    from core.excel_writer import REMARK_COL, build_workbook  # noqa: E402
    from core.remarks_rule import (SUMMARY1_KEYS, cashflow_rows_from_summary1,  # noqa: E402
                                   generate_leaf_remarks, summary1_drafts)
    from core.report import DETAIL_ROWS  # noqa: E402

    M = 1_000_000
    _ax = [_T("2026-07", "actual", tag="외상매출금", memo="매출", amount_in=3 * M),
           _T("2026-07", "actual", tag="급여", memo="7월 급여", amount_out=14 * M)]
    _px = [_T("2026-07", "plan", tag="급여", memo="8월 급여", amount_out=15 * M)]
    _ra = aggregate(_ax, opening_krw=100.0)

    # 4번 화면에서 고친 셈 치고 저장
    W = "2026-09"
    _texts = [""] * len(DETAIL_ROWS)
    for _i, (_dep, _lab, _k) in enumerate(DETAIL_ROWS):
        if _k == "expense.salary":
            _texts[_i] = "사용자수정-급여 999백만"
        if _k == "expense.total":
            _texts[_i] = "사용자수정-지출소계(엑셀에는 안 들어감)"
    save_report_remarks(W, "detail", _texts)
    _four = {"actual_in": "● 7월 : 사용자수정-매출 3백만", "plan_in": "○ 8월 : ",
             "actual_out": "● 7월 : 사용자수정-급여 14백만", "plan_out": "○ 8월 : 급여 15백만",
             "open_bal": "○ 기초잔고: 유동자금 100백만, 운용자금 2,400백만",
             "close_bal": "○ 기말잔고: 유동자금 90백만, 운용자금 2,400백만"}
    save_report_remarks(W, "summary1", [_four[k] for k in SUMMARY1_KEYS])
    save_report_remarks(W, "cash_flow", cashflow_rows_from_summary1(_four))

    # 5번 화면과 똑같은 방식으로 비고를 모읍니다
    _leaf = generate_leaf_remarks(_ra, 7)
    _saved = load_report_remarks(W, "detail", len(DETAIL_ROWS))
    for (_dep, _lab, _k), _t in zip(DETAIL_ROWS, _saved):
        if (_t or "").strip() and _k in ROW_NO:
            _leaf[_k] = _t
    _s1 = load_report_remarks(W, "summary1", len(SUMMARY1_KEYS))
    _d1 = summary1_drafts(_ax, _px, 7, 8)
    _drafts = {k: (_s1[i].strip() or _d1[k]) for i, k in enumerate(SUMMARY1_KEYS)}

    _rep = build_workbook(
        actual_txns=_ax, plan_txns=_px, actual_result=_ra,
        actual_month=7, plan_month=8, opening_liq_krw=100.0, opening_liq_usd=0.0,
        opening_inv_krw=2400.0, opening_inv_usd=0.0, fx_rate=0.0,
        leaf_remarks=_leaf, summary_drafts=_drafts, out_name="_test_remarks.xlsm")
    _wb = openpyxl.load_workbook(_rep.path, keep_vba=True)

    _d3 = _wb["유동자금세부"]
    _got = _d3.cell(row=ROW_NO["expense.salary"], column=REMARK_COL).value
    assert _got == "사용자수정-급여 999백만", repr(_got)
    ok("4번에서 고친 세부내역 비고가 엑셀 유동자금세부 N열에 그대로 들어감")

    _txt = "\n".join(str(c.value) for r in _wb["요약1"].iter_rows()
                     for c in r if isinstance(c.value, str))
    assert "사용자수정-급여 14백만" in _txt, _txt[-1500:]
    assert "사용자수정-매출 3백만" in _txt, _txt[-1500:]
    ok("4번에서 고친 요약1 네 문장이 엑셀 요약1에 그대로 들어감")
    Path(_rep.path).unlink(missing_ok=True)


# ── 9. 비고 칸을 지우면 "nan" 이 아니라 공란 ───────────────────────
from core.data_store import clean_text  # noqa: E402

_nan = float("nan")
assert clean_text(_nan) == "" and clean_text(None) == "" and clean_text("nan") == ""
assert clean_text(" 급여 15백만 ") == " 급여 15백만 "
save_report_remarks("2026-10", "detail", [_nan, "정상 문구", None, "NaN"])
assert load_report_remarks("2026-10", "detail", 4) == ["", "정상 문구", "", ""]
ok("비고를 지우면 'nan' 이 아니라 공란으로 저장·출력됨")


# ── 10. PPT 산출물 회귀 (2026-08-25 사용자 지적 6건) ───────────────
from core.ppt_writer import PPT_TEMPLATE  # noqa: E402

if not PPT_TEMPLATE.exists():
    print("- PPT 템플릿이 없어 PPT 검사는 건너뜀")
else:
    from pptx import Presentation  # noqa: E402
    from pptx.util import Pt  # noqa: E402

    from core.data_store import Txn as _T2  # noqa: E402
    from core.ppt_writer import build_ppt, detect_template_months  # noqa: E402
    from core.report import build_report  # noqa: E402

    _M = 1_000_000
    _a2 = [_T2("2026-07", "actual", tag="외상매출금", memo="매출", amount_in=3 * _M),
           _T2("2026-07", "actual", tag="급여", memo="7월 급여", amount_out=20 * _M),
           _T2("2026-07", "actual", tag="원천세", memo="원천세", amount_out=4 * _M)]
    _p2 = [_T2("2026-07", "plan", tag="급여", memo="8월 급여", amount_out=21 * _M),
           _T2("2026-07", "plan", tag="기타입금", memo="분담금", amount_in=1 * _M)]
    _rep2 = build_report(_a2, _p2, 7, 8, open_krw=362.5, open_usd=0.0,
                         plan_open_krw=321.4, plan_open_usd=0.0,
                         invest_open_krw=2100.0, invest_open_usd=0.0, fx=0.0)

    _prs0 = Presentation(str(PPT_TEMPLATE))
    assert detect_template_months(_prs0) == (6, 7), detect_template_months(_prs0)
    ok("PPT 템플릿이 몇 월 파일인지 자동으로 알아냄 (6월 실적 / 7월 계획)")

    _notes = {k: [""] * len(_rep2.tables[k]) for k in _rep2.tables}
    _notes["detail"][5] = "∙7월 : 정기예금 만기 6.4백만\n∙8월 : 없음"   # 여러 줄
    _notes["cash_flow"][1] = "● 7월 : 매출 3백만"
    _out = build_ppt(_rep2, remarks=_notes, out_name="_test_ppt.pptx")
    _prs = Presentation(str(_out.path))
    _sl = list(_prs.slides)

    def _tb(s):
        return [sh.table for sh in s.shapes if getattr(sh, "has_table", False)]

    # (1) 2장 '2. 유동자금 계획' 표 머리글이 계획월로
    _h = _tb(_sl[1])[1].cell(0, 1).text
    assert "8월" in _h and "7월" not in _h, _h
    ok("PPT 2장 유동자금 계획 표 머리글 = 8월 계획")

    # (2) 3장 '3. 운용자금 계획' 표 머리글이 실적월·계획월로
    _t3 = _tb(_sl[2])[0]
    assert "7월" in _t3.cell(0, 2).text and "8월" in _t3.cell(0, 5).text, \
        (_t3.cell(0, 2).text, _t3.cell(0, 5).text)
    ok("PPT 3장 운용자금 계획 표 머리글 = 7월 실적 / 8월 계획")

    # (3) 5장 연간 자금실적 — 실적월·계획월 두 칸이 모두 갱신
    _t5 = _tb(_sl[4])[0]
    _mcol = {}
    for _c in range(len(_t5.columns)):
        _mm = re.fullmatch(r"(\d{1,2})\s*월", _t5.cell(2, _c).text.strip())
        if _mm:
            _mcol[int(_mm.group(1))] = _c
    assert 7 in _mcol and 8 in _mcol, _mcol

    def _val(row, month):
        return float(_t5.cell(row, _mcol[month]).text.strip().replace(",", "") or 0)

    _cf = _rep2.tables["cash_flow"]
    assert abs(_val(3, 7) - round(_cf[0][1])) <= 1, (_val(3, 7), _cf[0][1])
    assert abs(_val(7, 7) - round(_cf[2][1])) <= 1, (_val(7, 7), _cf[2][1])   # C.지출 실적
    assert abs(_val(7, 8) - round(_cf[2][2])) <= 1, (_val(7, 8), _cf[2][2])   # C.지출 계획
    assert abs(_val(25, 8) - round(_cf[5][2])) <= 1, (_val(25, 8), _cf[5][2])  # F.기말 계획
    assert abs(_val(9, 7) - 20) <= 1, _val(9, 7)                              # 인건비 실적
    assert abs(_val(9, 8) - 21) <= 1, _val(9, 8)                              # 인건비 계획
    ok("PPT 5장 연간 자금실적 — 실적월·계획월 숫자가 모두 갱신됨")

    # (4) 6장 비고 — 여러 줄이 문단으로 나뉘고 8pt 맑은 고딕
    _t6 = _tb(_sl[5])[0]
    _cell = _t6.cell(2 + 5, 11)
    assert _cell.text == _notes["detail"][5], repr(_cell.text)
    assert len(_cell.text_frame.paragraphs) == 2, len(_cell.text_frame.paragraphs)
    for _pa in _cell.text_frame.paragraphs:
        for _r in _pa.runs:
            assert _r.font.name == "맑은 고딕" and _r.font.size == Pt(8), \
                (_r.font.name, _r.font.size)
    ok("PPT 6장 비고 — 줄바꿈이 문단으로 나뉘고 전부 맑은 고딕 8pt")

    # (5) 비운 비고는 'nan' 이 아니라 공란
    _blank = [_t6.cell(2 + i, 11).text for i in range(len(_rep2.tables["detail"]))
              if i != 5]
    assert all(b.strip() == "" for b in _blank), [b for b in _blank if b.strip()]
    assert "nan" not in _t6.cell(2 + 0, 11).text.lower()
    ok("PPT 비고를 비우면 공란으로 출력 ('nan' 안 나옴)")
    Path(_out.path).unlink(missing_ok=True)


# ── 11. 수기 입력 안내 (주택자금 / 삼성카드 법인카드) ──────────────
from core.data_store import load_recheck_items  # noqa: E402
from core.recheck_items import check_missing, condition_text  # noqa: E402

_rules = load_recheck_items()
assert len(_rules) == 2, [(r["pattern"], r["counterpart"]) for r in _rules]
assert {(r["pattern"], r["counterpart"]) for r in _rules} == {("주택자금", ""),
                                                              ("법인카드", "삼성카드")}
ok("수기 입력 규칙 2개만 존재 (옛 중복 규칙 정리됨)")

Y = "2026-07"
_base = [Txn(Y, "actual", tag="급여", memo="7월 급여", amount_out=1)]
assert {m["label"] for m in check_missing(_base)} == {"주택자금대출 이자 지원",
                                                      "삼성카드 법인카드 대금"}
ok("해당 거래가 없으면 두 항목 모두 '수기 입력 필요'로 잡힘")

# 계정과목이 아직 안 붙어 있어도 낱말만으로 찾습니다
_have = _base + [Txn(Y, "actual", tag="", memo="주택자금대출 이자 지원", amount_out=1)]
assert {m["label"] for m in check_missing(_have)} == {"삼성카드 법인카드 대금"}
ok("적요에 '주택자금' 이 있으면 계정과목이 비어 있어도 찾음")

# 거래처가 삼성카드일 때만 잡힙니다
_wrong = _have + [Txn(Y, "actual", counterpart="현대카드", memo="법인카드 대금", amount_out=1)]
assert {m["label"] for m in check_missing(_wrong)} == {"삼성카드 법인카드 대금"}
_right = _have + [Txn(Y, "actual", counterpart="삼성카드", memo="7월분 법인카드 대금",
                      amount_out=1)]
assert check_missing(_right) == []
ok("법인카드는 거래처가 '삼성카드' 일 때만 찾음 (현대카드는 해당 없음)")

assert "삼성카드" in condition_text(_rules[1]) or "삼성카드" in condition_text(_rules[0])
ok("안내창에 찾은 조건이 사람 말로 표시됨")

if AppTest is not None:
    from core.data_store import clear_period as _cp  # noqa: E402

    _cp(Y, "actual")
    save_transactions([Txn(Y, "actual", tag="급여", memo="7월 급여", amount_out=1_000_000)])

    def _run3():
        at = AppTest.from_file(str(ROOT / "app/pages/3_Manage_Data.py"), default_timeout=180)
        at.session_state["period"] = Y
        at.run()
        assert not at.exception, at.exception[0]
        return at

    _at3 = _run3()
    assert any("직접 넣어야 하는 항목" in w.value for w in _at3.warning), \
        [w.value for w in _at3.warning]
    _labels = [b.label for b in _at3.button]
    assert _labels.count("➕ 이 건 추가") == 2, _labels
    assert "이번 달은 없음 — 닫기" in _labels, _labels
    ok("3번 화면에 수기 입력 안내창이 뜨고, 창 안에서 바로 추가할 수 있음")

    save_transactions([
        Txn(Y, "actual", tag="복리후생비", memo="주택자금대출 이자 지원", amount_out=500_000),
        Txn(Y, "actual", tag="직원비용", counterpart="삼성카드", memo="7월분 법인카드 대금",
            amount_out=3_000_000)])
    _at4 = _run3()
    assert not [w for w in _at4.warning if "직접 넣어야 하는 항목" in w.value]
    ok("해당 거래를 넣으면 안내창·경고가 사라짐")

    # 왜 안 떴는지 — 어떤 거래 덕분에 통과했는지 보여줍니다
    from core.recheck_items import check_all  # noqa: E402

    _all = check_all(load_transactions(Y, "actual", include_excluded=True))
    assert all(hits for _it, hits in _all), [(i["label"], len(h)) for i, h in _all]
    _hit = {i["label"]: h[0].counterpart for i, h in _all}
    assert _hit["삼성카드 법인카드 대금"] == "삼성카드", _hit
    ok("통과한 항목이 '어떤 거래 때문인지' 확인 가능 (check_all)")


# ── 12. 비고 서식 · 표 이름 (2026-08-25 사용자 지적) ───────────────
from core.remarks_rule import (SUMMARY1_ORDER, balance_notes,  # noqa: E402
                               cashflow_rows_from_summary1, detail_drafts,
                               mark_summary1, summary1_drafts)
from core.report import TABLE_TITLE  # noqa: E402

assert TABLE_TITLE["cash_flow"].format(a=7, p=8) == "1. 8월 자금수지"
ok("'1. 자금수지' 표 이름이 보고 대상 월(계획월) 기준")

# 월 표기는 어떤 수정을 배워도 사라지지 않아야 합니다
_no_month = infer_style("급여 15백만 / 상여 3백만", 7, style_for("detail"))
assert "{m}" in _no_month["prefix"], _no_month
ok("사용자가 월을 뺀 문장을 써도 다음 초안의 '∙N월 :' 은 유지됨")

assert mark_summary1("actual_in", "7월 : 매출 3백만") == "● 7월 : 매출 3백만"
assert mark_summary1("plan_out", "8월 : 급여 15백만") == "○ 8월 : 급여 15백만"
assert mark_summary1("actual_in", "● 이미 있음") == "● 이미 있음"
ok("요약1 비고에 ● (실적월) / ○ (계획월) 기호가 자동으로 붙음")

if PPT_TEMPLATE.exists():
    _rep3 = build_report(_a2, _p2, 7, 8, open_krw=362.5, open_usd=0.0,
                         plan_open_krw=321.4, plan_open_usd=0.0,
                         invest_open_krw=2100.0, invest_open_usd=0.0, fx=0.0)
    _bn = balance_notes(_rep3)
    assert _bn["open_bal"].startswith("○ 기초잔고: 유동자금 ") and "운용자금 2,100백만" in _bn["open_bal"], _bn
    assert _bn["close_bal"].startswith("○ 기말잔고: 유동자금 "), _bn
    ok("기초·기말잔고 비고를 원본 PPT 문장 그대로 자동 작성")

    _s1 = summary1_drafts(_a2, _p2, 7, 8, rep=_rep3)
    assert set(_s1) == set(SUMMARY1_ORDER), sorted(_s1)
    assert _s1["actual_out"].startswith("● 7월 : "), _s1["actual_out"]
    assert _s1["plan_out"].startswith("○ 8월 : "), _s1["plan_out"]

    _notes3 = {k: [""] * len(_rep3.tables[k]) for k in _rep3.tables}
    _notes3["cash_flow"] = cashflow_rows_from_summary1(_s1)
    _notes3["detail"] = detail_drafts(_rep3.actual, _rep3.plan, 7, 8)
    _out3 = build_ppt(_rep3, remarks=_notes3, out_name="_test_ppt2.pptx")
    _t2 = [sh.table for sh in list(Presentation(str(_out3.path)).slides)[1].shapes
           if getattr(sh, "has_table", False)][0]
    _n = [_t2.cell(r, 4).text for r in range(9)]
    assert _n[1].startswith("○ 기초잔고:"), _n[1]
    assert _n[2].startswith("● 7월"), _n[2]          # 실적월은 윗줄
    assert _n[3].startswith("○ 8월") or _n[3] == "", _n[3]   # 계획월은 아랫줄
    assert _n[4].startswith("● 7월"), _n[4]
    assert _n[5].startswith("○ 8월"), _n[5]
    assert _n[8].startswith("○ 기말잔고:"), _n[8]
    assert "\n" not in _n[4], _n[4]                  # 한 칸에 두 줄을 몰아넣지 않음
    ok("PPT 2장 비고 — ● 실적월/○ 계획월이 원본처럼 위아래 두 줄로 (지난달 문장 안 남음)")

    # 세부내역: 저장 안 한 줄도 화면에 보이던 초안 그대로 출력
    _t6b = [sh.table for sh in list(Presentation(str(_out3.path)).slides)[5].shapes
            if getattr(sh, "has_table", False)][0]
    _di = [i for i, (_, _, k) in enumerate(DETAIL_ROWS) if k == "expense.salary"][0]
    _cell6 = _t6b.cell(2 + _di, 11).text
    assert "7월" in _cell6 and "8월" in _cell6, repr(_cell6)
    ok("PPT 세부내역 비고에 '7월 / 8월' 이 언제 생긴 거래인지 함께 나옴")
    Path(_out3.path).unlink(missing_ok=True)


# ── 13. cmd 창을 껐다 켜도 남는지 (재시작) ─────────────────────────
if AppTest is not None:
    from core.uploads import forget, load_state, recall, remember, save_state  # noqa: E402

    X = "2026-07"
    _kf2 = keep("actual", "일일자금계획.xlsx", _tmpx.read_bytes())
    remember("actual", [_kf2.to_dict()])
    save_state("period", X)

    # 세션이 완전히 빈 상태(= 앱을 새로 켠 직후)
    _h = AppTest.from_file(str(ROOT / "app/Home.py"), default_timeout=180)
    _h.run()
    assert not _h.exception, _h.exception[0]
    assert _h.session_state["period"] == X, _h.session_state["period"]
    assert any("/ 4" in m.value for m in _h.metric), [m.value for m in _h.metric]
    ok("앱을 새로 켜도 실적 월이 유지되고, 진행 현황이 그려짐")

    _a4 = AppTest.from_file(str(ROOT / "app/pages/1_Upload_Actual.py"), default_timeout=180)
    _a4.session_state["period"] = X
    _a4.run()
    assert not _a4.exception, _a4.exception[0]
    assert "시트별 읽기 결과" in [x.value for x in _a4.subheader], \
        [x.value for x in _a4.subheader]
    ok("앱을 새로 켜도 지난번 올린 파일을 그대로 씀 (다시 안 올려도 됨)")

    forget("actual")
    assert recall("actual") == [] and load_state("actual") is None
    _a5 = AppTest.from_file(str(ROOT / "app/pages/1_Upload_Actual.py"), default_timeout=180)
    _a5.session_state["period"] = X
    _a5.run()
    assert any("파일을 올리면" in i.value for i in _a5.info), [i.value for i in _a5.info]
    ok("[파일 비우기] 를 누르면 다음에 켰을 때 다시 물어봄")


# ── 14. 계정과목 자동 교정 ─────────────────────────────────────────
from core.tag_classifier import apply_retag, retag  # noqa: E402

_t1 = Txn("2026-07", "actual", counterpart="삼성카드", memo="공용법인카드대금(7/20)",
          tag="기타지출", amount_out=4172)
assert retag(_t1) == ("직원비용", "카드사 법인카드 대금은 직원비용으로 분류"), retag(_t1)
_t2 = Txn("2026-07", "actual", counterpart="KT", memo="통신비", tag="기타지출", amount_out=1)
assert retag(_t2) is None
# 거래처가 카드사가 아니면 안 건드립니다 (소프트웨어 서비스명에 '법인카드'가 들어간 경우)
_t4 = Txn("2026-07", "actual", counterpart="주식회사 세포아소프트",
          memo="법인카드데이타 연동서비스", tag="기타지출", amount_out=110000)
assert retag(_t4) is None, retag(_t4)
_t3 = Txn("2026-07", "plan", counterpart="삼성카드", memo="삼성법인카드대금",
          tag="직원비용", amount_out=190000)
assert retag(_t3) is None                      # 이미 맞으면 안 건드림
_lst = [_t1, _t2, _t3, _t4]
_ch = apply_retag(_lst)
assert len(_ch) == 1 and _t1.tag == "직원비용" and _t2.tag == "기타지출", (_ch, _t1.tag)
ok("카드사 법인카드 대금만 기타지출 → 직원비용으로 교정 (소프트웨어 서비스는 제외)")


# ── 15. 비운 비고 vs 안 건드린 비고 (2026-08-25 사용자 지적) ────────
from core.data_store import (clear_report_remarks,  # noqa: E402
                             load_report_remarks_raw)

W2 = "2026-05"
clear_report_remarks(W2, "detail")
_raw = load_report_remarks_raw(W2, "detail", 4)
assert _raw == [None, None, None, None], _raw
ok("저장한 적 없는 비고는 None (→ 초안을 씁니다)")

save_report_remarks(W2, "detail", ["직접 쓴 글", "", "", "또 다른 글"])
_raw = load_report_remarks_raw(W2, "detail", 4)
assert _raw == ["직접 쓴 글", "", "", "또 다른 글"], _raw
ok("일부러 비운 줄은 '' 로 남음 (→ 초안이 되살아나지 않습니다)")

_draft = ["초안0", "초안1", "초안2", "초안3"]
_shown = [(_draft[i] if s is None else s) for i, s in enumerate(_raw)]
assert _shown == ["직접 쓴 글", "", "", "또 다른 글"], _shown
clear_report_remarks(W2, "detail")
_shown2 = [(_draft[i] if s is None else s)
           for i, s in enumerate(load_report_remarks_raw(W2, "detail", 4))]
assert _shown2 == _draft, _shown2
ok("[초안으로 되돌리기] 를 누르면 저장 기록이 지워져 초안이 살아남")


# ── 16. 위젯 만든 뒤 값 바꾸기 금지 (StreamlitAPIException) ─────────
if AppTest is not None:
    V = "2026-07"
    _at6 = AppTest.from_file(str(ROOT / "app/pages/4_Preview_Report.py"),
                             default_timeout=180)
    _at6.session_state["period"] = V
    _at6.session_state[f"pr_planopen_{V}"] = 999.0      # 기말잔고와 다르게
    _at6.run()
    assert not _at6.exception, _at6.exception[0]
    _btn = [b for b in _at6.button if "맞추기" in b.label]
    assert _btn, [b.label for b in _at6.button]
    _btn[0].click().run()
    # 예전에는 여기서 "cannot be modified after the widget is instantiated" 가 났습니다
    assert not _at6.exception, _at6.exception[0]
    assert abs(_at6.session_state[f"pr_planopen_{V}"] - 999.0) > 1e-6
    ok("[기말잔고로 맞추기] 버튼이 오류 없이 값을 바꿈 (on_click 콜백)")

    # 값이 맞으면 버튼을 아예 안 보여줍니다
    _at7 = AppTest.from_file(str(ROOT / "app/pages/4_Preview_Report.py"),
                             default_timeout=180)
    _at7.session_state["period"] = V
    _at7.run()
    assert not [b for b in _at7.button if "맞추기" in b.label], \
        [b.label for b in _at7.button]
    ok("기말잔고와 같으면 '맞추기' 버튼이 안 보임")


# ── 18. 비고만 고치고 바로 받기 (재사용) ───────────────────────────
if AppTest is not None and TEMPLATE_PATH.exists():
    import time as _time  # noqa: E402

    from core.data_store import load_month_balance as _lmb  # noqa: E402
    from core.outputs import build_outputs, cached_files  # noqa: E402

    T = "2026-07"
    _mb = _lmb(T)
    _v = {"open_krw": float(_mb["open_krw"]) if _mb else 100.0, "open_usd": 0.0,
          "plan_open_krw": float(_mb["close_krw"]) if _mb else 100.0,
          "plan_open_usd": 0.0, "invest_open_krw": 2100.0, "invest_open_usd": 0.0,
          "fx": 0.0}

    _o1 = build_outputs(T, _v, 7, 8, True, False, force=True)
    assert _o1.files and not _o1.reused, _o1
    _t0 = _time.time()
    _o2 = build_outputs(T, _v, 7, 8, True, False)
    _dt = _time.time() - _t0
    assert _o2.reused and _dt < 1.0, (_o2.reused, _dt)
    ok(f"바뀐 게 없으면 지난 파일 그대로 재사용 ({_dt:.2f}초)")

    save_report_remarks(T, "detail", ["바꿔봄"] + [""] * 35)
    _hit, _st2 = cached_files(T, _v, 7, 8)
    assert _hit is None, "비고를 바꿨는데도 재사용하려 함"
    ok("비고를 고치면 '내용이 바뀌었습니다' 로 다시 만들게 됨")

    _v2 = dict(_v, fx=1380.0)
    assert cached_files(T, _v2, 7, 8)[0] is None
    ok("환율·잔고를 바꿔도 다시 만들게 됨")

    # 4번 화면 안에서 만들고 받기 (화면 이동 0회)
    def _run4():
        at = AppTest.from_file(str(ROOT / "app/pages/4_Preview_Report.py"),
                               default_timeout=300)
        at.session_state["period"] = T
        at.run()
        assert not at.exception, at.exception[0]
        return at

    _a4 = _run4()
    _mk = [b for b in _a4.button if "파일 만들기" in b.label]
    assert _mk, [b.label for b in _a4.button]
    _mk[0].click().run()
    assert not _a4.exception, _a4.exception[0]
    _dls = [d.label for d in _a4.download_button]
    assert "⬇ 엑셀 받기" in _dls and "⬇ PPT 받기" in _dls, _dls
    ok("4번 화면에서 바로 만들고 받을 수 있음 (화면 이동 0회)")

    # 시작 화면에서도 만들어 둔 파일을 바로 받을 수 있어야 합니다
    _ah = AppTest.from_file(str(ROOT / "app/Home.py"), default_timeout=300)
    _ah.session_state["period"] = T
    _ah.run()
    assert not _ah.exception, _ah.exception[0]
    _hd = [d.label for d in _ah.download_button]
    assert "⬇ 엑셀 받기" in _hd, _hd
    ok("시작 화면에서 만들어 둔 파일을 바로 받을 수 있음")

    # 5번 페이지는 없어졌습니다 (4번으로 흡수)
    assert not (ROOT / "app/pages/5_Export_Excel.py").exists()
    _pages = sorted(p.name for p in (ROOT / "app/pages").glob("*.py"))
    assert _pages == ["1_Upload_Actual.py", "2_Upload_Plan.py", "3_Manage_Data.py",
                      "4_Preview_Report.py", "6_Remark_Labels.py",
                      "7_Tag_Dictionary.py"], _pages
    ok("화면이 4개(+가끔 쓰는 2개)로 정리됨")


# ── 19. 진행 현황 (사용자가 체크 안 해도 자동) ─────────────────────
if AppTest is not None:
    from core.data_store import clear_period as _cp2  # noqa: E402
    from core.progress import auto_vals, first_todo, steps  # noqa: E402

    G = "2026-03"
    _cp2(G, "actual")
    _cp2(G, "plan")
    _s = steps(G, auto_vals(G), output_ready=False)
    assert [x.done for x in _s] == [False, False, True, False], [x.detail for x in _s]
    assert first_todo(_s).no == 1
    ok("아무것도 안 올렸으면 1단계부터 (자료에서 자동 판정)")

    save_transactions([Txn(G, "actual", tag="급여", memo="급여", amount_out=1, txn_date="0331")])
    _s = steps(G, auto_vals(G), output_ready=False)
    assert _s[0].done and not _s[1].done, [x.detail for x in _s]
    assert "1건" in _s[0].detail and "시트 1개" in _s[0].detail, _s[0].detail
    assert first_todo(_s).no == 2
    ok("실적만 올리면 2단계(계획 업로드)가 다음 할 일")

    save_transactions([Txn(G, "plan", dept="회계팀", tag="급여", memo="급여", amount_out=1)])
    _s = steps(G, auto_vals(G), output_ready=False)
    assert _s[1].done and "회계팀" in _s[1].detail, _s[1].detail
    ok("계획을 올리면 어느 부서가 들어왔는지 보여줌 (대기 상황 판단용)")

    # 계정과목이 비면 3단계가 빨간불
    save_transactions([Txn(G, "actual", tag="", memo="확인 필요", amount_out=1, txn_date="0331")])
    _s = steps(G, auto_vals(G), output_ready=False)
    assert not _s[2].done and "1건" in _s[2].detail, _s[2].detail
    assert first_todo(_s).no == 3
    ok("계정과목이 비면 3단계가 남은 일로 잡힘")

    # 시작 화면에 상태바·단계 카드가 그려지는지
    _ah2 = AppTest.from_file(str(ROOT / "app/Home.py"), default_timeout=300)
    _ah2.session_state["period"] = G
    _ah2.run()
    assert not _ah2.exception, _ah2.exception[0]
    assert any("/ 4" in m.value for m in _ah2.metric), [m.value for m in _ah2.metric]
    _html = " ".join(str(m.value) for m in _ah2.markdown)
    assert "fp-todo" in _html and "fp-ok" in _html
    ok("시작 화면에 상태바 + 단계별 카드(남은 칸은 깜빡임)가 나옴")

print(f"\n{n_ok}개 항목 전부 통과했습니다.")
