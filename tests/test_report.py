"""보고서 표(미리보기·PPT 공용) 회귀 테스트 — 2026-08-24에 실제로 터졌던 문제만 담았습니다.

외부 파일 없이 돕니다.  실행:  python tests/test_report.py
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.data_store import Txn                                  # noqa: E402
from core.excel_writer import _MONTH_RE, OLD_ACTUAL_MONTH, OLD_PLAN_MONTH  # noqa: E402
from core.report import build_report                             # noqa: E402

M = 1_000_000
FX = 1400.0


def txn(tag, direction, amount, currency="KRW"):
    return Txn(period="2026-07", kind="actual", tag=tag, currency=currency,
               amount_in=amount if direction == "in" else 0,
               amount_out=amount if direction == "out" else 0, memo="테스트")


def sample():
    actual = [txn("외상매출금", "in", 100 * M), txn("이자수익", "in", 50, "USD"),
              txn("급여", "out", 30 * M), txn("지급이자", "out", 7 * M),
              txn("외주용역비", "out", 3 * M), txn("금융상품", "in", 300 * M)]
    plan = [txn("외상매출금", "in", 20 * M), txn("급여", "out", 5 * M),
            txn("지급이자", "out", 1 * M)]
    return build_report(actual, plan, 7, 8,
                        open_krw=500.0, open_usd=10.0,
                        plan_open_krw=700.0, plan_open_usd=60.0,
                        invest_open_krw=2400.0, invest_open_usd=0.0, fx=FX)


def near(a, b, tol=1e-9):
    return abs(a - b) < tol


rep = sample()
rows = {r[0].strip(): r for r in rep.tables["detail"]}


# ── 1. 월 치환이 연쇄되지 않는다 (6월→7월→8월 버그) ──────────────
table = {str(OLD_ACTUAL_MONTH): 7, str(OLD_PLAN_MONTH): 8}


def swap(text):
    return _MONTH_RE.sub(lambda m: f"{table[m.group(1)]}월", text)


assert swap("1. 6월 자금수지") == "1. 7월 자금수지"
assert swap("2. 7월 유동자금 계획") == "2. 8월 유동자금 계획"
assert swap("6월 실적 및 7월 계획") == "7월 실적 및 8월 계획"
assert swap("16월") == "16월", "앞에 숫자가 붙은 건 건드리면 안 됩니다"
print("✓ 월 표기 치환 — 실적 7월 / 계획 8월이 서로 안 섞임")


# ── 2. 지출 소계 = 눈에 보이는 하위 항목의 합 (지급이자 행 없음) ──
kids = ["인건비", "마케팅비", "제세공과금", "자산매입", "투자금", "기타"]
for col, who in ((1, "실적 원화"), (4, "계획 원화")):
    total = rows["지출"][col]
    child = sum(rows[k][col] for k in kids)
    assert near(total, child), f"{who}: 지출 {total} vs 하위합 {child}"
etc_rest = [r for r in rep.tables["detail"] if r[0].strip() == "기타"][-1]
assert near(etc_rest[1], 10.0), f"기타>기타에 지급이자가 안 들어감: {etc_rest[1]}"
print("✓ 세부내역 지출 소계 = 하위 항목 합 (지급이자는 '기타>기타'에 포함)")


# ── 3. 기말 유동잔고가 표마다 같은 값 (환산계를 원화 칸에 넣던 버그) ──
det_g = [r for r in rep.tables["detail"] if r[0].startswith("G.")][0]
liq_i = [r for r in rep.tables["liquid_plan"] if r[0].startswith("I.")][0]
stock_krw, stock_usd = rep.tables["stock"][0], rep.tables["stock"][1]
assert near(det_g[4], liq_i[1]) and near(det_g[5], liq_i[2]), "세부내역 G ≠ 유동자금계획 I"
assert near(det_g[4], stock_krw[3]) and near(det_g[5], stock_usd[3]), "세부내역 G ≠ stock A"
assert near(liq_i[3], det_g[6]), "환산계가 서로 다름"
# 외화가 두 번 더해지지 않았는지: 원화 칸에 환산계가 들어가면 아래가 깨집니다
assert near(liq_i[3], liq_i[1] + liq_i[2] * FX / 1000), "환산계 = 원화 + 외화×환율 이어야 함"
assert near(rep.tables["liquid_plan"][6][1] - rep.tables["liquid_plan"][7][1], liq_i[1]), "G-H=I 불일치"
print("✓ 기말 유동잔고가 세부내역·유동자금계획·stock 세 표에서 동일 (외화 이중계산 없음)")


# ── 4. FX 거래가 있으면 구조 경고가 뜬다 ─────────────────────────
plain = sample()
assert plain.meta["structure_warnings"] == [], "FX가 없는데 경고가 뜨면 안 됩니다"
with_fx = build_report([txn("외환매도", "in", 50 * M)], [], 7, 8,
                       open_krw=0.0, open_usd=0.0, plan_open_krw=0.0, plan_open_usd=0.0,
                       invest_open_krw=0.0, invest_open_usd=0.0, fx=FX)
assert with_fx.meta["structure_warnings"], "FX 거래가 있으면 경고가 떠야 합니다"
print("✓ FX 거래가 있으면 '표에 FX 행이 없음' 경고 발생")


# ── 5. PPT 글상자의 월 표기가 run으로 쪼개져 있어도 바뀐다 ────────
class _Run:
    def __init__(self, t):
        self.text = t


class _Para:
    def __init__(self, parts):
        self.runs = [_Run(p) for p in parts]

    @property
    def text(self):
        return "".join(r.text for r in self.runs)


from core.ppt_writer import _sub_digits_across_runs                # noqa: E402

pat = re.compile(r"(?<!\d)(6|7)월")
para = _Para(["Ⅰ. ", "7", "월 자금운용계획(안)"])
_sub_digits_across_runs(para, pat, lambda m: str({"6": 7, "7": 8}[m.group(1)]))
assert para.text == "Ⅰ. 8월 자금운용계획(안)", para.text
assert len(para.runs) == 3, "run 개수가 유지돼야 글꼴·색이 안 깨집니다"
print("✓ 슬라이드 글상자 — 월 표기가 여러 run에 쪼개져 있어도 치환됨 (서식 유지)")


# ── 6. PPT 맨 위 요약 문장의 숫자가 실제로 채워진다 ──────────────
from pathlib import Path as _P                                     # noqa: E402
_TPL = _P(__file__).resolve().parents[1] / "templates" / "자금운용계획_template.pptx"
if _TPL.exists():
    from pptx import Presentation                                  # noqa: E402
    from core.ppt_writer import build_ppt                          # noqa: E402

    r2 = build_report([txn("외상매출금", "in", 100 * M), txn("급여", "out", 30 * M)],
                      [txn("외상매출금", "in", 90 * M), txn("급여", "out", 5 * M)],
                      7, 8, open_krw=500.0, open_usd=0.0,
                      plan_open_krw=570.0, plan_open_usd=0.0,
                      invest_open_krw=2400.0, invest_open_usd=0.0, fx=FX)
    out = build_ppt(r2, out_name="_test_headline.pptx")
    assert not out.warnings, out.warnings
    prs = Presentation(str(out.path))

    def headline(idx, needle):
        for sh in prs.slides[idx].shapes:
            if not sh.has_text_frame:
                continue
            for pa in sh.text_frame.paragraphs:
                t = "".join(x.text for x in pa.runs)
                if needle in t and "백만" in t:
                    return t, len(pa.runs)
        return "", 0

    s2, n2 = headline(1, "flow")
    assert "3,055.0" in s2 and "2,970.0" in s2 and "+85.0" in s2, s2
    assert "8월 자금수지" in s2, s2
    assert n2 > 8, "run이 합쳐져 서식이 깨졌습니다"
    s3, n3 = headline(2, "stock")
    assert "3,055.0" in s3 and "655.0" in s3 and "2,400.0" in s3, s3
    assert "8월말 자금" in s3, s3
    out.path.unlink(missing_ok=True)
    print("✓ PPT 2·3장 맨 위 요약 문장 숫자 자동 채움 (서식 유지)")
else:
    print("- PPT 요약 문장 테스트는 건너뜀 (templates/자금운용계획_template.pptx 없음)")


# ── 7. 엑셀 산출물 — 2026-08-25 요청분 회귀 ──────────────────────
_XL = _P(__file__).resolve().parents[1] / "templates" / "자금운용계획_template.xlsm"
if _XL.exists():
    import warnings as _w                                          # noqa: E402
    _w.filterwarnings("ignore")
    import openpyxl                                                # noqa: E402
    from core.engine import aggregate                              # noqa: E402
    from core.excel_writer import build_workbook                   # noqa: E402
    from core.remarks_rule import generate_leaf_remarks            # noqa: E402

    ax = [txn("외상매출금", "in", 3 * M), txn("급여", "out", 14 * M),
          txn("원천세", "out", 2 * M), txn("전산유지비", "out", 1 * M),
          txn("광고판촉비", "out", 5 * M)]
    px = [txn("외상매출금", "in", 2 * M), txn("급여", "out", 15 * M),
          txn("지급수수료", "out", 1 * M)]
    ra_ = aggregate(ax, opening_krw=100.0)
    rep_xl = build_workbook(
        actual_txns=ax, plan_txns=px, actual_result=ra_,
        actual_month=7, plan_month=8,
        opening_liq_krw=100.0, opening_liq_usd=0.0,
        opening_inv_krw=2400.0, opening_inv_usd=0.0, fx_rate=0.0,
        leaf_remarks=generate_leaf_remarks(ra_, 7), summary_drafts={},
        out_name="_test_excel.xlsm")
    wbx = openpyxl.load_workbook(rep_xl.path, keep_vba=True)

    # 7-0 3p 집계 수식이 태그 사전 기준과 일치 (원본은 라벨셀 기준이라 0이 나왔음)
    d3 = wbx["유동자금세부"]
    from core.engine import LEAF as _LEAF, ACCOUNT_MAP as _AM             # noqa: E402
    for _k, _no, _lab, _dir, _ in _LEAF:
        want = sorted({a for (a, dd), kk in _AM.items() if kk == _k and dd == _dir})
        if not want:
            continue
        for _col in (7, 10):
            f = str(d3.cell(_no, _col).value or "")
            got = sorted(set(re.findall(r'\$A:\$A,"([^"]+)"', f)))
            assert got == want, (_no, _lab, _col, want, got)
    print("✓ 유동자금세부 집계 수식이 태그 사전과 정확히 일치 (실적·계획 양쪽)")

    # 7-1 원시 데이터 시트 이름
    assert "실적월" in wbx.sheetnames and "계획월" in wbx.sheetnames, wbx.sheetnames
    assert "6월데이터" not in wbx.sheetnames and "7월데이터" not in wbx.sheetnames
    print("✓ 원시 데이터 시트 이름이 '실적월 / 계획월'")

    # 7-2 요약1 두 표의 시작 열이 같다
    #     ※ templates/서식기준.xlsm 이 있으면 **사용자가 손본 배치**를 그대로 쓰므로
    #        아래 배치 검사(7-2·7-3)는 건너뜁니다.
    from core.excel_writer import STYLE_REF_PATH                     # noqa: E402
    _use_ref = STYLE_REF_PATH.exists()
    s1 = wbx["요약1"]

    def first_col(ws, row):
        for c in range(1, 20):
            if ws.cell(row=row, column=c).value not in (None, ""):
                return c
        return None

    r_head1 = next(r for r in range(1, 40) if s1.cell(r, 3).value == "구분")
    r_head2 = next(r for r in range(r_head1 + 1, 60)
                   if str(s1.cell(r, 3).value or "").strip() == "구분")
    if _use_ref:
        print("- 요약1 배치 검사는 건너뜀 (서식기준 파일의 배치를 따름)")
    else:
        assert first_col(s1, r_head1) == first_col(s1, r_head2) == 3, \
            (first_col(s1, r_head1), first_col(s1, r_head2))
        print("✓ 요약1 — 자금수지 표와 유동자금 계획 표가 같은 열(C)에서 시작")

    # 7-3 요약1·요약2의 표 사이 빈 줄 개수가 같다
    def gap(ws):
        rows = [r for r in range(1, ws.max_row + 1)
                if any(c.value not in (None, "") for c in ws[r])]
        for i in range(1, len(rows)):
            if rows[i] - rows[i - 1] > 3:
                return rows[i] - rows[i - 1] - 1
        return 0

    if _use_ref:
        # 기준 파일을 쓸 때는 그 파일의 열 너비·병합이 그대로 왔는지를 본다
        ref_wb = openpyxl.load_workbook(STYLE_REF_PATH)
        for nm in ("요약1", "요약2"):
            rw = {k: v.width for k, v in ref_wb[nm].column_dimensions.items() if v.width}
            ow = {k: v.width for k, v in wbx[nm].column_dimensions.items() if v.width}
            assert rw == ow, (nm, {k: (rw.get(k), ow.get(k)) for k in set(rw) | set(ow)
                                   if rw.get(k) != ow.get(k)})
            assert {str(m) for m in ref_wb[nm].merged_cells.ranges} == \
                   {str(m) for m in wbx[nm].merged_cells.ranges}, nm
        print("✓ 서식기준 파일의 열 너비·병합이 산출물에 그대로 유지됨")
    else:
        assert gap(s1) == gap(wbx["요약2"]), (gap(s1), gap(wbx["요약2"]))
        print(f"✓ 요약1·요약2 표 사이 간격이 동일 ({gap(s1)}줄)")

    # 7-4 자금실적 — 라벨 자리에 맞게 들어갔는지 + 계획월도 채워졌는지
    ys = wbx["2026년 자금실적"]
    labels = {}
    for r in range(11, 46):
        for c in (2, 3, 4, 5):
            v = ys.cell(r, c).value
            if isinstance(v, str) and v.strip():
                labels.setdefault(v.strip(), r)
                break
    assert abs(ys.cell(labels["제세공과금"], 14).value - 2.0) < 1e-6, "7월 원천세가 제세공과금 자리에 없음"
    assert abs(ys.cell(labels["지급수수료"], 15).value - 1.0) < 1e-6, "8월 계획이 안 채워짐"
    assert labels.get("마케팅,지급이자,외주용역비"), "마케팅 행이 자동 삽입되지 않음"
    assert abs(ys.cell(labels["마케팅,지급이자,외주용역비"], 14).value - 5.0) < 1e-6
    # '비경상'은 수입(위)·지출(아래) 두 군데 → 지출 쪽 소계 수식이 살아있는지 본다
    nonop = [r for r in range(11, 46)
             if any(str(ys.cell(r, c).value or "").strip() == "비경상" for c in (2, 3, 4, 5))]
    assert str(ys.cell(nonop[-1], 14).value).startswith("=SUM"), "지출 비경상 소계 수식이 깨짐"
    assert abs(ys.cell(nonop[0], 14).value) < 1e-9, "수입 비경상이 0이어야 함"
    assert abs(ys.cell(nonop[0] - 1, 14).value - 3.0) < 1e-6, "수입 경상 미반영"
    print("✓ 2026년 자금실적 — 라벨 기준으로 실적월·계획월 모두 반영, 마케팅 행 자동 삽입")

    # 7-4b 실적월 머리글에 (A), 계획월·이후 달은 그대로
    hdr = next(r for r in range(1, 20) if str(ys.cell(r, 8).value or "").startswith("1월"))
    assert ys.cell(hdr, 14).value == "7월(A)", ys.cell(hdr, 14).value
    assert ys.cell(hdr, 15).value == "8월", ys.cell(hdr, 15).value
    assert ys.cell(hdr, 16).value == "9월", ys.cell(hdr, 16).value
    print("✓ 자금실적 머리글 — 실적월만 '7월(A)', 계획월·이후 달은 그대로")

    # 7-5 연간 — 계획월(8월) 열도 채워짐
    an = wbx["연간"]
    assert abs(an["AB15"].value - 15.0) < 1e-6, an["AB15"].value   # 8월 급여
    assert abs(an["Y15"].value - 14.0) < 1e-6, an["Y15"].value     # 7월 급여
    print("✓ 연간 — 실적월·계획월 둘 다 채워짐")

    # 7-6 유동자금세부 N열에 지난달 문구가 남지 않음
    d3 = wbx["유동자금세부"]
    stale = [d3.cell(r, 14).value for r in range(6, 50)
             if d3.cell(r, 14).value and "6월" in str(d3.cell(r, 14).value)]
    assert not stale, stale
    print("✓ 유동자금세부 N열 — 지난달 문구 잔재 없음")

    # 7-7 없는 시트를 가리키는 수식이 하나도 없어야 한다
    #     (시트 이름을 바꿀 때 '6월데이터'!$E:$E 처럼 행 번호 없는 참조를 놓쳐
    #      파일 전체가 #REF!/#VALUE!로 깨졌던 사고의 재발 방지)
    names = set(wbx.sheetnames)
    prefix = re.compile(r"'((?:[^']|'')+)'!|(?<![A-Za-z0-9_.!$])([A-Za-z_가-힣][A-Za-z0-9_.가-힣]*)!")
    dangling = []
    for sh in wbx.worksheets:
        for row_ in sh.iter_rows():
            for c in row_:
                if not isinstance(c.value, str) or not c.value.startswith("="):
                    continue
                for m in prefix.finditer(c.value):
                    nm = (m.group(1) or m.group(2) or "").replace("''", "'")
                    if nm and nm not in names and not nm.isupper():
                        dangling.append((sh.title, c.coordinate, nm))
    assert not dangling, dangling[:5]
    print(f"✓ 없는 시트를 가리키는 수식 0건 (시트 {len(names)}개 전수 검사)")

    # 7-8 시트 순서
    from core.excel_writer import SHEET_ORDER                        # noqa: E402
    assert wbx.sheetnames == [n for n in SHEET_ORDER if n in names], wbx.sheetnames
    print("✓ 시트 순서 — 요약1 · 요약2 · 유동자금세부 · 연간 · 실적월 · 계획월 · 2025/2026년 자금실적")

    rep_xl.path.unlink(missing_ok=True)
    _n_items = 17
else:
    print("- 엑셀 산출 테스트는 건너뜀 (templates/자금운용계획_template.xlsm 없음)")
    _n_items = 6


# ── 8. 속도 관련 회귀 (2026-08-25) ───────────────────────────────
# 8-1 캐시 함수 인자가 밑줄로 시작하면 스트림릿이 그 인자를 **무시**합니다.
#     'stamp'(파일 지문)를 '_stamp'로 써서 파일을 고쳐 올려도 옛 내용이 나오던 사고 방지.
import ast as _ast                                                  # noqa: E402

_pages = (_P(__file__).resolve().parents[1] / "app" / "pages")
_bad = []
for _f in sorted(_pages.glob("*.py")):
    tree = _ast.parse(_f.read_text(encoding="utf-8"))
    for node in _ast.walk(tree):
        if not isinstance(node, _ast.FunctionDef):
            continue
        cached = any(
            "cache_data" in _ast.dump(d) or "cache_resource" in _ast.dump(d)
            for d in node.decorator_list)
        if not cached:
            continue
        for arg in node.args.args:
            if arg.arg.startswith("_"):
                _bad.append(f"{_f.name}:{node.name}({arg.arg})")
assert not _bad, f"캐시 함수의 밑줄 인자는 무시됩니다: {_bad}"
print("✓ 캐시 함수 인자에 밑줄 없음 (지문이 무시되던 버그 방지)")

# 8-2 daily_bundle 은 엑셀을 **한 번만** 연다
_DAILY = _P(__file__).resolve().parents[1] / "data" / "_upload_actual.xlsx"
if _DAILY.exists():
    import openpyxl as _ox                                          # noqa: E402
    from core import parsers as _pr                                 # noqa: E402
    _orig_open, _cnt = _ox.load_workbook, [0]

    def _spy(*a, **k):
        _cnt[0] += 1
        return _orig_open(*a, **k)

    _pr.openpyxl.load_workbook = _spy
    try:
        b = _pr.daily_bundle(str(_DAILY), 7)
    finally:
        _pr.openpyxl.load_workbook = _orig_open
    assert _cnt[0] == 1, f"파일을 {_cnt[0]}번 열었습니다 (1번이어야 함)"
    assert b["sheets"] and b["rows"], b
    assert b["prev_sheet"] is None or b["prev_sheet"] in b["rows"]
    print(f"✓ daily_bundle — 엑셀을 1번만 열어 시트 {len(b['rows'])}개를 한꺼번에 읽음")
    _n_items += 1
else:
    print("- daily_bundle 테스트는 건너뜀 (data/_upload_actual.xlsx 없음)")

print(f"\n{_n_items + 1}개 항목 전부 통과했습니다.")
