import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import pandas as pd
import streamlit as st

from core.nav import apply_nav
from core.data_store import (clear_period, init_db, save_month_balance,
                             save_transactions)
from core.parsers import (SECTION_ACTUAL, SECTION_PLAN, daily_bundle,
                          daily_sheet_subtotals, parse_daily_sheet,
                          read_cash_status)
from core.period import fmt
from core.tag_classifier import apply_retag, suggest_tag
from core.uploads import KeptFile, forget, keep, recall, remember

st.set_page_config(page_title="1. 실적 업로드", layout="wide")
apply_nav(st)
init_db()
st.title("1️⃣ 실적 업로드 — 일일자금계획")


# ── 빠르게 만드는 장치 ────────────────────────────────────────────
# 스트림릿은 화면에서 뭘 누를 때마다 이 파일을 처음부터 다시 실행합니다.
# 그대로 두면 그때마다 엑셀을 9번씩 다시 열었습니다. → 한 번만 열고 결과를 기억합니다.
#   (`stamp` = 파일 내용 지문. 내용이 바뀌면 지문이 달라져 자동으로 다시 읽습니다.
#    ※ 인자 이름을 밑줄로 시작하면 스트림릿이 그 인자를 무시하니 주의)
@st.cache_data(show_spinner="엑셀을 읽는 중입니다…", max_entries=6)
def c_bundle(path: str, stamp: str, month: int) -> dict:
    """파일을 **한 번만** 열어 시트 목록·시트별 내용·전월 시트를 한꺼번에 가져옵니다."""
    return daily_bundle(path, month)


@st.cache_data(show_spinner=False, max_entries=12)
def c_parse(path: str, stamp: str, month: int, picked: tuple, period_: str,
            include_plan_: bool) -> tuple:
    """고른 시트를 거래 목록으로 바꿉니다 (같은 조건이면 다시 계산하지 않습니다)."""
    rows_by = c_bundle(path, stamp, month)["rows"]
    out, per_sheet_, compare_, retagged_ = [], [], [], []
    for name in picked:
        rows = rows_by.get(name, [])
        acts = parse_daily_sheet(rows, period_, name, SECTION_ACTUAL)
        plans = parse_daily_sheet(rows, period_, name, SECTION_PLAN) if include_plan_ else []
        # 원본이 적어 둔 '소계'와 우리가 읽은 합계를 나란히 놓습니다 (대조표용)
        for st_ in daily_sheet_subtotals(rows):
            if st_["구간"] == SECTION_PLAN and not include_plan_:
                continue
            src = acts if st_["구간"] == SECTION_ACTUAL else plans
            mine = [t for t in src if (t.currency or "KRW") == st_["통화"]]
            got_in = sum(t.amount_in for t in mine)
            got_out = sum(t.amount_out for t in mine)
            d_in = got_in - (st_["수입"] or 0)
            d_out = got_out - (st_["지출"] or 0)
            compare_.append({
                "시트명": name, "표": st_["구간"], "통화": st_["통화"],
                "원본 소계 수입": st_["수입"] or 0, "앱이 읽은 수입": got_in,
                "원본 소계 지출": st_["지출"] or 0, "앱이 읽은 지출": got_out,
                "결과": "일치" if abs(d_in) < 1 and abs(d_out) < 1
                        else f"차이 수입 {d_in:+,.0f} / 지출 {d_out:+,.0f}",
            })
        for t in plans:
            t.kind = "actual"              # 같은 달 실적 집계에 합침
        txns = acts + plans
        retagged_ += apply_retag(txns)     # 원본 계정과목 자동 교정
        out.extend(txns)
        per_sheet_.append({"시트명": name,
                           "실적": f"{len(acts):,}건",
                           "계획": f"{len(plans):,}건",
                           "합계": f"{len(txns):,}건",
                           "계정과목 있는 항목(계)": f"{sum(1 for t in txns if t.tag):,}건"})
    return out, per_sheet_, compare_, retagged_

period = st.session_state.get("period", "")
if not period:
    st.error("먼저 시작 화면(Home)에서 실적 월을 정해주세요.")
    st.stop()
st.info(f"📅 **{fmt(period)} 실적** 을 읽습니다. "
        "해당 월의 일자 시트(MMDD, MMDD-N)를 **전부 찾아서 합칩니다.**")
target_month = int(period.split("-")[1])

# ── 올린 파일은 화면을 옮겨도 그대로 남습니다 ─────────────────────
# 스트림릿은 다른 화면에 갔다 오면 업로드 칸을 비워버립니다(원래 동작).
# 그래서 올린 파일을 따로 보관해 두고, 돌아왔을 때 그걸 계속 씁니다.
# cmd 창을 껐다 켜도 남도록, 세션뿐 아니라 디스크에도 기록해 둡니다.
KEEP_KEY = "actual_file"

up = st.file_uploader("일일자금계획.xlsx 파일을 올려주세요", type=["xlsx", "xlsm"])
if up is not None:
    rec = keep("actual", up.name, bytes(up.getbuffer())).to_dict()
    st.session_state[KEEP_KEY] = rec
    remember("actual", [rec])

kept = KeptFile.from_dict(st.session_state.get(KEEP_KEY))
if kept is None:                      # 앱을 새로 켠 경우 — 지난번 파일을 되살립니다
    prev = recall("actual")
    kept = KeptFile.from_dict(prev[0]) if prev else None
    if kept:
        st.session_state[KEEP_KEY] = kept.to_dict()
if kept is None:
    st.session_state.pop(KEEP_KEY, None)
    st.info("파일을 올리면 아래에 읽기 결과가 나옵니다.")
    st.stop()

if up is None:
    k1, k2 = st.columns([4, 1])
    k1.success(f"직전에 올리신 **{kept.name}** 을 그대로 씁니다. "
               "다른 파일로 바꾸시려면 위에 새로 올려주세요.")
    with k2:
        st.write("")
        if st.button("↺ 파일 비우기", use_container_width=True):
            st.session_state.pop(KEEP_KEY, None)
            forget("actual")
            st.rerun()

tmp, stamp = Path(kept.path), kept.stamp

bundle = c_bundle(str(tmp), stamp, target_month)
months = bundle["months"]
if not months:
    st.error("MMDD 형식(예: 0731)의 시트를 못 찾았습니다.")
    st.stop()
if target_month not in months:
    st.warning(f"{target_month}월 시트가 없습니다. 파일에 있는 달: "
               + ", ".join(f"{m}월" for m in months))
    target_month = st.selectbox("합칠 달", months, format_func=lambda m: f"{m}월")
    bundle = c_bundle(str(tmp), stamp, target_month)   # 고른 달로 다시

sheets = bundle["sheets"]
st.success(f"**{target_month}월 시트 {len(sheets)}개**를 찾았습니다.")
st.caption(" · ".join(sheets))

pick = st.multiselect("합칠 시트 (기본값: 전부)", sheets, default=sheets)
if not pick:
    st.stop()

st.divider()
which = st.radio(
    "시트 안의 어느 표를 읽을까요?",
    ["실적 + 계획 표 둘 다 (권장)", "실적 표만"],
    horizontal=True,
    help="일일자금계획 시트에는 '1. 일일자금수지 실적'과 '2. 일일자금수지 계획' 두 표가 있습니다. "
         "둘 다 그 달에 실제로 나간/들어온 자금이므로 기본값은 둘 다 합치기입니다.",
)
include_plan = which.startswith("실적 + 계획")

# ── 읽기 (같은 파일·같은 선택이면 다시 계산하지 않습니다) ──
preview, per_sheet, compare, retagged = c_parse(
    str(tmp), stamp, target_month, tuple(pick), period, include_plan)

st.subheader("시트별 읽기 결과")
st.dataframe(pd.DataFrame(per_sheet), use_container_width=True, hide_index=True)

n_gap = sum(1 for c in compare if c["결과"] != "일치")
with st.expander(f"🔍 원본 소계와 대조하기 — {'전부 일치' if not n_gap else f'{n_gap}건 차이'}"):
    st.caption("일일자금계획 시트에 **원본이 직접 적어 둔 '소계' 금액**과, 앱이 읽어들인 "
               "합계를 나란히 놓았습니다. 두 값이 같으면 빠뜨린 줄이 없다는 뜻입니다. "
               "(단위: 원. 외화 구간은 해당 통화 단위)")
    if compare:
        st.dataframe(pd.DataFrame(compare), use_container_width=True, hide_index=True)
        if n_gap:
            st.warning("차이가 나는 구간은 원본 시트에서 그 표를 직접 확인해주세요. "
                       "숨김 행이나 소계에 포함되지 않은 줄이 있을 수 있습니다.")
    else:
        st.info("소계 행을 못 찾았습니다. 원본 표 양식이 바뀌었을 수 있습니다.")

if retagged:
    with st.expander(f"🔧 계정과목 자동 교정 {len(retagged)}건 — 눌러서 확인"):
        st.caption("원본 파일의 계정과목이 실제 내용과 달라 아래처럼 바꿨습니다. "
                   "합계는 그대로이고 분류만 정확해집니다.")
        st.dataframe(pd.DataFrame(retagged), use_container_width=True, hide_index=True)

if not preview:
    st.error("읽어들인 거래가 0건입니다. 위 '원본 소계와 대조하기'에서 표를 찾았는지 확인해주세요.")
    st.stop()

st.divider()

# ── 계정과목이 있는 것 / 없는 것으로 나눔 ──
tagged = [t for t in preview if t.tag]
untagged = [t for t in preview if not t.tag]

st.subheader(f"✅ 가져올 거래 — {len(tagged):,}건")
st.caption("계정과목이 적혀 있는 거래입니다.")
st.dataframe(pd.DataFrame([{
    "시트": t.txn_date, "계정과목": t.tag, "거래처": t.counterpart,
    "적요": t.memo, "통화": t.currency, "입금": t.amount_in, "출금": t.amount_out,
} for t in tagged]), use_container_width=True, height=320)

chosen_extra = []
if untagged:
    st.subheader(f"⚠️ 계정과목 미기재 항목 — {len(untagged):,}건")
    st.caption("자금대체(MMDA)·계좌잔액처럼 집계 대상이 아닌 행일 가능성이 큽니다. "
               "**기본은 가져오지 않습니다.** 꼭 포함할 건만 아래에서 체크해주세요.")
    labels = [f"{t.txn_date} | {t.counterpart} | {t.memo} | "
              f"{'입금' if t.amount_in else '출금'} {t.amount:,.0f}원"
              for t in untagged]
    picked = st.multiselect("포함할 거래 (기본: 없음)", options=list(range(len(untagged))),
                            format_func=lambda i: labels[i], default=[])
    chosen_extra = [untagged[i] for i in picked]
    with st.expander(f"항목 {len(untagged)}건 상세 보기"):
        st.dataframe(pd.DataFrame([{
            "시트": t.txn_date, "거래처": t.counterpart, "적요": t.memo,
            "통화": t.currency, "입금": t.amount_in, "출금": t.amount_out,
        } for t in untagged]), use_container_width=True)

final = tagged + chosen_extra
st.divider()
c1, c2, c3 = st.columns(3)
c1.metric("저장할 건수", f"{len(final):,}건")
c2.metric("입금 합계", f"{sum(t.amount_in for t in final):,.0f}원")
c3.metric("출금 합계", f"{sum(t.amount_out for t in final):,.0f}원")

# ══ 유동/운용 잔고 자동 인식 ═════════════════════════════════════
st.divider()
st.subheader("💰 잔고 자동 인식")
st.caption("일일자금계획의 「3. 자금현황」 표에서 당일잔액 합계를 읽어옵니다. "
           "행 위치가 달라져도 글자를 보고 찾습니다.")

prev_sheet = bundle["prev_sheet"]
last_sheet = bundle["sheets"][-1]

open_cs = read_cash_status(bundle["rows"].get(prev_sheet, [])) if prev_sheet else {}
close_cs = read_cash_status(bundle["rows"].get(last_sheet, []))

M = 1_000_000


def _mil(v):
    return (v or 0) / M


bal = {
    "open_krw": _mil(open_cs.get("krw")), "open_usd": (open_cs.get("usd") or 0) / 1000,
    "close_krw": _mil(close_cs.get("krw")), "close_usd": (close_cs.get("usd") or 0) / 1000,
    "invest_krw": _mil(close_cs.get("invest")), "invest_usd": 0.0,
    # 기초 운용자금도 **전월 마지막 시트(예: 0630-1)** 에서 그대로 읽습니다.
    # 예전에는 지난달 산출 이력에서만 가져와, 첫 달이나 이력이 없으면 0이 됐습니다.
    "open_invest_krw": _mil(open_cs.get("invest")), "open_invest_usd": 0.0,
    "open_sheet": prev_sheet or "", "close_sheet": last_sheet,
}

c1, c2 = st.columns(2)
with c1:
    st.markdown(f"**기초잔고** — 전월말 `{prev_sheet or '못 찾음'}` 시트")
    if prev_sheet and open_cs.get("krw") is not None:
        st.metric("원화 (백만원)", f"{bal['open_krw']:,.6f}")
        st.caption(f"외화 {bal['open_usd']:,.3f}천$ · 원본 {open_cs['krw']:,.0f}원  \n"
                   f"기초 운용자금 {bal['open_invest_krw']:,.1f}백만 "
                   "(4번 화면 '기초 운용자금'에 이 값이 들어갑니다)")
    else:
        st.error("전월 마지막 시트를 못 찾았습니다. 3번 화면에서 직접 입력해주세요.")
with c2:
    st.markdown(f"**기말잔고** — 당월말 `{last_sheet}` 시트")
    if close_cs.get("krw") is not None:
        st.metric("원화 (백만원)", f"{bal['close_krw']:,.6f}")
        st.caption(f"외화 {bal['close_usd']:,.3f}천$ · 운용자금 {bal['invest_krw']:,.1f}백만")
    else:
        st.error("「3. 자금현황」 표를 못 찾았습니다.")

# 계산값(A) vs 실제값(B) 대조
calc_change = (sum(t.amount_in for t in final) - sum(t.amount_out for t in final)) / M
calc_close = bal["open_krw"] + calc_change
diff = bal["close_krw"] - calc_close
st.markdown("**대조 — 계산값 vs 파일 실제값**")
d1, d2, d3 = st.columns(3)
d1.metric("계산값 (기초 + 이번 달 수지)", f"{calc_close:,.6f}")
d2.metric("파일 실제값", f"{bal['close_krw']:,.6f}")
d3.metric("차이", f"{diff:,.6f}", delta=None)

if abs(diff) < 0.000001:
    st.success("✅ 계산값과 파일 실제값이 일치합니다.")
else:
    st.warning(f"⚠️ **{abs(diff):,.6f}백만원 차이** — 실적이 누락됐을 수 있습니다.  \n"
               "위 시트 목록에서 빠진 시트가 없는지, 계정과목이 비어 제외된 거래가 없는지 "
               "확인해주세요. 다음 달 기초잔고에는 **파일 실제값**을 사용합니다.")

st.divider()
mode = st.radio("저장 방식", ["기존 실적을 지우고 새로 저장", "기존 실적에 덧붙이기"],
                horizontal=True)

if st.button("💾 실적으로 저장하기", type="primary"):
    auto = 0
    for t in final:
        if not t.tag:
            sug = suggest_tag(t)
            if sug.level == "auto":
                t.tag = sug.tag
                auto += 1
    if mode.startswith("기존 실적을 지우고"):
        clear_period(period, "actual")
    n = save_transactions(final)
    save_month_balance(period, **bal)
    skipped = len(untagged) - len(chosen_extra)
    st.success(f"실적 {n}건을 저장했습니다."
               + (f" (계정과목이 비어 있던 {auto}건은 적요로 자동 추정)" if auto else "")
               + (f" / 계정과목 없는 {skipped}건은 제외했습니다." if skipped else ""))
    st.info("잔고도 함께 저장했습니다 — 기초 "
            f"{bal['open_krw']:,.6f}백만 / 기말 {bal['close_krw']:,.6f}백만.  \n"
            "다음은 **'2 계획 업로드'** 로 넘어가세요.")

last = bundle["last_sheet"]
st.caption(f"참고 — 가장 늦은 날짜 시트는 **{last}** 입니다 (V2 잔고 대사 기준).")
