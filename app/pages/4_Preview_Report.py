"""4. 최종 점검 — 보고서를 만들기 직전 수치를 검증하고 출력물을 미리 봅니다.

화면은 위·아래 두 덩어리입니다.
  ① 수치 확인   : 기초·기말 잔고 → 총자금 → 검증 게이트 (문제는 여기서 바로 고칩니다)
  ② 보고서 미리보기 : 실제 산출물 모양 그대로 + 비고 직접 수정
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import pandas as pd
import streamlit as st

from core.nav import apply_nav
from core.data_store import (clean_text, clear_gate_ack, clear_report_remarks,
                             init_db, load_gate_acks, load_month_balance,
                             load_remark_style, load_report_remarks,
                             load_report_remarks_raw, load_tags,
                             load_transactions, resolve_open_invest,
                             save_gate_ack, save_remark_style,
                             save_report_remarks, txn_fingerprint,
                             update_transaction)
from core.engine import ROW_NO, aggregate, validate
from core.period import actual_month_of, fmt, next_period, plan_month_of
from core.remark_style import learn, style_for
from core.remarks_rule import (SUMMARY1_KEYS, SUMMARY1_LABEL, SUMMARY1_ORDER,
                               cashflow_rows_from_summary1, detail_drafts,
                               mark_summary1, summary1_drafts,
                               summary1_from_cashflow)
from core.output_ui import render_output_box
from core.report import DETAIL_ROWS, TABLE_TITLE, build_report, headers

st.set_page_config(page_title="4. 최종 점검", layout="wide")
apply_nav(st)
init_db()

st.title("4️⃣ 최종 점검")
st.caption("보고서 만들기 직전 — 입력된 수치에 이상이 없는지 확인하고, 출력물을 미리 봅니다.")

period = st.session_state.get("period", "")
if not period:
    st.error("먼저 시작 화면(Home)에서 실적 월을 정해주세요.")
    st.stop()
plan_period = next_period(period)
st.info(f"📅 실적 = **{fmt(period)}**  /  계획 = **{fmt(plan_period)}**")


# ── 입력값이 페이지를 옮겨도 살아 있도록 ──────────────────────────
# 스트림릿은 화면을 옮기면 이 파일을 처음부터 다시 실행합니다. 그때 number_input 에
# value= 를 매번 새로 주면 사용자가 고친 값이 기본값으로 되돌아갑니다.
# → key 를 붙이고 **처음 한 번만** 기본값을 넣어 둡니다.
def keep(col, label, key: str, default: float, **kw):
    if key not in st.session_state:
        st.session_state[key] = float(default)
    return col.number_input(label, key=key, **kw)


bal = load_month_balance(period)
if not bal:
    st.warning("저장된 잔고가 없습니다. **1 실적 업로드**에서 일일자금계획 파일을 먼저 올려주세요. "
               "아래에서 직접 입력해도 됩니다.")

# 기초 운용자금 = **전월말** 값. 일일자금계획 전월 마지막 시트(예: 0630-1)에서 읽습니다.
inv_dflt, inv_dflt_usd, inv_src = resolve_open_invest(period)

# ══════════════════════════════════════════════════════════════════
# ①  수 치  확 인
# ══════════════════════════════════════════════════════════════════
st.header("① 수치 확인")

with st.expander("❓ 기초잔고는 어떻게 정해지나요?"):
    st.markdown(f"""
잔고는 **1 실적 업로드**에서 일일자금계획 파일을 읽을 때 자동으로 잡힙니다.

| 무엇 | 어디서 오나 |
|---|---|
| **{fmt(period)} 기초잔고** | 전월 마지막 시트 「3. 자금현황」 당일잔액 합계 |
| **{fmt(period)} 기말잔고** | 당월 마지막 시트 「3. 자금현황」 당일잔액 합계 (실제값) |
| **{fmt(plan_period)} 기초잔고** | 위 **실제값을 그대로 이어받습니다** |

실적 기말은 두 가지로 나옵니다.
**계산값** = 기초 + 이번 달 거래 수지 / **실제값** = 파일에 적힌 잔고.
둘이 다르면 실적이 누락됐을 수 있어 경고가 뜨고, 다음 달 기초에는 **실제값**을 씁니다.
""")

st.subheader("기초 잔고 · 환율")
c1, c2, c3, c4 = st.columns(4)
open_krw = keep(c1, f"{fmt(period)} 기초 유동잔고 원화 (백만원)", f"pr_open_krw_{period}",
                float(bal["open_krw"]) if bal else 0.0, format="%.6f")
open_usd = keep(c2, "기초 유동잔고 외화 (천$)", f"pr_open_usd_{period}",
                float(bal["open_usd"]) if bal else 0.0, format="%.3f")
inv_open = keep(c3, "기초 운용자금 원화 (백만원)", f"pr_inv_krw_{period}",
                inv_dflt, format="%.6f", help=inv_src)
fx = keep(c4, "환율 (원/외화)", f"pr_fx_{period}", 0.0, step=1.0)
st.caption(f"💰 기초 운용자금 출처: {inv_src} · "
           "이 칸의 값은 화면을 옮겼다 돌아와도 그대로 남습니다.")

actual = load_transactions(period, "actual")
plan = load_transactions(period, "plan")

# ── 실적 기말: 계산값 vs 실제값 ───────────────────────────────────
res_a = aggregate(actual, opening_krw=open_krw, opening_usd=open_usd, fx_rate=fx or None)
close_calc = res_a.closing
close_real = float(bal["close_krw"]) if bal else None

st.subheader("기말 잔고 대조")
m1, m2, m3 = st.columns(3)
m1.metric(f"{fmt(period)} 기초 유동잔고", f"{open_krw:,.2f}")
m2.metric("기말 — 계산값", f"{close_calc:,.2f}")
if close_real is not None:
    diff = close_real - close_calc
    m3.metric("기말 — 파일 실제값", f"{close_real:,.2f}",
              delta=f"{diff:+,.2f}" if abs(diff) >= 1e-6 else "일치")
    if abs(diff) >= 1e-6:
        st.warning(f"⚠️ **{abs(diff):,.2f}백만원 차이** — 실적 거래가 누락됐을 수 있습니다. "
                   "**1 실적 업로드**에서 시트가 빠지지 않았는지, 계정과목이 비어 제외된 "
                   "거래가 없는지 확인해주세요.")
    else:
        st.success("✅ 계산값과 파일 실제값이 일치합니다.")
else:
    m3.metric("기말 — 파일 실제값", "—")

# ── 계획월 기초 유동잔고 ──────────────────────────────────────────
# 원칙: 다음 달 기초 = 이번 달 기말(파일 실제값). 처음엔 자동으로 채워집니다.
# 그런데 이 칸은 화면을 옮겨도 값이 남기 때문에, 실적을 고쳐 기말이 바뀌면
# 옛 값이 그대로 남을 수 있습니다. 그때만 '맞추기' 버튼을 보여줍니다.
#
# ⚠️ 버튼 안에서 st.session_state[키] = … 로 바꾸면
#    "cannot be modified after the widget ... is instantiated" 오류가 납니다
#    (이미 위에서 number_input 이 만들어진 뒤라서). on_click 콜백은 화면을
#    다시 그리기 **전에** 실행되므로 안전합니다.
_PLANOPEN_KEY = f"pr_planopen_{period}"
_want_open = float(close_real if close_real is not None else close_calc)


def _sync_plan_open():
    st.session_state[_PLANOPEN_KEY] = _want_open


_cur_open = st.session_state.get(_PLANOPEN_KEY, _want_open)
_off = abs(float(_cur_open) - _want_open) >= 1e-6

p1, p2 = st.columns([3, 1])
plan_open = keep(p1, f"{fmt(plan_period)} 기초 유동잔고 원화 (백만원)",
                 _PLANOPEN_KEY, _want_open, format="%.6f",
                 help="원칙: 다음 달 기초잔고 = 이번 달 기말잔고(파일 실제값)")
with p2:
    st.write("")
    if _off:
        st.button(f"🔄 기말잔고({_want_open:,.2f})로 맞추기",
                  on_click=_sync_plan_open, use_container_width=True)
if _off:
    st.caption(f"⚠️ 이번 달 기말잔고 **{_want_open:,.2f}** 와 다릅니다. "
               "일부러 다르게 넣으신 게 아니면 위 버튼을 눌러주세요.")
plan_open_usd = float(bal["close_usd"]) if bal else 0.0

rep = build_report(actual, plan, actual_month_of(period), plan_month_of(period),
                   open_krw=open_krw, open_usd=open_usd,
                   plan_open_krw=plan_open, plan_open_usd=plan_open_usd,
                   invest_open_krw=inv_open, invest_open_usd=inv_dflt_usd, fx=fx or 0.0)
hdr = headers(rep)
st.session_state["report_inputs"] = {
    "open_krw": open_krw, "open_usd": open_usd,
    "plan_open_krw": plan_open, "plan_open_usd": plan_open_usd,
    "invest_open_krw": inv_open, "invest_open_usd": inv_dflt_usd, "fx": fx or 0.0,
}
for _w in rep.meta.get("structure_warnings", []):
    st.warning("⚠️ " + _w)

# ── 유동 · 운용 · 총자금 ──────────────────────────────────────────
st.subheader("유동자금 · 운용자금 · 총자금")
q1, q2, q3 = st.columns(3)
q1.metric(f"{fmt(period)} 말 유동자금", f"{rep.meta['actual_liquid_close']:,.2f}백만",
          help="세부내역 G. 기말유동잔고와 같은 값입니다.")
q2.metric(f"{fmt(period)} 말 운용자금", f"{rep.meta['invest_actual_end']:,.2f}백만")
q3.metric(f"{fmt(period)} 말 총자금", f"{rep.meta['actual_close_total']:,.2f}백만")
r1, r2, r3 = st.columns(3)
r1.metric(f"{fmt(plan_period)} 말 유동자금", f"{rep.meta['plan_liquid_close']:,.2f}백만")
r2.metric(f"{fmt(plan_period)} 말 운용자금", f"{rep.meta['invest_plan_end']:,.2f}백만")
r3.metric(f"{fmt(plan_period)} 말 총자금", f"{rep.meta['plan_close_total']:,.2f}백만",
          delta=f"{rep.meta['diff_total']:+,.2f}")

# ══ 검증 게이트 ═══════════════════════════════════════════════════
st.subheader("🔎 검증 게이트 6종")
st.caption("빨간 칸은 사람이 봐야 할 곳입니다. 문제 거래는 **이 화면에서 바로 고칠 수 있고**, "
           "고치면 3번 화면에도 즉시 반영됩니다. "
           "고칠 게 없는 항목은 [확인함]을 눌러 초록색으로 바꿔주세요.")

TAG_OPTIONS = [""] + sorted(load_tags().keys())
if st.session_state.pop("_fix_msg", None):
    st.success("거래를 고쳤습니다. 3번 화면에도 바로 반영됩니다.")


def fix_panel(scope: str, res, key_prefix: str) -> bool:
    """문제 거래를 이 화면에서 바로 고치는 표. 고쳤으면 True."""
    bad: dict[int, object] = {}
    for t in list(res.unmapped) + list(res.unregistered):
        if t.id:
            bad[t.id] = t
    if not bad:
        st.caption("고칠 거래가 없습니다.")
        return False
    # 거래 id 는 표에 보이지 않게 '행 이름'으로 넣어 둡니다.
    df = pd.DataFrame(
        [{"시트/부서": t.txn_date or t.dept, "거래처": t.counterpart,
          "적요": t.memo, "방향": "입금" if t.direction == "in" else "출금",
          "통화": t.currency, "금액": float(t.amount), "계정과목": t.tag or "",
          "집계 제외": bool(t.excluded)} for t in bad.values()],
        index=list(bad.keys()))
    ed = st.data_editor(
        df, hide_index=True, use_container_width=True,
        disabled=["시트/부서", "거래처", "적요", "방향", "통화", "금액"],
        column_config={
            "계정과목": st.column_config.SelectboxColumn(options=TAG_OPTIONS, required=False),
            "금액": st.column_config.NumberColumn(format="%.0f"),
        },
        key=f"fix_{key_prefix}")
    if st.button("💾 고친 내용 저장", key=f"fixsave_{key_prefix}", type="primary"):
        n = 0
        for tid, row in ed.iterrows():
            t = bad[int(tid)]
            new_tag = str(row["계정과목"] or "").strip()
            new_exc = 1 if row["집계 제외"] else 0
            if new_tag != (t.tag or "") or new_exc != int(t.excluded):
                update_transaction(int(tid), tag=new_tag, excluded=new_exc)
                n += 1
        st.session_state["_fix_msg"] = f"{n}건을 고쳤습니다. 3번 화면에도 바로 반영됩니다."
        st.rerun()
    return False


for scope, label, txns, res, ext in (
        ("actual", f"{fmt(period)} 실적", actual, res_a, close_real),
        ("plan", f"{fmt(plan_period)} 계획", plan, rep.plan, None)):
    if not txns:
        continue
    st.markdown(f"**{label}**")
    fp = txn_fingerprint(txns)
    acks = load_gate_acks(period, scope)
    gates = validate(txns, res, external_balance=ext, fx_rate=fx or None)

    cols = st.columns(3)
    for i, g in enumerate(gates):
        acked = acks.get(g.id, {}).get("fingerprint") == fp
        stale = (not g.passed) and g.id in acks and not acked
        with cols[i % 3]:
            if g.passed:
                st.success(f"**{g.id} {g.name}**\n\n특이사항 없음")
            elif acked:
                st.success(f"**{g.id} {g.name}**\n\n특이사항 없음 "
                           f"(사용자 확인 {acks[g.id]['updated_at'][:16]})")
            else:
                st.error(f"**{g.id} {g.name}**\n\n{g.detail}")
                if stale:
                    st.caption("⟳ 확인 이후 거래가 바뀌어 확인이 풀렸습니다.")
                if g.items:
                    st.caption(" / ".join(str(x) for x in g.items[:5]))
                with st.popover("확인 / 수정", use_container_width=True):
                    st.markdown(f"**{g.id} {g.name}** — {g.detail}")
                    if g.id in ("V1", "V4"):
                        fix_panel(scope, res, f"{scope}_{g.id}")
                    elif g.items:
                        st.write(g.items)
                    st.divider()
                    if st.button("✅ 확인함 — 문제 없습니다",
                                 key=f"ack_{scope}_{g.id}", type="primary"):
                        save_gate_ack(period, scope, g.id, fp)
                        st.rerun()
            if (g.passed is False) and acked:
                if st.button("확인 취소", key=f"unack_{scope}_{g.id}"):
                    clear_gate_ack(period, scope, g.id)
                    st.rerun()

if bal:
    st.caption(f"잔고 출처 — 기초: `{bal['open_sheet']}` 시트 · 기말: `{bal['close_sheet']}` 시트 "
               f"· 당월말 운용자금 {bal['invest_krw']:,.2f}백만원")

# ══════════════════════════════════════════════════════════════════
# ②  보 고 서  미 리 보 기
# ══════════════════════════════════════════════════════════════════
st.divider()
st.header("② 보고서 미리보기")
st.caption("실제 산출물(PPT·엑셀)에 들어갈 모양 그대로입니다. "
           "**비고는 표 아래에서 직접 고치고 [비고 저장]을 누르면 산출물에 그대로 들어갑니다.**")


def dash(v):
    if isinstance(v, str):
        return v or "&nbsp;"
    if v is None or abs(v) < 0.05:
        return "-"
    return f"{v:,.1f}"


def signed(v):
    if isinstance(v, str):
        return v or "&nbsp;"
    if v is None or abs(v) < 0.05:
        return "-"
    return f"{v:+,.1f}"


CSS = """
<style>
.rpt {border-collapse:collapse; width:100%; font-size:13px;}
.rpt th {background:#dce6f1; border:1px solid #a8b8cc; padding:7px; text-align:center;
         font-weight:600;}
.rpt td {border:1px solid #c9c9c9; padding:6px 9px;}
.rpt td.lbl {background:#fafcff; font-weight:600; white-space:nowrap;}
.rpt td.sub {background:#fafcff; color:#555; text-align:center; white-space:nowrap;}
.rpt td.num {text-align:right; font-variant-numeric:tabular-nums;}
.rpt td.neg {color:#1f4fd8;}
.rpt td.pos {color:#d92b2b;}
.rpt td.note {font-size:12px; color:#333;}
.rpt tr.last td {background:#fffbe6; font-weight:700;}
.rpt tr.last td.note {font-weight:400;}
.rpt th.hide {display:none;}
.rpt td.lbl {white-space:pre;}
</style>
"""


def render(head, rows, remarks, sub_col=False, diff_col=None, note_w="32%",
           group=None, ncols=None):
    """ncols = 값 열을 몇 개까지 그릴지 (None이면 전부)."""
    if group:
        th = ("<tr>" + "".join(group) + "</tr>"
              "<tr>" + "".join(f"<th>{h}</th>" for h in head[1:]) + "</tr>")
    else:
        th = "<tr>" + "".join(f"<th>{h}</th>" for h in head) + \
             f"<th style='width:{note_w}'>비고</th></tr>"
    body = []
    for i, row in enumerate(rows):
        tds = [f"<td class='lbl'>{row[0] or '&nbsp;'}</td>"]
        start = 1
        if sub_col:
            tds.append(f"<td class='sub'>{row[1] or '&nbsp;'}</td>")
            start = 2
        stop = len(row) if ncols is None else min(len(row), start + ncols)
        for j, v in enumerate(row[start:stop], start=start):
            if diff_col is not None and j == diff_col:
                txt, cls = signed(v), ("neg" if isinstance(v, (int, float)) and v < -0.05
                                       else ("pos" if isinstance(v, (int, float)) and v > 0.05
                                             else ""))
            else:
                txt = dash(v)
                cls = "neg" if isinstance(v, (int, float)) and v < -0.05 else ""
            tds.append(f"<td class='num {cls}'>{txt}</td>")
        note = (remarks[i] if i < len(remarks) else "") or ""
        tds.append(f"<td class='note'>{note.replace(chr(10), '<br>')}</td>")
        cls = " class='last'" if i == len(rows) - 1 else ""
        body.append(f"<tr{cls}>" + "".join(tds) + "</tr>")
    return CSS + f"<table class='rpt'><thead>{th}</thead><tbody>{''.join(body)}</tbody></table>"


def flat(s: str) -> str:
    """편집표는 한 줄짜리 칸이라, 여러 줄 초안을 ' / ' 로 이어 보여줍니다."""
    return " / ".join(x.strip() for x in (s or "").split("\n") if x.strip())


# ── 비고 초안 (과거 PPT 서식을 따라 앱이 미리 만들어 둡니다) ──────
# 요약1(엑셀 1-1)과 '1. 자금수지' 표는 **같은 네 문장**을 씁니다.
# 그래서 사용자는 네 문장만 손보면 됩니다.
S1_DRAFT = summary1_drafts(actual, plan, rep.actual_month, rep.plan_month,
                           style_for("summary1", load_remark_style("summary1")), rep)
_s1_saved = load_report_remarks_raw(period, "summary1", len(SUMMARY1_KEYS))
# 7차 이전 버전에서 '자금수지' 표에 직접 써 두신 글도 이어받습니다
_legacy = summary1_from_cashflow(load_report_remarks(period, "cash_flow", 6))
# 저장한 적 없으면(None) 초안 · 일부러 비웠으면("") 빈칸 그대로
S1_NOW = {k: (_s1_saved[i] if _s1_saved[i] is not None
              else (_legacy.get(k, "") or S1_DRAFT[k]))
          for i, k in enumerate(SUMMARY1_KEYS)}

DRAFTS = {
    "detail": [flat(x) for x in detail_drafts(
        res_a, rep.plan, rep.actual_month, rep.plan_month,
        style_for("detail", load_remark_style("detail")))],
}

KEYS = ("cash_flow", "liquid_plan", "fund_plan", "stock", "detail")
tabs = st.tabs([TABLE_TITLE[k].format(a=rep.actual_month, p=rep.plan_month) for k in KEYS])

for tab, key in zip(tabs, KEYS):
    with tab:
        rows, head = rep.tables[key], hdr[key]

        saved = load_report_remarks_raw(period, key, len(rows))
        drafts = DRAFTS.get(key, [""] * len(rows))
        # None = 아직 저장 안 함 → 초안 / "" = 일부러 비움 → 빈칸 그대로
        shown = [(drafts[i] if i < len(drafts) else "") if s is None else s
                 for i, s in enumerate(saved)]
        if key == "cash_flow":
            shown = cashflow_rows_from_summary1(S1_NOW)

        if key == "detail":
            # b-a 열은 화면에서 빼고, 비고를 마지막 열로 둡니다.
            group = [
                "<th rowspan='2'>구분</th>",
                f"<th colspan='3'>{rep.actual_month}월(실적)</th>",
                f"<th colspan='3'>{rep.plan_month}월(계획)</th>",
                "<th rowspan='2' style='width:26%'>비고</th>",
            ]
            st.caption("(단위 : 백만원, 천$)")
            st.markdown(render(head[:7], rows, shown, group=group, ncols=6),
                        unsafe_allow_html=True)
        else:
            sub = key in ("fund_plan", "stock")
            diff = (len(head) - 1) if key in ("cash_flow", "stock") else None
            st.markdown(render(head, rows, shown, sub_col=sub, diff_col=diff),
                        unsafe_allow_html=True)

        # ── 자금수지 표: 네 문장만 고치면 요약1·PPT가 같이 바뀝니다 ──
        if key == "cash_flow":
            st.markdown("**✏️ 비고 — 원본 PPT 서식 그대로 (● 실적월 / ○ 계획월)**")
            st.caption("여기서 확정한 문장이 **엑셀 요약1**과 **PPT 2장 표**에 그대로 들어갑니다 "
                       ""
                       "● / ○ 기호는 안 붙이셔도 저장할 때 자동으로 붙습니다.")
            new4 = {}
            for k in SUMMARY1_ORDER:
                new4[k] = st.text_input(
                    SUMMARY1_LABEL[k].format(a=rep.actual_month, p=rep.plan_month),
                    S1_NOW[k], key=f"s1_{k}_{period}")
            new4 = {k: mark_summary1(k, v) for k, v in new4.items()}
            g1, g2, _g3 = st.columns([1, 1, 3])
            if g1.button("💾 비고 저장", key="s1save", type="primary",
                         use_container_width=True):
                texts = [new4[k] for k in SUMMARY1_KEYS]
                save_report_remarks(period, "summary1", texts)
                # 자금수지 표(PPT 2장)에도 같은 문장을 넣어 둡니다
                save_report_remarks(period, "cash_flow", cashflow_rows_from_summary1(new4))
                msg = "비고를 저장했습니다. 요약1·PPT에 그대로 들어갑니다."
                learned = learn("summary1", [S1_DRAFT[k] for k in SUMMARY1_KEYS],
                                texts, rep.actual_month)
                if learned:
                    save_remark_style("summary1", learned)
                    msg += " 쓰시는 방식도 배웠습니다."
                st.session_state["_rm_msg"] = msg
                st.rerun()
            if g2.button("↩ 초안으로 되돌리기", key="s1reset", use_container_width=True):
                clear_report_remarks(period, "summary1")
                clear_report_remarks(period, "cash_flow")
                for k in SUMMARY1_KEYS:
                    st.session_state.pop(f"s1_{k}_{period}", None)
                st.rerun()
            continue

        st.markdown("**✏️ 비고 고치기** — 칸을 고친 뒤 아래 [비고 저장]을 누르세요.")
        if key in DRAFTS:
            st.caption("앱이 과거 PPT 서식대로 초안을 채워 뒀습니다. 자유롭게 고치시면 "
                       "**쓰시는 방식**(금액 단위·소수 자릿수·항목 개수·구분 기호)을 배워 "
                       "다음 달 초안에 반영합니다. 항목이 여러 개면 ` / ` 로 나눠 적어주세요.")
        if key == "detail":
            st.caption("💡 **굵은 소계 줄**(수입·지출·인건비·기타 등)의 비고는 PPT에만 들어갑니다. "
                       "엑셀 유동자금세부 N열에는 **맨 아래 단계 항목**의 비고만 들어갑니다.")
        cols = {"구분": [str(r[0]).replace("　", "  ") for r in rows]}
        if key == "detail":
            # 엑셀 유동자금세부 N열에 들어가는 줄만 표시 (소계 줄은 PPT 전용)
            cols["엑셀"] = ["○" if k in ROW_NO else "—" for _, _, k in DETAIL_ROWS]
        cols["비고"] = shown

        # ⚠️ 편집표를 form 으로 감쌉니다.
        #    form 밖에 두면, 칸에 글자를 쓰고 **곧바로 [저장]을 누를 때** 그 칸의 값이
        #    아직 서버로 안 넘어와 고친 내용이 통째로 빠집니다(실제로 겪은 문제).
        #    form 은 [저장]을 누르는 순간 표 전체를 한꺼번에 보냅니다.
        with st.form(f"rmform_{key}_{period}", border=False):
            ed = st.data_editor(
                pd.DataFrame(cols),
                hide_index=True, use_container_width=True, disabled=["구분", "엑셀"],
                column_config={"구분": st.column_config.TextColumn(width="medium"),
                               "엑셀": st.column_config.TextColumn(
                                   width="small", help="○ = 엑셀 유동자금세부 N열에도 들어감"),
                               "비고": st.column_config.TextColumn(width="large")},
                key=f"rmed_{key}_{period}_{st.session_state.get(f'rmver_{key}', 0)}")
            f1, _f2 = st.columns([1, 4])
            saved_now = f1.form_submit_button("💾 비고 저장", type="primary",
                                              use_container_width=True)

        if saved_now:
            # 칸을 지우면 판다스가 NaN 을 주는데, 그냥 str() 하면 "nan" 이 됩니다.
            texts = [clean_text(x) for x in ed["비고"].tolist()]
            save_report_remarks(period, key, texts)
            n_filled = sum(1 for x in texts if x.strip())
            msg = f"비고 {n_filled}줄을 저장했습니다. 표와 산출물에 바로 반영됩니다."
            if key in DRAFTS:
                learned = learn(key, drafts, texts, rep.actual_month)
                if learned:
                    save_remark_style(key, learned)
                    msg += " 쓰시는 방식도 배웠습니다."
            st.session_state["_rm_msg"] = msg
            st.rerun()

        if key in DRAFTS and st.button("↩ 초안으로 되돌리기", key=f"rmreset_{key}"):
            clear_report_remarks(period, key)      # 지워야 '저장 안 함'으로 되돌아갑니다
            # 편집표가 옛 내용을 붙들고 있지 않도록 위젯을 새로 만듭니다
            st.session_state[f"rmver_{key}"] = st.session_state.get(f"rmver_{key}", 0) + 1
            st.rerun()

        _now = load_report_remarks_raw(period, key, len(rows))
        n_edit = sum(1 for i, x in enumerate(_now)
                     if x is not None and x != (drafts[i] if i < len(drafts) else ""))
        st.caption(f"✅ 수정내역 {n_edit}건 전부 반영" if n_edit
                   else "초안 그대로 나갑니다 (고치신 내용 없음)")

_msg = st.session_state.pop("_rm_msg", None)
if _msg:
    st.toast(_msg, icon="💾")

# ══════════════════════════════════════════════════════════════════
# ③  엑 셀 · P P T  받 기   (화면을 옮기지 않고 여기서 바로)
# ══════════════════════════════════════════════════════════════════
st.divider()
st.header("③ 엑셀·PPT 받기")
st.caption("비고만 고치고 바로 받으실 수 있습니다. 내용이 그대로면 지난 파일을 즉시 드립니다.")
render_output_box(period, st.session_state["report_inputs"],
                  rep.actual_month, rep.plan_month, key="p4")
