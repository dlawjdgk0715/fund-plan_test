import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import pandas as pd
import streamlit as st

from core.nav import apply_nav
from core.data_store import (Txn, clear_period_dept, delete_transaction, dept_counts,
                             init_db, load_fund_balance, load_tags, load_transactions,
                             previous_period, save_transactions, update_transaction)
from core.engine import DISPLAY, aggregate, validate
from core.period import fmt, next_period
from core.recheck_items import check_all, check_missing, condition_text

st.set_page_config(page_title="3. 실적/계획 관리", layout="wide")
apply_nav(st)
init_db()
st.title("3️⃣ 실적 / 계획 관리")

period = st.session_state.get("period", "")
if not period:
    st.error("먼저 시작 화면(Home)에서 실적 월을 정해주세요.")
    st.stop()
st.info(f"📅 실적 = **{fmt(period)}**  /  계획 = **{fmt(next_period(period))}**")
st.caption("여기서 계정과목을 확인·수정한 뒤 아래 **[재계산]** 을 눌러주세요. "
           "고친 내용은 4번 화면에도 바로 반영됩니다.")

prev = previous_period(period)
bal = load_fund_balance(prev)
if bal:
    st.caption(f"참고 — 저장된 운용자금 잔고({prev} 기말): "
               f"원화 {bal['invest_end_krw']:,.1f}백만 / 외화 {bal['invest_end_usd']:,.1f}천$")

kind = st.radio("어느 쪽을 볼까요?",
                [f"{fmt(period)} 실적", f"{fmt(next_period(period))} 계획"],
                horizontal=True)
kind_key = "actual" if kind.endswith("실적") else "plan"
txns = load_transactions(period, kind_key, include_excluded=True)

# ── 매달 손으로 챙겨야 하는 항목 ──────────────────────────────────
# 취합 파일에 안 나오는 항목(주택자금대출 이자 지원, 삼성카드 법인카드 대금)이
# 이번 달에 하나도 없으면 **안내창**을 띄웁니다. 창 안에서 바로 넣을 수도 있습니다.
missing = check_missing(txns)


def _add_manual(item: dict, cp: str, memo: str, tag: str, amount: float, dirc: str):
    save_transactions([Txn(
        period=period, kind=kind_key, dept="", txn_date="",
        counterpart=cp, memo=memo, tag=tag, currency="KRW",
        amount_in=amount if dirc == "in" else 0.0,
        amount_out=amount if dirc == "out" else 0.0,
        source="manual",
    )])


@st.dialog("🖐 직접 넣어야 하는 항목이 있습니다")
def _recheck_dialog(items):
    st.caption(f"{kind} 자료에서 아래 항목을 못 찾았습니다. "
               "매달 취합 파일에 안 나오는 항목이라 직접 넣으셔야 합니다.  \n"
               "**해당 없는 달이면 [이번 달은 없음]** 을 누르고 넘어가세요.")
    names = sorted(load_tags().keys())
    for it in items:
        st.divider()
        st.markdown(f"**{it['label']}**")
        st.caption(f"찾은 조건: {condition_text(it)}")
        if it.get("hint"):
            st.caption(it["hint"])
        with st.form(f"rc_{it['id']}", border=False):
            a1, a2 = st.columns(2)
            cp = a1.text_input("거래처", it.get("counterpart") or "",
                               key=f"rc_cp_{it['id']}")
            memo = a2.text_input("적요", it["label"], key=f"rc_mm_{it['id']}")
            b1, b2, b3 = st.columns([2, 1, 1])
            dflt = (it.get("def_tag") or "")
            tag = b1.selectbox("계정과목", names,
                               index=names.index(dflt) if dflt in names else 0,
                               key=f"rc_tg_{it['id']}")
            amount = b2.number_input("금액 (원)", min_value=0.0, step=10000.0,
                                     format="%.0f", key=f"rc_am_{it['id']}")
            dirc = b3.selectbox("구분", ["출금", "입금"],
                                index=0 if (it.get("def_dir") or "out") == "out" else 1,
                                key=f"rc_dr_{it['id']}")
            if st.form_submit_button("➕ 이 건 추가", type="primary",
                                     use_container_width=True):
                if amount <= 0:
                    st.error("금액을 넣어주세요.")
                else:
                    _add_manual(it, cp, memo, tag,
                                float(amount), "out" if dirc == "출금" else "in")
                    st.session_state[_RC_KEY] = True
                    st.rerun()
    st.divider()
    if st.button("이번 달은 없음 — 닫기", use_container_width=True):
        st.session_state[_RC_KEY] = True
        st.rerun()


# 찾은 항목은 **어느 거래 덕분에 통과했는지** 보여줍니다.
# (없는 줄 알았는데 팝업이 안 떠서 헷갈리는 일을 막습니다)
_found = [(it, hits) for it, hits in check_all(txns) if hits]
if _found:
    with st.expander(f"🖐 직접 넣어야 하는 항목 {len(_found)}건 — 확인됨 (팝업 안 뜸)"):
        st.caption("아래 거래가 있어서 안내창이 뜨지 않습니다. "
                   "이 거래가 맞는지 한 번 봐주세요.")
        st.dataframe(pd.DataFrame([{
            "항목": it["label"], "조건": condition_text(it).replace("**", ""),
            "찾은 거래": f"{h.counterpart} | {h.memo}",
            "계정과목": h.tag or "(없음)",
            "금액": f"{h.amount:,.0f}원",
        } for it, hits in _found for h in hits]),
            use_container_width=True, hide_index=True)

_RC_KEY = f"recheck_seen_{period}_{kind_key}"
if missing:
    labels = ", ".join(it["label"] for it in missing)
    st.warning(f"**직접 넣어야 하는 항목** — {kind} 자료에 안 보입니다: {labels}  \n"
               "(해당 없는 달이면 그냥 넘어가셔도 됩니다.)")
    if not st.session_state.get(_RC_KEY):
        _recheck_dialog(missing)
    elif st.button("🖐 안내창 다시 보기"):
        st.session_state.pop(_RC_KEY, None)
        st.rerun()

if not txns:
    st.info("저장된 거래가 없습니다. 업로드 화면에서 파일을 먼저 올리거나, 아래에서 직접 추가해주세요.")

# ── 중복 점검 ─────────────────────────────────────────────────────
if txns:
    counts = dept_counts(period, kind_key)
    st.caption("지금 저장된 자료 — " +
               " · ".join(f"{d} {n}건" for d, n in sorted(counts.items())))

    groups: dict[tuple, list] = {}
    for t in txns:
        if t.excluded:
            continue
        groups.setdefault(
            (t.counterpart.strip(), t.memo.strip(), round(t.amount, 2), t.direction), []
        ).append(t)
    dupes = {k: v for k, v in groups.items() if len(v) > 1}

    if dupes:
        n_extra = sum(len(v) - 1 for v in dupes.values())
        st.error(f"🔁 **중복 의심: {n_extra}건** (같은 내용 {len(dupes)}가지)  \n"
                 "실제로 두 번 집행하는 건인지 확인한 뒤, 아니면 지워주세요.")
        with st.expander("어떤 것이 겹치는지 보기", expanded=True):
            for (cp, memo, amt, d) in sorted(dupes, key=lambda k: -k[2]):
                items = dupes[(cp, memo, amt, d)]
                depts = {i.dept or "(부서없음)" for i in items}
                cause = ("👥 **여러 부서가 같은 내용을 올렸습니다** — "
                         f"({', '.join(sorted(depts))}) 각각 집행하는 건인지 확인해주세요"
                         if len(depts) > 1 else
                         "📄 **같은 부서에 같은 내용이 여러 번 있습니다** — "
                         "파일을 두 번 올렸을 가능성이 큽니다")
                st.markdown(f"**{memo}** — {cp} / "
                            f"{'입금' if d == 'in' else '출금'} {amt:,.0f}원 "
                            f"× **{len(items)}건**  \n{cause}")
                st.dataframe(pd.DataFrame([{
                    "id": i.id, "부서": i.dept or "(없음)", "일자": i.txn_date,
                    "계정과목": i.tag, "출처": i.source,
                    "입금": i.amount_in, "출금": i.amount_out,
                } for i in items]), use_container_width=True, hide_index=True)
                keep_id = items[0].id
                if st.button(f"이 항목 1건만 남기고 {len(items)-1}건 지우기",
                             key=f"dedup_{keep_id}"):
                    for i in items[1:]:
                        delete_transaction(i.id)
                    st.success(f"{len(items)-1}건을 지웠습니다.")
                    st.rerun()

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🧹 중복 의심 건 한꺼번에 정리 — 항목마다 1건만 남김",
                         type="primary"):
                removed = 0
                for items in dupes.values():
                    for i in items[1:]:
                        delete_transaction(i.id)
                        removed += 1
                st.success(f"{removed}건을 지웠습니다.")
                st.rerun()
        with c2:
            wipe = st.selectbox("부서 하나를 통째로 지우고 다시 올리기",
                                ["(선택 안 함)"] + sorted(counts))
            if wipe != "(선택 안 함)" and st.button(f"'{wipe}' {counts[wipe]}건 지우기"):
                clear_period_dept(period, kind_key,
                                  "" if wipe == "(부서없음)" else wipe)
                st.success(f"'{wipe}' 자료를 지웠습니다.")
                st.rerun()

st.divider()
df = pd.DataFrame([{
    "id": t.id, "부서": t.dept, "일자": t.txn_date, "거래처": t.counterpart,
    "적요": t.memo, "계정과목": t.tag, "통화": t.currency,
    "입금": t.amount_in, "출금": t.amount_out, "제외": bool(t.excluded),
} for t in txns])

tag_names = sorted(load_tags().keys())
edited = st.data_editor(
    df, use_container_width=True, num_rows="dynamic", key="editor",
    column_config={
        "id": st.column_config.NumberColumn("id", disabled=True),
        "계정과목": st.column_config.SelectboxColumn(options=tag_names),
        "통화": st.column_config.SelectboxColumn(options=["KRW", "USD", "JPY", "EUR", "CNY"]),
    },
)

c1, c2 = st.columns(2)
with c1:
    if st.button("표에서 고친 내용 저장", type="primary"):
        for _, r in edited.iterrows():
            if pd.isna(r.get("id")):
                save_transactions([Txn(
                    period=period, kind=kind_key, dept=str(r.get("부서") or ""),
                    txn_date=str(r.get("일자") or ""),
                    counterpart=str(r.get("거래처") or ""), memo=str(r.get("적요") or ""),
                    tag=str(r.get("계정과목") or ""), currency=str(r.get("통화") or "KRW"),
                    amount_in=float(r.get("입금") or 0), amount_out=float(r.get("출금") or 0),
                    source="manual", excluded=int(bool(r.get("제외"))),
                )])
            else:
                update_transaction(
                    int(r["id"]), dept=str(r.get("부서") or ""),
                    txn_date=str(r.get("일자") or ""),
                    counterpart=str(r.get("거래처") or ""), memo=str(r.get("적요") or ""),
                    tag=str(r.get("계정과목") or ""), currency=str(r.get("통화") or "KRW"),
                    amount_in=float(r.get("입금") or 0), amount_out=float(r.get("출금") or 0),
                    excluded=int(bool(r.get("제외"))),
                )
        # 표에서 지운 행은 DB에서도 삭제
        left = {int(x) for x in edited["id"].dropna()}
        for t in txns:
            if t.id not in left:
                delete_transaction(t.id)
        st.success("저장했습니다.")
        st.rerun()

with c2:
    fx = st.number_input("환율(원/외화) — 외화 거래가 있을 때만", value=0.0, step=1.0)

st.divider()
if st.button("🔄 재계산", type="primary"):
    live = load_transactions(period, kind_key)
    res = aggregate(live, opening_krw=0.0, fx_rate=fx or None)
    gates = validate(live, res, fx_rate=fx or None)

    st.subheader("유동자금세부 집계 (백만원)")
    st.dataframe(pd.DataFrame(
        [{"구분": label, "금액": round(res.values.get(k, 0.0), 6)} for k, label in DISPLAY]
    ), use_container_width=True, height=500)
    st.metric("순증감 (수입 - 지출)", f"{res.values['income.total'] - res.values['expense.total']:,.6f} 백만")

    st.subheader("검증 게이트")
    for g in gates:
        (st.success if g.passed else st.error)(f"**{g.id} {g.name}** — {g.detail}")
        if g.items:
            st.write(g.items)
