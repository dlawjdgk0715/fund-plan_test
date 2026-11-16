"""자금운용계획 자동화 — 시작 화면.
실행:  python -m streamlit run app/Home.py
"""
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from core import gitstore
from core.data_store import annual_saved_months, init_db, load_fund_balance, previous_period
from core.period import actual_month_of, fmt, next_period, plan_month_of
from core.nav import apply_nav, set_show_optional, show_optional
from core.outputs import inputs_stamp, last_outputs
from core.progress import auto_vals, first_todo, steps
from core.uploads import load_state, save_state

st.set_page_config(page_title="자금운용계획 자동화", page_icon="💰",
                   layout="wide", initial_sidebar_state="expanded")
init_db()
apply_nav(st)

st.title("💰 자금운용계획 자동화")
st.caption("㈜카탈리스터 전략금융실 — 매달 자금운용계획 보고서를 자동으로 만듭니다.")

# 실적 월은 cmd 창을 껐다 켜도 지난번 값을 그대로 씁니다.
if "period" not in st.session_state:
    saved = load_state("period")
    if saved:
        st.session_state.period = str(saved)
    else:
        d = date.today()
        prev = f"{d.year}-{d.month:02d}"
        st.session_state.period = previous_period(prev)   # 보통 전월 실적으로 작업

col1, col2 = st.columns([1, 2])
with col1:
    st.session_state.period = st.text_input(
        "📊 실적 월 (YYYY-MM)", st.session_state.period,
        help="이번에 만들 보고서의 '실적' 기준 월입니다. 계획 월은 자동으로 다음 달이 됩니다.",
    )
with col2:
    st.write("")
    try:
        st.success(
            f"**{fmt(st.session_state.period)} 실적** + **{fmt(next_period(st.session_state.period))} 계획** "
            "으로 보고서를 만듭니다."
        )
    except Exception:
        st.error("월 형식이 잘못됐습니다. `2026-07` 처럼 적어주세요.")
        st.stop()

save_state("period", st.session_state.period)
am, pm = actual_month_of(st.session_state.period), plan_month_of(st.session_state.period)

st.divider()

EXTRA_PAGES = [
    ("pages/6_Remark_Labels.py",  "비고 라벨 관리", "비고에 쓸 짧은 라벨 수정"),
    ("pages/7_Tag_Dictionary.py", "태그 사전 관리", "계정과목 자동 분류 기준 관리"),
]

base = Path(__file__).resolve().parent

# ── 진행 현황 ─────────────────────────────────────────────────────
# 사용자가 체크할 필요 없습니다. 이미 저장된 자료에서 매번 계산합니다.
# (따로 기록해 두면 실제 자료와 어긋날 수 있어서 계산으로만 알아냅니다)
st.markdown("""
<style>
@keyframes fpblink { 0%,100% {background:#fdecea;} 50% {background:#ffffff;} }
.fp-todo {animation: fpblink 1.6s ease-in-out infinite; border:1px solid #f0a9a0;
          border-radius:6px; padding:7px 10px; color:#a1160a; font-size:13px;}
.fp-ok   {padding:7px 10px; color:#4b6b4b; font-size:13px;}
</style>
""", unsafe_allow_html=True)

_last = last_outputs(st.session_state.period)
_ri = st.session_state.get("report_inputs")
_vals = _ri or auto_vals(st.session_state.period)
_fresh = bool(_last) and (
    inputs_stamp(st.session_state.period, _vals, am, pm) == _last["stamp"])
_steps = steps(st.session_state.period, _vals, output_ready=_fresh)
_done = sum(1 for s in _steps if s.done)

h1, h2 = st.columns([3, 2])
h1.subheader("진행 현황")
h2.metric("완료한 단계", f"{_done} / {len(_steps)}")
st.progress(_done / len(_steps))
_todo = first_todo(_steps)
if _todo:
    st.caption(f"빨갛게 깜빡이는 칸이 아직 남은 일입니다. 지금 하실 일 → **{_todo.label.strip()}**")
else:
    st.caption("모든 단계가 끝났습니다. 아래에서 파일을 받으세요.")

for s in _steps:
    with st.container(border=True):
        if (base / s.page).exists():
            try:
                st.page_link(s.page, label=s.label, icon="✅" if s.done else "⬜")
            except Exception:
                st.write(f"**{s.label}** — 왼쪽 사이드바에서 여세요")
        else:
            st.error(f"{s.label} — 파일 없음")
        st.markdown(
            f"<div class='{'fp-ok' if s.done else 'fp-todo'}'>{s.detail}</div>",
            unsafe_allow_html=True)

# ── 가끔 쓰는 화면 (평소엔 사이드바에서 숨김) ─────────────────────
_extra = st.toggle("가끔 쓰는 화면도 보기 (비고 라벨 · 태그 사전)", value=show_optional())
if _extra != show_optional():
    set_show_optional(_extra)
    st.rerun()
if _extra:
    xcols = st.columns(3)
    for i, (path, label, desc) in enumerate(EXTRA_PAGES):
        with xcols[i % 3]:
            if (base / path).exists():
                try:
                    st.page_link(path, label=label, icon="🔧")
                except Exception:
                    st.write(f"**{label}**")
                st.caption(desc)

# ── 이미 만들어 둔 파일 바로 받기 ─────────────────────────────────
# 아무것도 안 고치고 파일만 다시 받고 싶을 때, 4번까지 안 가도 됩니다.
st.divider()
st.subheader("📥 만들어 둔 파일 받기")
if not _last:
    st.info("아직 만든 파일이 없습니다. 위 진행 현황에서 남은 단계를 마치고 만들어주세요.")
    try:
        st.page_link("pages/4_Preview_Report.py",
                     label="4 최종 점검 · 산출 에서 만들기", icon="📄")
    except Exception:
        pass
else:
    if not _fresh:
        st.warning("이 파일을 만든 뒤로 **내용이 바뀌었습니다.** "
                   "최신으로 받으시려면 4번 화면에서 다시 만들어주세요.")
    dcols = st.columns(len(_last["files"]))
    for col, (label, path, mime) in zip(dcols, _last["files"]):
        col.download_button(f"⬇ {label} 받기", path.read_bytes(),
                            file_name=path.name, mime=mime,
                            use_container_width=True, key=f"home_dl_{label}")
    st.caption(f"만든 시각: {_last['made_at'] or '기록 없음'} · 저장 위치: "
               f"`{_last['files'][0][1].parent}`")

# ── 자료 보관 상태 ────────────────────────────────────────────────
# 웹(Streamlit Cloud)은 앱을 다시 띄울 때 디스크를 되돌립니다.
# 그래서 저장한 내용을 깃허브에 올려 두고, 시작할 때 다시 내려받습니다.
st.divider()
_gs = gitstore.status()
with st.expander(
        ("☁️ 자료 보관 — " + ("저장 안 된 변경 있음" if _gs["dirty"] else "최신"))
        if _gs["enabled"] else "💾 자료 보관 — 이 PC에만 저장됩니다",
        expanded=bool(_gs["enabled"] and _gs["dirty"])):
    if not _gs["enabled"]:
        st.warning(
            "**지금은 이 컴퓨터에만 저장됩니다.**\n\n"
            "웹(Streamlit Cloud)에서 쓰실 때는 앱을 다시 띄울 때마다 입력한 자료가 "
            "사라집니다. 아래를 앱 설정의 **Secrets** 에 넣어주세요.",
            icon="⚠️")
        st.code('FUNDPLAN_GH_TOKEN  = "github_pat_여기에"\n'
                f'FUNDPLAN_GH_REPO   = "{_gs["repo"]}"\n'
                f'FUNDPLAN_GH_BRANCH = "{_gs["branch"]}"', language="toml")
    else:
        c1, c2, c3 = st.columns([2, 2, 1])
        c1.metric("마지막 불러오기", _gs["pulled_at"] or "—")
        c2.metric("마지막 저장", _gs["pushed_at"] or "—")
        if c3.button("☁ 지금 저장", use_container_width=True):
            with st.spinner("저장소에 올리는 중…"):
                r = gitstore.flush("수동 저장")
            (st.success if r.get("ok") else st.error)(r.get("msg", ""))
            st.rerun()
        if _gs["dirty"]:
            st.warning("아직 저장소에 안 올라간 변경이 있습니다. "
                       "잠시 뒤 자동으로 올라가며, 급하면 [☁ 지금 저장]을 눌러주세요.")
        else:
            st.success("저장소와 같은 상태입니다.")
        st.caption(f"보관 위치 `{_gs['repo']}` · 브랜치 `{_gs['branch']}` · "
                   f"`{_gs['path']}`")
    if _gs["last_msg"]:
        st.caption(f"최근 기록: {_gs['last_at']} — {_gs['last_msg']}")

    _saved = annual_saved_months(int(str(st.session_state.period)[:4]))
    if _saved:
        st.caption("연간 시트에 보관된 달: "
                   + ", ".join(f"{m}월" for m in _saved))

st.divider()
prev = previous_period(st.session_state.period)
bal = load_fund_balance(prev)
if bal:
    st.info(f"저장된 운용자금 잔고 ({fmt(prev)} 기말): "
            f"원화 {bal['invest_end_krw']:,.1f}백만 / 외화 {bal['invest_end_usd']:,.1f}천$")
else:
    st.warning(f"{fmt(prev)} 운용자금 기말 잔고 기록이 없습니다. "
               "첫 달은 '4 최종 점검 · 산출' 화면에서 직접 입력해주세요.")

with st.expander("❓ '실적 월'이 뭔가요?"):
    st.markdown(f"""
보고서 한 세트(실적 + 계획)를 묶는 **이름표**입니다. 실적 월만 정하면 나머지는 자동입니다.

| 정한 값 | 자동으로 정해지는 것 |
|---|---|
| 실적 월 = **{st.session_state.period}** | 실적 = **{am}월** / 계획 = **{pm}월** |
| | 운용자금 기초잔고 = **{fmt(prev)}** 기말값을 이어받음 |

예를 들어 **{pm}월 계획을 짜려면 실적 월에 `{st.session_state.period}`** 를 적으면 됩니다.
1번 화면은 {am}월 시트를, 3번 화면은 {pm}월 계획 파일을 받습니다.
""")

with st.expander("🔧 화면이 안 보이거나 오류가 날 때 — 이렇게 다시 켜세요"):
    st.markdown("**1단계.** 아래 폴더를 탐색기에서 연 뒤, 주소창에 `cmd` 를 치고 Enter")
    st.code(str(base.parent), language=None)
    st.markdown("**2단계.** 열린 검은 창에 아래를 붙여넣고 Enter")
    st.code("python -m streamlit run app/Home.py", language=None)
    st.markdown("팀원과 같이 쓸 때는 이걸 쓰세요")
    st.code("python -m streamlit run app/Home.py --server.address 0.0.0.0", language=None)
    st.caption("※ start 앱(바로가기)이 아니라 반드시 위 폴더에서 cmd 창으로 실행해야 합니다. "
               "`core` 폴더 파일을 바꿨다면 cmd 창을 껐다 켜야 반영됩니다(F5로는 안 됨).")
