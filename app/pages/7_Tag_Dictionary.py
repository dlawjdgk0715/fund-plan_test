import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import pandas as pd
import streamlit as st

from core.nav import apply_nav
from core.data_store import (init_db, load_tags, set_tag_examples, set_tag_flags,
                             tag_usage_counts)
from core.engine import ACCOUNT_MAP, ANNUAL_GROUP, EXCLUDED_ACCOUNTS, ROW_LABEL

st.set_page_config(page_title="7. 태그 사전 관리", layout="wide")
apply_nav(st)
init_db()
st.title("7️⃣ 태그 사전 관리")

with st.expander("❓ 태그 사전이 뭔가요?", expanded=False):
    st.markdown("""
계정과목마다 **"이런 표현이 나오면 이 계정과목이다"** 라는 예시를 모아둔 목록입니다.
부서 계획서에는 계정과목 칸이 없어서, 앱이 이 예시와 사유를 비교해 계정과목을 추천합니다.

| 언제 쓰나요 | 이렇게 하세요 |
|---|---|
| 추천이 자꾸 엉뚱하게 나올 때 | 그 계정과목에 실제 쓰는 표현을 예시로 추가 |
| 새 표현이 생겼을 때 | 미리 등록해두면 다음 달부터 자동 분류 |
| 검색이 안 될 때 | 예시에 넣어두면 계획 화면 드롭다운에서 검색됨 |

**예시는 자동으로도 쌓입니다.** 계획 업로드 화면에서 계정과목을 직접 고르면 그 사유가 자동 저장됩니다.
여기서는 미리 채워두거나, 잘못 쌓인 걸 지울 때 쓰시면 됩니다.

⚠️ **계정과목 이름 자체는 바꾸지 마세요.** 원본 자금운용계획 엑셀의 수식이 이 이름을 그대로 찾습니다.
""")

tags = load_tags()
usage = tag_usage_counts()

# ── 한눈에 보기 ──
st.subheader("전체 계정과목")
rows = []
for name, info in sorted(tags.items()):
    grp = ""
    for rk, (g, c) in ANNUAL_GROUP.items():
        pass
    directions = sorted({d for (a, d) in ACCOUNT_MAP if a == name})
    rows.append({
        "계정과목": name,
        "방향": " / ".join("입금" if d == "in" else "출금" for d in directions) or "—",
        "유동자금세부 행": " / ".join(
            ROW_LABEL.get(ACCOUNT_MAP[(name, d)], "") for d in directions) or "(집계 제외)",
        "예시 개수": len(info.example_list),
        "쓰인 건수": usage.get(name, 0),
        "집계 제외": "예" if info.exclude_agg else "",
        "매달 확인": "예" if info.manual_input else "",
    })
df = pd.DataFrame(rows)
only_empty = st.checkbox("예시가 없는 것만 보기")
st.dataframe(df[df["예시 개수"] == 0] if only_empty else df,
             use_container_width=True, height=380)

st.divider()

# ── 하나 골라서 편집 ──
st.subheader("예시 편집")
names = sorted(tags.keys())
sel = st.selectbox("편집할 계정과목", names,
                   format_func=lambda n: f"{n}  (예시 {len(tags[n].example_list)}개)")
info = tags[sel]

directions = sorted({d for (a, d) in ACCOUNT_MAP if a == sel})
if directions:
    where = " / ".join(
        f"{'입금' if d == 'in' else '출금'} → {ROW_LABEL.get(ACCOUNT_MAP[(sel, d)], '?')}"
        for d in directions)
    st.caption(f"집계 위치: {where}")
elif sel in EXCLUDED_ACCOUNTS:
    st.caption("집계에서 제외되는 계정과목입니다 (계좌간 대체 등).")

text = st.text_area(
    "적요 예시 — 한 줄에 하나씩",
    "\n".join(info.example_list), height=200,
    help="실제 부서 계획서·일일자금계획에 나오는 표현을 그대로 적어주세요. "
         "예) 대표번호 전화요금 / 인터넷 회선료",
)

c1, c2 = st.columns(2)
manual = c1.checkbox("매달 재확인 필요 항목 (취합 파일에 안 나타남)",
                     value=bool(info.manual_input))
exclude = c2.checkbox("집계에서 제외", value=bool(info.exclude_agg))

if st.button("💾 저장", type="primary"):
    set_tag_examples(sel, text.splitlines())
    set_tag_flags(sel, int(manual), int(exclude))
    st.success(f"'{sel}' 예시 {len([x for x in text.splitlines() if x.strip()])}개를 저장했습니다. "
               "다음 분류부터 바로 반영됩니다.")
    st.rerun()

st.divider()
with st.expander("🧪 여기서 바로 시험해보기"):
    from core.data_store import Txn
    from core.tag_classifier import suggest_tag
    q = st.text_input("사유를 입력해보세요", placeholder="예: 대표번호 전화요금")
    if q:
        sug = suggest_tag(Txn("", "plan", memo=q, amount_out=1))
        mark = {"auto": "🟢 자동", "suggest": "🟡 추천", "manual": "🔴 못 찾음"}[sug.level]
        st.write(f"{mark} — **{sug.tag or '(없음)'}** / {sug.score}점 / {sug.reason}")
