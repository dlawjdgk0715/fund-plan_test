import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

import pandas as pd
import streamlit as st

from core.nav import apply_nav
from core.data_store import init_db, load_remark_labels, load_transactions, save_remark_label
from core.remark_labels import get_label, seed_defaults

st.set_page_config(page_title="6. 비고 라벨 관리", layout="wide")
apply_nav(st)
init_db()
st.title("6️⃣ 비고 라벨 관리")
st.caption("비고(N열)에 적요를 그대로 쓰면 너무 길어서, 짧은 라벨로 바꿔 씁니다. "
           "여기서 고치면 **바로 다음 재계산·엑셀 만들기부터** 반영됩니다.")

if st.button("기본 예시 채우기"):
    seed_defaults()
    st.success("기본 라벨을 채웠습니다.")

period = st.session_state.get("period", "")
rows = []
seen = set()
for kind in ("actual", "plan"):
    for t in load_transactions(period, kind) if period else []:
        key = (t.tag, t.memo)
        if key in seen or not t.memo:
            continue
        seen.add(key)
        rows.append({"계정과목": t.tag, "적요": t.memo, "라벨": get_label(t.tag, t.memo)})

for (tag, memo), label in load_remark_labels().items():
    if (tag, memo) not in seen:
        rows.append({"계정과목": tag, "적요": memo, "라벨": label})

if not rows:
    st.info("아직 저장된 거래가 없습니다. 파일을 먼저 올려주세요.")
    st.stop()

edited = st.data_editor(pd.DataFrame(rows), use_container_width=True, height=600,
                        column_config={"계정과목": st.column_config.TextColumn(disabled=True),
                                       "적요": st.column_config.TextColumn(disabled=True)})

if st.button("라벨 저장", type="primary"):
    for _, r in edited.iterrows():
        save_remark_label(str(r["계정과목"]), str(r["적요"]), str(r["라벨"]))
    st.success("저장했습니다. 다음 재계산부터 바로 반영됩니다.")
