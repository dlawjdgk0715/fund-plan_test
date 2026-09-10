"""엑셀·PPT 만들기/받기 칸.

여기만 고치면 이 칸을 쓰는 화면이 함께 바뀝니다.
"""
from __future__ import annotations

import streamlit as st

from core.outputs import build_outputs, cached_files


def render_output_box(period: str, vals: dict, actual_month: int, plan_month: int,
                      key: str = "out", compact: bool = False) -> None:
    """만들기 버튼 + 다운로드 버튼. 바뀐 게 없으면 지난 파일을 바로 받습니다."""
    want_xl = st.session_state.get(f"{key}_xl", True)
    want_pt = st.session_state.get(f"{key}_pt", True)

    c1, c2 = st.columns(2)
    want_xl = c1.checkbox("📄 엑셀 (.xlsm)", value=want_xl, key=f"{key}_xl")
    want_pt = c2.checkbox("📊 PPT (.pptx)", value=want_pt, key=f"{key}_pt")
    if not (want_xl or want_pt):
        st.warning("만들 파일을 하나 이상 골라주세요.")
        return

    hit, _stamp = cached_files(period, vals, actual_month, plan_month)
    got = {n for n, _p, _m in (hit or [])}
    ready = bool(hit) and (not want_xl or "엑셀" in got) and (not want_pt or "PPT" in got)

    made = st.session_state.get(f"{key}_made")
    if made and made.get("stamp") == _stamp:
        ready = True

    if not ready:
        st.info("내용이 바뀌었습니다 — 아래 버튼을 눌러 새로 만들어주세요.")
        if st.button("🚀 파일 만들기", type="primary", use_container_width=True,
                     key=f"{key}_build"):
            with st.spinner("만드는 중입니다…"):
                out = build_outputs(period, vals, actual_month, plan_month,
                                    want_xl, want_pt)
            st.session_state[f"{key}_made"] = {"stamp": out.stamp,
                                               "warnings": out.warnings,
                                               "fund": out.fund}
            st.rerun()
        return

    # ── 받기 ──────────────────────────────────────────────────────
    out = build_outputs(period, vals, actual_month, plan_month, want_xl, want_pt)
    st.success("✅ 최신 상태입니다 — 바로 받으세요.")
    files = [f for f in out.files if f[1].exists()]
    if not files:
        st.warning("파일을 못 찾았습니다. [새로 만들기] 를 눌러주세요.")
    else:
        cols = st.columns(len(files))
        for col, (label, path, mime) in zip(cols, files):
            col.download_button(f"⬇ {label} 받기", path.read_bytes(),
                                file_name=path.name, mime=mime,
                                use_container_width=True,
                                key=f"{key}_dl_{label}")
        st.caption(f"📁 저장 위치: `{files[0][1].parent}`")

    for w in (st.session_state.get(f"{key}_made", {}).get("warnings")
              or out.warnings or []):
        st.warning(w)

    if not compact:
        ch = out.fund or st.session_state.get(f"{key}_made", {}).get("fund") or {}
        if ch:
            st.caption(f"이번 달 운용자금 — 해지 {ch.get('release_krw', 0):,.1f}백만 / "
                       f"예치 {ch.get('deposit_krw', 0):,.1f}백만 → "
                       f"기말 {ch.get('end_krw', 0):,.1f}백만 (저장 완료)")

    if st.button("🔄 새로 만들기", key=f"{key}_force"):
        with st.spinner("다시 만드는 중입니다…"):
            out = build_outputs(period, vals, actual_month, plan_month,
                                want_xl, want_pt, force=True)
        st.session_state[f"{key}_made"] = {"stamp": out.stamp,
                                           "warnings": out.warnings,
                                           "fund": out.fund}
        st.rerun()
