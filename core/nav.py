"""사이드바 메뉴 감추기 — '보조 메뉴'는 평소에 숨겨 둡니다.

왜
    매달 쓰는 화면은 1~4번이면 충분합니다. 5·6·7번은 가끔만 쓰는데
    사이드바에 늘 나와 있으면 '어디서 뭘 하지'가 헷갈립니다.

어떻게
    스트림릿은 `app/pages/` 안의 파일을 **무조건** 사이드바에 넣습니다.
    그래서 목록에서 빼는 대신, 해당 링크만 CSS 로 감춥니다.
    (링크 주소로 찾으므로 파일 이름이 그대로면 안전합니다)

켜고 끄기
    시작 화면(Home)의 [보조 메뉴 보기] 스위치. 껐다 켜도 기억합니다.
"""
from __future__ import annotations

from core.uploads import load_state, save_state

# URL 조각(파일 이름에서 앞의 숫자를 뗀 것) → 화면 이름
OPTIONAL_PAGES = {
    "Remark_Labels": "비고 라벨 관리",
    "Tag_Dictionary": "태그 사전 관리",
}

_KEY = "show_optional"


def show_optional() -> bool:
    return bool(load_state(_KEY, False))


def set_show_optional(on: bool) -> None:
    save_state(_KEY, bool(on))


def apply_nav(st) -> None:
    """보조 메뉴를 숨깁니다. 각 화면 맨 위에서 한 번 불러주세요."""
    if show_optional():
        return
    sel = ", ".join(
        f'[data-testid="stSidebarNav"] a[href$="/{slug}"]' for slug in OPTIONAL_PAGES)
    st.markdown(f"<style>{sel} {{display:none;}}</style>", unsafe_allow_html=True)
