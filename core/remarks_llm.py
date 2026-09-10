"""방식 B(생성형 AI로 비고 문장 작성) — 자리만 비워둔 모듈.

회사 내부망에서 외부 API를 못 쓸 수 있어 지금은 구현하지 않습니다.
나중에 필요하면 여기에 API 호출을 넣고 remarks_rule 대신 부르면 됩니다.
"""
from __future__ import annotations


def generate_leaf_remarks(*args, **kwargs):
    raise NotImplementedError("방식 B(LLM)는 아직 구현하지 않았습니다. remarks_rule을 사용하세요.")
