"""경로·설정 한 곳 모음.

DB 위치를 회사 Google Drive 동기화 폴더로 바꾸려면 환경변수 FUNDPLAN_DB 를 쓰거나
아래 DB_PATH 를 직접 고치세요.
  예) Windows:  set FUNDPLAN_DB=G:\\내 드라이브\\자금운용\\tags.db
"""
from __future__ import annotations

import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

DB_PATH = Path(os.environ.get("FUNDPLAN_DB", ROOT / "db" / "tags.db"))
SCHEMA_PATH = ROOT / "db" / "schema.sql"

TEMPLATE_DIR = ROOT / "templates"

# templates 폴더에서 '내용물'로만 걸러낼 파일들 (템플릿 후보가 아님)
_NOT_TEMPLATE = ("서식기준.xlsm",)


def find_template(preferred: Path, suffixes: tuple[str, ...]) -> Path:
    """템플릿 파일을 **이름과 상관없이** 찾습니다.

    ① 정해진 이름의 파일이 있으면 그걸 씁니다.
    ② 없으면 templates 폴더에서 확장자가 맞는 파일 중 **가장 최근 것**을 씁니다.
       (`_fast_…` 는 앱이 만든 속도용 사본, `~$…` 는 엑셀 임시파일이라 제외)
    """
    if preferred.exists():
        return preferred
    folder = preferred.parent
    if folder.exists():
        cands = [p for p in folder.iterdir()
                 if p.is_file() and p.suffix.lower() in suffixes
                 and not p.name.startswith(("_fast_", "~$", "."))
                 and p.name not in _NOT_TEMPLATE]
        if cands:
            return max(cands, key=lambda p: p.stat().st_mtime)
    return preferred


def _env_or_find(env_key: str, name: str, suffixes: tuple[str, ...]) -> Path:
    v = os.environ.get(env_key, "").strip()
    return Path(v) if v else find_template(TEMPLATE_DIR / name, suffixes)


# 원본 워크북 템플릿 (사용자가 직접 넣어야 하는 파일 — .gitignore 대상)
TEMPLATE_PATH = _env_or_find("FUNDPLAN_TEMPLATE",
                             "자금운용계획_template.xlsm", (".xlsm", ".xltm"))

# 보고 PPT 템플릿 (지난달 보고서를 그대로 씁니다)
PPT_TEMPLATE_PATH = _env_or_find("FUNDPLAN_PPT_TEMPLATE",
                                 "자금운용계획_template.pptx", (".pptx", ".potx"))

OUTPUT_DIR = ROOT / "data" / "output"

# 한국은행 ECOS API 키 (환율 자동 조회용). 환경변수로 넣는 것을 권장합니다.
ECOS_API_KEY = os.environ.get("ECOS_API_KEY", "")

# 금액 표시 단위: 원본 워크북은 백만원 단위
MILLION = 1_000_000
