"""올린 파일을 **화면을 옮겨도 잃지 않도록** 보관합니다.

왜 필요한가
    스트림릿은 화면(페이지)을 옮기면 그 화면의 위젯 상태를 지웁니다.
    `st.file_uploader` 도 예외가 아니라서, 1번에서 파일을 올리고 2번에 갔다가
    돌아오면 "파일을 올려주세요" 로 되돌아갑니다. **오류가 아니라 원래 동작**입니다.

어떻게 해결하나
    올린 파일을 `data/_uploads/` 에 **내용 지문(sha1)을 이름으로** 저장해 두고,
    그 목록만 세션에 기억합니다. 세션 값은 화면을 옮겨도 남습니다.

왜 지문을 파일 이름에 쓰나
    예전에는 늘 같은 이름(`_upload_actual.xlsx`)에 덮어썼습니다. 그러면
    "A 올림 → B 올림 → 다시 A 올림" 순서에서, 마지막 A 는 캐시에 걸려 저장을
    건너뛰는데 디스크에는 B 가 남아 **엉뚱한 파일을 읽는** 문제가 생깁니다.
    지문을 이름에 넣으면 파일 이름과 내용이 항상 짝이 맞습니다.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

from core.config import ROOT

KEEP_DIR = ROOT / "data" / "_uploads"
KEEP_PER_PREFIX = 8          # 종류별로 최근 몇 개까지 남길지


def _rel(p) -> str:
    """기록에는 **프로젝트 폴더 기준 상대경로**로 적습니다.

    절대경로로 적으면 폴더를 옮기거나 다른 PC에서 열었을 때 못 찾습니다.
    """
    try:
        return Path(p).resolve().relative_to(ROOT).as_posix()
    except Exception:
        return str(p)


def _abs(s) -> Path:
    p = Path(str(s))
    return p if p.is_absolute() else (ROOT / p)


def stamp_of(data: bytes) -> str:
    """파일 내용 지문. 내용이 한 글자라도 바뀌면 값이 달라집니다."""
    return hashlib.sha1(data).hexdigest()[:16]


@dataclass
class KeptFile:
    """보관해 둔 업로드 파일. `st.file_uploader` 결과와 같은 방식으로 씁니다."""
    name: str          # 사용자가 올린 원래 파일 이름
    path: str          # 보관 위치
    stamp: str         # 내용 지문 (캐시 열쇠로 씁니다)

    def getbuffer(self) -> memoryview:
        return memoryview(_abs(self.path).read_bytes())

    def to_dict(self) -> dict:
        return {"name": self.name, "path": _rel(self.path), "stamp": self.stamp}

    @staticmethod
    def from_dict(d: dict | None) -> "KeptFile | None":
        """기록을 되살립니다. 파일이 지워졌으면 None."""
        if not d:
            return None
        p = _abs(d.get("path", ""))
        if not p.exists():
            return None
        return KeptFile(d.get("name", p.name), str(p), d.get("stamp", ""))

    @staticmethod
    def list_from(records) -> list["KeptFile"]:
        return [k for k in (KeptFile.from_dict(d) for d in (records or [])) if k]


def keep(prefix: str, name: str, data: bytes) -> KeptFile:
    """업로드 파일을 보관하고 KeptFile 로 돌려줍니다 (같은 내용이면 다시 안 씁니다)."""
    KEEP_DIR.mkdir(parents=True, exist_ok=True)
    st = stamp_of(data)
    suffix = Path(name).suffix.lower() or ".xlsx"
    path = KEEP_DIR / f"{prefix}_{st}{suffix}"
    if not path.exists() or path.stat().st_size != len(data):
        path.write_bytes(data)
    prune(prefix)
    return KeptFile(name, str(path), st)


def prune(prefix: str, keep_n: int = KEEP_PER_PREFIX) -> None:
    """오래된 보관 파일을 지웁니다 (디스크가 계속 불어나지 않도록)."""
    if not KEEP_DIR.exists():
        return
    files = sorted((p for p in KEEP_DIR.glob(f"{prefix}_*") if p.is_file()),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    for p in files[keep_n:]:
        try:
            p.unlink()
        except OSError:
            pass


# ── 앱을 껐다 켜도 기억하기 ────────────────────────────────────────
# 세션(st.session_state)은 cmd 창을 닫으면 사라집니다. 그래서 마지막으로 올린
# 파일이 무엇이었는지를 작은 기록 파일에 남겨 둡니다. 거래 내역 자체는 DB에
# 남아 있으므로, 이 기록은 '다시 안 올려도 되게' 하기 위한 편의 장치입니다.
STATE_DIR = ROOT / "data" / "_state"


def _state_path(kind: str) -> Path:
    return STATE_DIR / f"{kind}.json"


def save_state(kind: str, obj) -> None:
    """작은 기록을 디스크에 남깁니다 (앱을 껐다 켜도 남도록)."""
    import json
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        _state_path(kind).write_text(
            json.dumps(obj, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def load_state(kind: str, default=None):
    import json
    p = _state_path(kind)
    if not p.exists():
        return default
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return default


def remember(kind: str, records: list[dict]) -> None:
    """마지막으로 올린 파일 목록을 디스크에 남깁니다."""
    save_state(kind, records)


def recall(kind: str) -> list[dict]:
    """지난번에 올린 파일 목록 (파일이 아직 있는 것만)."""
    recs = load_state(kind, []) or []
    return [r for r in recs if isinstance(r, dict) and _abs(r.get("path", "")).exists()]


def forget(kind: str) -> None:
    try:
        _state_path(kind).unlink(missing_ok=True)
    except OSError:
        pass
