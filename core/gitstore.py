"""깃허브를 **영구 저장소**로 쓰기.

왜 필요한가
    Streamlit Cloud 는 앱을 다시 띄울 때마다 디스크를 **깃허브 내용으로 되돌립니다.**
    그래서 Cloud 안에 저장한 거래·비고·잔고는 한 달을 못 버티고 사라집니다.
    (2026-10 산출물에서 연간 시트 8·9월이 비어 버린 것이 이 때문입니다)

어떻게
    DB 파일(`db/tags.db`) 하나를 **데이터 전용 브랜치**에 올려 둡니다.
        앱 시작 → 깃허브에서 내려받기
        저장     → 깃허브에 올리기 (약 1초 뒤 한 번에 모아서)
    코드가 든 브랜치(main)는 건드리지 않으므로 앱이 다시 배포되지 않습니다.

설정 (Streamlit  Settings → Secrets 에 붙여넣기)
    FUNDPLAN_GH_TOKEN  = "github_pat_..."   ← 필수
    FUNDPLAN_GH_REPO   = "사용자/리포이름"  ← 생략하면 아래 기본값
    FUNDPLAN_GH_BRANCH = "data"             ← 생략 가능
    FUNDPLAN_GH_PATH   = "store/tags.db"   ← 생략 가능

설정이 없으면
    아무 일도 하지 않습니다. 내 PC에서 쓸 때처럼 그냥 로컬 파일만 씁니다.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

from core.config import DB_PATH, ROOT

API = "https://api.github.com"
DEFAULT_REPO = "dlawjdgk0715/fund-plan_test"
DEFAULT_BRANCH = "data"
DEFAULT_PATH = "store/tags.db"

_STATE = ROOT / "data" / "_state" / "gitstore.json"
_PUSH_DELAY = 1.2          # 연달아 저장해도 이만큼 모았다가 한 번에 올립니다
_TIMEOUT = 25

_lock = threading.Lock()
_timer: threading.Timer | None = None
_last: dict = {"ok": None, "msg": "", "at": "", "sha": "", "pulled": False}


# ── 설정 읽기 ─────────────────────────────────────────────────────
def _secret(key: str, default: str = "") -> str:
    v = os.environ.get(key, "").strip()
    if v:
        return v
    try:                                    # Streamlit Cloud 의 Secrets
        import streamlit as st
        return str(st.secrets.get(key, "") or "").strip() or default
    except Exception:
        return default


def config() -> dict:
    return {
        "token": _secret("FUNDPLAN_GH_TOKEN"),
        "repo": _secret("FUNDPLAN_GH_REPO", DEFAULT_REPO),
        "branch": _secret("FUNDPLAN_GH_BRANCH", DEFAULT_BRANCH),
        "path": _secret("FUNDPLAN_GH_PATH", DEFAULT_PATH),
    }


def enabled() -> bool:
    return bool(config()["token"])


# ── 지난 기록 ─────────────────────────────────────────────────────
def _load_state() -> dict:
    try:
        return json.loads(_STATE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_state(d: dict) -> None:
    try:
        _STATE.parent.mkdir(parents=True, exist_ok=True)
        _STATE.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _digest(p: Path) -> str:
    try:
        return hashlib.sha1(p.read_bytes()).hexdigest()
    except OSError:
        return ""


# ── 깃허브 호출 ───────────────────────────────────────────────────
def _req(method: str, url: str, token: str, body: dict | None = None):
    data = json.dumps(body).encode() if body is not None else None
    r = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "Content-Type": "application/json",
        "User-Agent": "fund-plan-app",
    })
    with urllib.request.urlopen(r, timeout=_TIMEOUT) as resp:
        raw = resp.read()
    return json.loads(raw) if raw else {}


def _ensure_branch(cfg: dict) -> None:
    """데이터 브랜치가 없으면 기본 브랜치에서 하나 만듭니다."""
    repo, br, tok = cfg["repo"], cfg["branch"], cfg["token"]
    try:
        _req("GET", f"{API}/repos/{repo}/git/ref/heads/{br}", tok)
        return
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    info = _req("GET", f"{API}/repos/{repo}", tok)
    head = _req("GET", f"{API}/repos/{repo}/git/ref/heads/{info['default_branch']}", tok)
    _req("POST", f"{API}/repos/{repo}/git/refs", tok,
         {"ref": f"refs/heads/{br}", "sha": head["object"]["sha"]})


def _remote(cfg: dict) -> dict | None:
    """원격 파일 정보 (없으면 None)."""
    url = f"{API}/repos/{cfg['repo']}/contents/{cfg['path']}?ref={cfg['branch']}"
    try:
        return _req("GET", url, cfg["token"])
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


# ── 내려받기 ──────────────────────────────────────────────────────
def pull(force: bool = False) -> dict:
    """깃허브에 있는 DB 를 내려받습니다. 앱이 시작할 때 **한 번** 부릅니다."""
    cfg = config()
    if not cfg["token"]:
        return {"ok": None, "msg": "저장소 설정 없음 (로컬 파일만 사용)"}
    st = _load_state()
    if st.get("pulled_session") == _boot_id() and not force:
        return {"ok": True, "msg": "이미 내려받음"}
    try:
        _ensure_branch(cfg)
        info = _remote(cfg)
        if not info:
            _mark(True, "저장소에 아직 DB 가 없습니다 — 처음 저장할 때 만들어집니다")
            st["pulled_session"] = _boot_id()
            _save_state(st)
            return {"ok": True, "msg": "원격 DB 없음"}
        blob = base64.b64decode(info.get("content", "") or "")
        if not blob and info.get("sha"):          # 1MB 넘으면 content 가 비어 옵니다
            g = _req("GET", f"{API}/repos/{cfg['repo']}/git/blobs/{info['sha']}",
                     cfg["token"])
            blob = base64.b64decode(g.get("content", "") or "")
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        DB_PATH.write_bytes(blob)
        st.update({"sha": info["sha"], "digest": _digest(DB_PATH),
                   "pulled_at": _now(), "pulled_session": _boot_id()})
        _save_state(st)
        _mark(True, f"저장소에서 불러왔습니다 ({len(blob):,} 바이트)", sha=info["sha"],
              pulled=True)
        return {"ok": True, "msg": "내려받기 완료"}
    except Exception as e:                        # noqa: BLE001
        _mark(False, f"불러오기 실패: {e}")
        return {"ok": False, "msg": str(e)}


# ── 올리기 ────────────────────────────────────────────────────────
def push(reason: str = "") -> dict:
    """DB 를 깃허브에 올립니다. 내용이 그대로면 아무 것도 하지 않습니다."""
    cfg = config()
    if not cfg["token"]:
        return {"ok": None, "msg": "저장소 설정 없음"}
    if not DB_PATH.exists():
        return {"ok": None, "msg": "DB 파일 없음"}
    with _lock:
        st = _load_state()
        dig = _digest(DB_PATH)
        if dig and dig == st.get("digest") and st.get("sha"):
            return {"ok": True, "msg": "바뀐 내용 없음"}
        try:
            _ensure_branch(cfg)
            info = _remote(cfg)
            body = {
                "message": f"자료 저장 {_now()}" + (f" — {reason}" if reason else ""),
                "content": base64.b64encode(DB_PATH.read_bytes()).decode(),
                "branch": cfg["branch"],
            }
            if info:
                body["sha"] = info["sha"]
            res = _req("PUT",
                       f"{API}/repos/{cfg['repo']}/contents/{cfg['path']}",
                       cfg["token"], body)
            sha = res.get("content", {}).get("sha", "")
            st.update({"sha": sha, "digest": dig, "pushed_at": _now()})
            _save_state(st)
            _mark(True, "저장소에 올렸습니다", sha=sha)
            return {"ok": True, "msg": "올리기 완료"}
        except urllib.error.HTTPError as e:
            if e.code == 409:
                _mark(False, "다른 분이 먼저 저장했습니다 — 화면을 새로고침한 뒤 "
                             "다시 저장해주세요")
                return {"ok": False, "msg": "충돌"}
            _mark(False, f"올리기 실패: HTTP {e.code}")
            return {"ok": False, "msg": f"HTTP {e.code}"}
        except Exception as e:                    # noqa: BLE001
            _mark(False, f"올리기 실패: {e}")
            return {"ok": False, "msg": str(e)}


# ── 저장할 때마다 부르는 곳 ───────────────────────────────────────
def mark_dirty(reason: str = "") -> None:
    """DB 에 뭔가 썼다고 알립니다. 잠시 모았다가 한 번에 올립니다.

    연달아 여러 번 저장해도(예: 부서 7개 계획 저장) 올리기는 **한 번**입니다.
    """
    global _timer
    if not enabled():
        return
    with _lock:
        if _timer is not None:
            _timer.cancel()
        _timer = threading.Timer(_PUSH_DELAY, lambda: push(reason))
        _timer.daemon = True
        _timer.start()


def flush(reason: str = "") -> dict:
    """기다리지 않고 지금 올립니다 (수동 [저장] 버튼용)."""
    global _timer
    with _lock:
        if _timer is not None:
            _timer.cancel()
            _timer = None
    return push(reason)


# ── 화면에 보여줄 상태 ────────────────────────────────────────────
def status() -> dict:
    st = _load_state()
    cfg = config()
    dirty = bool(DB_PATH.exists() and _digest(DB_PATH) != st.get("digest"))
    return {
        "enabled": bool(cfg["token"]),
        "repo": cfg["repo"], "branch": cfg["branch"], "path": cfg["path"],
        "pulled_at": st.get("pulled_at", ""),
        "pushed_at": st.get("pushed_at", ""),
        "dirty": dirty,
        "last_ok": _last["ok"], "last_msg": _last["msg"], "last_at": _last["at"],
    }


# ── 잡다한 것 ─────────────────────────────────────────────────────
_BOOT = time.strftime("%Y%m%d%H%M%S")


def _boot_id() -> str:
    return _BOOT


def _now() -> str:
    from core.period import now_kst
    return now_kst()


def _mark(ok, msg, sha="", pulled=False) -> None:
    _last.update({"ok": ok, "msg": msg, "at": _now()})
    if sha:
        _last["sha"] = sha
    if pulled:
        _last["pulled"] = True
