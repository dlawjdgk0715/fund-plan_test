"""산출물(엑셀·PPT) 만들기.

왜 따로 뺐나
    예전에는 산출 화면 안에만 있어서, 비고 하나 고치려면 화면을 옮겨야 했습니다.
    이제 4번 화면과 시작 화면이 같은 코드를 씁니다.

다시 안 만들기 (재사용)
    산출물에 영향을 주는 것들(거래·비고·잔고·환율·템플릿)을 하나의 **지문**으로
    묶어 둡니다. 지문이 그대로면 지난번 파일을 그대로 돌려줍니다(즉시).
    바뀌었으면 다시 만듭니다.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from pathlib import Path

from core.config import OUTPUT_DIR, ROOT, TEMPLATE_PATH
from core.data_store import (load_remark_style, load_report_remarks,
                             load_report_remarks_raw, load_transactions,
                             save_fund_balance, txn_fingerprint)
from core.engine import ROW_NO, aggregate, compute_fund_balance_change
from core.excel_writer import STYLE_REF_PATH, build_workbook
from core.ppt_writer import PPT_TEMPLATE, build_ppt
from core.remarks_rule import (SUMMARY1_KEYS, cashflow_rows_from_summary1,
                               detail_drafts, summary1_drafts,
                               summary1_from_cashflow)
from core.report import DETAIL_ROWS, build_report

XL_MIME = "application/vnd.ms-excel.sheet.macroEnabled.12"
PPT_MIME = ("application/vnd.openxmlformats-officedocument."
            "presentationml.presentation")

TABLE_KEYS = ("cash_flow", "liquid_plan", "fund_plan", "stock", "detail", "summary1")


@dataclass
class Outputs:
    files: list[tuple[str, Path, str]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    fund: dict = field(default_factory=dict)
    stamp: str = ""
    reused: bool = False


# ── 무엇이 바뀌면 다시 만들어야 하는가 ────────────────────────────
def _file_mark(p: Path) -> str:
    try:
        s = p.stat()
        return f"{p.name}:{s.st_size}:{int(s.st_mtime)}"
    except OSError:
        return f"{p.name}:none"


def inputs_stamp(period: str, vals: dict, actual_month: int, plan_month: int) -> str:
    """산출물에 영향을 주는 모든 것의 지문."""
    parts = [
        f"m:{actual_month}/{plan_month}",
        "v:" + json.dumps({k: round(float(v or 0), 6) for k, v in sorted(vals.items())}),
        "a:" + txn_fingerprint(load_transactions(period, "actual")),
        "p:" + txn_fingerprint(load_transactions(period, "plan")),
    ]
    for key in TABLE_KEYS:
        raw = load_report_remarks_raw(period, key, 40)
        parts.append(f"r:{key}:" + json.dumps(raw, ensure_ascii=False))
    for key in ("summary1", "detail", "cash_flow"):
        parts.append(f"s:{key}:" + json.dumps(load_remark_style(key), ensure_ascii=False,
                                              sort_keys=True))
    for p in (TEMPLATE_PATH, PPT_TEMPLATE, STYLE_REF_PATH):
        parts.append("t:" + _file_mark(p))
    return hashlib.sha1("\n".join(parts).encode("utf-8")).hexdigest()[:16]


# ── 지난번 결과 기억 ──────────────────────────────────────────────
_CACHE = ROOT / "data" / "_state" / "outputs.json"


def _load_cache(period: str, stamp: str) -> list[tuple[str, Path, str]] | None:
    try:
        d = json.loads(_CACHE.read_text(encoding="utf-8"))
    except Exception:
        return None
    if d.get("period") != period or d.get("stamp") != stamp:
        return None
    files = []
    for label, rel, mime in d.get("files", []):
        p = Path(rel)
        p = p if p.is_absolute() else (ROOT / p)
        if not p.exists():
            return None
        files.append((label, p, mime))
    return files or None


def _save_cache(period: str, stamp: str, files, fund: dict, warnings) -> None:
    import time
    _CACHE.parent.mkdir(parents=True, exist_ok=True)

    def rel(p: Path) -> str:
        try:
            return p.resolve().relative_to(ROOT).as_posix()
        except Exception:
            return str(p)
    try:
        _CACHE.write_text(json.dumps({
            "period": period, "stamp": stamp,
            "files": [[n, rel(p), m] for n, p, m in files],
            "fund": fund, "warnings": warnings,
            "made_at": time.strftime("%Y-%m-%d %H:%M"),
        }, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def last_outputs(period: str | None = None) -> dict | None:
    """마지막으로 만든 산출물 기록 (지문이 맞든 안 맞든).

    시작 화면에서 '이미 만들어 둔 파일 바로 받기' 에 씁니다.
    반환: {"period", "stamp", "made_at", "files": [(라벨, Path, mime), …]}
    """
    try:
        d = json.loads(_CACHE.read_text(encoding="utf-8"))
    except Exception:
        return None
    if period and d.get("period") != period:
        return None
    files = []
    for label, rel_, mime in d.get("files", []):
        p = Path(rel_)
        p = p if p.is_absolute() else (ROOT / p)
        if p.exists():
            files.append((label, p, mime))
    if not files:
        return None
    return {"period": d.get("period", ""), "stamp": d.get("stamp", ""),
            "made_at": d.get("made_at", ""), "files": files}


def cached_files(period: str, vals: dict, actual_month: int, plan_month: int):
    """지금 상태로 이미 만들어 둔 파일이 있으면 (파일들, 지문), 없으면 (None, 지문)."""
    stamp = inputs_stamp(period, vals, actual_month, plan_month)
    return _load_cache(period, stamp), stamp


# ── 비고 모으기 (4번 화면에 보이던 값 그대로) ─────────────────────
def collect_remarks(period: str, actual_txns, plan_txns, res, plan_res,
                    actual_month: int, plan_month: int, rep):
    """(요약1 문장들, 세부내역 36줄, 유동자금세부 N열용 dict)"""
    s1_raw = load_report_remarks_raw(period, "summary1", len(SUMMARY1_KEYS))
    s1_draft = summary1_drafts(actual_txns, plan_txns, actual_month, plan_month, rep=rep)
    legacy = summary1_from_cashflow(load_report_remarks(period, "cash_flow", 6))
    summary = {k: (s1_raw[i] if s1_raw[i] is not None
                   else (legacy.get(k, "") or s1_draft[k]))
               for i, k in enumerate(SUMMARY1_KEYS)}

    d_draft = detail_drafts(res, plan_res, actual_month, plan_month)
    d_raw = load_report_remarks_raw(period, "detail", len(DETAIL_ROWS))
    detail = [d_draft[i] if d_raw[i] is None else d_raw[i] for i in range(len(DETAIL_ROWS))]

    leaf = {}
    for (_d, _lab, key), txt in zip(DETAIL_ROWS, detail):
        if (txt or "").strip() and key in ROW_NO:
            leaf[key] = txt
    return summary, detail, leaf


# ── 만들기 ────────────────────────────────────────────────────────
def build_outputs(period: str, vals: dict, actual_month: int, plan_month: int,
                  want_excel: bool = True, want_ppt: bool = True,
                  force: bool = False) -> Outputs:
    """엑셀·PPT를 만듭니다. 바뀐 게 없으면 지난번 파일을 그대로 돌려줍니다."""
    stamp = inputs_stamp(period, vals, actual_month, plan_month)
    if not force:
        hit = _load_cache(period, stamp) or []
        got = {n for n, _p, _m in hit}
        # 필요한 것이 전부 이미 있을 때만 재사용합니다
        if hit and (not want_excel or "엑셀" in got) and (not want_ppt or "PPT" in got):
            keep = [f for f in hit
                    if (f[0] == "엑셀" and want_excel) or (f[0] == "PPT" and want_ppt)]
            try:
                d = json.loads(_CACHE.read_text(encoding="utf-8"))
            except Exception:
                d = {}
            return Outputs(files=keep, warnings=list(d.get("warnings", [])),
                           fund=dict(d.get("fund", {})), stamp=stamp, reused=True)

    actual = load_transactions(period, "actual")
    plan = load_transactions(period, "plan")
    fx = float(vals.get("fx") or 0.0)
    res = aggregate(actual, opening_krw=float(vals.get("open_krw") or 0.0),
                    opening_usd=float(vals.get("open_usd") or 0.0), fx_rate=fx or None)
    plan_res = aggregate(plan, opening_krw=0.0, fx_rate=fx or None)

    rep = build_report(
        actual, plan, actual_month, plan_month,
        open_krw=float(vals.get("open_krw") or 0.0),
        open_usd=float(vals.get("open_usd") or 0.0),
        plan_open_krw=float(vals.get("plan_open_krw") or res.closing),
        plan_open_usd=float(vals.get("plan_open_usd") or 0.0),
        invest_open_krw=float(vals.get("invest_open_krw") or 0.0),
        invest_open_usd=float(vals.get("invest_open_usd") or 0.0), fx=fx)

    summary, detail_notes, leaf = collect_remarks(
        period, actual, plan, res, plan_res, actual_month, plan_month, rep)

    out = Outputs(stamp=stamp)
    if want_excel:
        rp = build_workbook(
            actual_txns=actual, plan_txns=plan, actual_result=res,
            actual_month=actual_month, plan_month=plan_month,
            opening_liq_krw=float(vals.get("open_krw") or 0.0),
            opening_liq_usd=float(vals.get("open_usd") or 0.0),
            opening_inv_krw=float(vals.get("invest_open_krw") or 0.0),
            opening_inv_usd=float(vals.get("invest_open_usd") or 0.0),
            fx_rate=fx, leaf_remarks=leaf, summary_drafts=summary)
        out.warnings += rp.warnings
        out.files.append(("엑셀", rp.path, XL_MIME))

    if want_ppt:
        notes = {k: load_report_remarks(period, k, len(rep.tables[k]))
                 for k in ("cash_flow", "liquid_plan", "fund_plan", "stock", "detail")}
        notes["cash_flow"] = cashflow_rows_from_summary1(summary)
        notes["detail"] = detail_notes
        pp = build_ppt(rep, remarks=notes)
        out.warnings += pp.warnings
        out.files.append(("PPT", pp.path, PPT_MIME))

    # 운용자금 기말 저장 (다음 달 기초로 이어집니다)
    change = compute_fund_balance_change(actual,
                                         float(vals.get("invest_open_krw") or 0.0),
                                         float(vals.get("invest_open_usd") or 0.0))
    save_fund_balance(period, change["end_krw"], change["end_usd"])
    out.fund = {k: float(v) for k, v in change.items()}

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    _save_cache(period, stamp, out.files, out.fund, out.warnings)
    return out
