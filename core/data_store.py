"""DB 저장/조회 계층. 화면(app/)은 전부 이 모듈을 통해서만 DB에 접근합니다."""
from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from core.config import DB_PATH, SCHEMA_PATH


# ── 연결 ──────────────────────────────────────────────────────────
def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_columns(conn, table: str, cols: list[tuple[str, str]]) -> None:
    """이미 쓰고 있던 DB에 새 열을 더합니다.

    `CREATE TABLE IF NOT EXISTS` 는 **이미 있는 표에는 아무 일도 하지 않습니다.**
    그래서 스키마 파일에 열을 추가해도 기존 DB에는 안 생깁니다 — 여기서 채워 넣습니다.
    """
    have = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
    for name, ddl in cols:
        if name not in have:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {name} {ddl}")


# ── 매달 손으로 챙겨야 하는 항목 (취합 파일에 안 나오는 것들) ──────
# 이번 달 거래에 아래 조건에 맞는 건이 하나도 없으면 3번 화면에 안내창을 띄웁니다.
#   pattern     : 적요(또는 거래처)에 들어 있어야 하는 낱말
#   counterpart : 거래처 조건 (빈칸이면 안 봄)
#   tag         : 계정과목 조건 (빈칸이면 안 봄 — 계정과목이 아직 안 붙었어도 잡힙니다)
RECHECK_DEFAULTS = [
    # (pattern, counterpart, tag, label, hint, def_tag, def_dir)
    ("주택자금", "", "", "주택자금대출 이자 지원",
     "적요에 '주택자금' 이 들어간 거래가 이번 달에 없습니다. "
     "지원금이 나갔다면 직접 넣어주세요.", "복리후생비", "out"),
    ("법인카드", "삼성카드", "", "삼성카드 법인카드 대금",
     "거래처가 '삼성카드' 이면서 적요에 '법인카드' 가 들어간 거래가 없습니다. "
     "카드대금 결제가 있었다면 직접 넣어주세요.", "직원비용", "out"),
]


def _seed_recheck(conn) -> None:
    """기본 재확인 규칙을 넣거나(없으면) 최신 조건으로 맞춥니다."""
    for pat, cp, tag, label, hint, dtag, ddir in RECHECK_DEFAULTS:
        rows = conn.execute("SELECT id FROM recheck_items WHERE pattern=? ORDER BY id",
                            (pat,)).fetchall()
        # 같은 낱말로 만들어져 있던 옛 규칙은 하나만 남기고 지웁니다
        for extra in rows[1:]:
            conn.execute("DELETE FROM recheck_items WHERE id=?", (extra["id"],))
        if rows:
            conn.execute(
                """UPDATE recheck_items SET counterpart=?, tag=?, label=?, hint=?,
                       def_tag=?, def_dir=?, active=1 WHERE id=?""",
                (cp, tag, label, hint, dtag, ddir, rows[0]["id"]))
        else:
            conn.execute(
                """INSERT INTO recheck_items
                   (tag, pattern, label, counterpart, hint, def_tag, def_dir)
                   VALUES (?,?,?,?,?,?,?)""",
                (tag, pat, label, cp, hint, dtag, ddir))


def init_db() -> None:
    """앱 시작 시 1회. 테이블이 없으면 만들고 태그 사전을 채웁니다."""
    conn = connect()
    conn.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))
    _ensure_columns(conn, "month_balances", [
        # 전월 마지막 시트(예: 0630-1)의 「3. 자금현황 > 3) 운용자금」 에서 읽은 기초값
        ("open_invest_krw", "REAL DEFAULT 0"),
        ("open_invest_usd", "REAL DEFAULT 0"),
    ])
    _ensure_columns(conn, "recheck_items", [
        ("counterpart", "TEXT DEFAULT ''"),
        ("hint", "TEXT DEFAULT ''"),
        ("def_tag", "TEXT DEFAULT ''"),
        ("def_dir", "TEXT DEFAULT 'out'"),
    ])
    n = conn.execute("SELECT COUNT(*) FROM account_tags").fetchone()[0]
    if n == 0:
        import sys
        sys.path.insert(0, str(SCHEMA_PATH.parent))
        from seed_tags import seed  # type: ignore
        seed(conn)
    _seed_recheck(conn)
    conn.commit()
    conn.close()


# ── 거래 ──────────────────────────────────────────────────────────
@dataclass
class Txn:
    period: str
    kind: str            # 'actual' | 'plan'
    tag: str = ""
    memo: str = ""
    counterpart: str = ""
    dept: str = ""
    txn_date: str = ""
    currency: str = "KRW"
    amount_in: float = 0.0
    amount_out: float = 0.0
    source: str = "upload"
    excluded: int = 0
    id: int | None = None

    @property
    def direction(self) -> str:
        return "in" if self.amount_in > 0 else "out"

    @property
    def amount(self) -> float:
        return self.amount_in if self.amount_in > 0 else self.amount_out


def save_transactions(txns: list[Txn]) -> int:
    conn = connect()
    cur = conn.cursor()
    for t in txns:
        cur.execute(
            """INSERT INTO transactions
               (period, kind, dept, txn_date, counterpart, memo, tag, currency,
                amount_in, amount_out, source, excluded)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (t.period, t.kind, t.dept, t.txn_date, t.counterpart, t.memo, t.tag,
             t.currency, t.amount_in, t.amount_out, t.source, t.excluded),
        )
    conn.commit()
    conn.close()
    return len(txns)


def load_transactions(period: str, kind: str | None = None,
                      include_excluded: bool = False) -> list[Txn]:
    conn = connect()
    q = "SELECT * FROM transactions WHERE period=?"
    args: list[Any] = [period]
    if kind:
        q += " AND kind=?"
        args.append(kind)
    if not include_excluded:
        q += " AND excluded=0"
    rows = conn.execute(q + " ORDER BY id", args).fetchall()
    conn.close()
    return [Txn(**{k: r[k] for k in r.keys() if k != "created_at"}) for r in rows]


def update_transaction(txn_id: int, **fields) -> None:
    if not fields:
        return
    conn = connect()
    sets = ", ".join(f"{k}=?" for k in fields)
    conn.execute(f"UPDATE transactions SET {sets} WHERE id=?",
                 [*fields.values(), txn_id])
    conn.commit()
    conn.close()


def delete_transaction(txn_id: int) -> None:
    conn = connect()
    conn.execute("DELETE FROM transactions WHERE id=?", (txn_id,))
    conn.commit()
    conn.close()


def clear_period(period: str, kind: str) -> int:
    conn = connect()
    cur = conn.execute("DELETE FROM transactions WHERE period=? AND kind=?", (period, kind))
    conn.commit()
    n = cur.rowcount
    conn.close()
    return n


def clear_period_dept(period: str, kind: str, dept: str) -> int:
    """특정 부서의 이번 달 데이터만 지웁니다 (다시 업로드할 때 중복 방지)."""
    conn = connect()
    cur = conn.execute("DELETE FROM transactions WHERE period=? AND kind=? AND dept=?",
                       (period, kind, dept))
    conn.commit()
    n = cur.rowcount
    conn.close()
    return n


def count_transactions(period: str, kind: str, dept: str | None = None) -> int:
    conn = connect()
    q = "SELECT COUNT(*) FROM transactions WHERE period=? AND kind=?"
    args = [period, kind]
    if dept is not None:
        q += " AND dept=?"
        args.append(dept)
    n = conn.execute(q, args).fetchone()[0]
    conn.close()
    return n


def dept_counts(period: str, kind: str) -> dict[str, int]:
    conn = connect()
    rows = conn.execute(
        "SELECT dept, COUNT(*) AS n FROM transactions WHERE period=? AND kind=? GROUP BY dept",
        (period, kind)).fetchall()
    conn.close()
    return {(r["dept"] or "(부서없음)"): r["n"] for r in rows}


# ── 태그 사전 ─────────────────────────────────────────────────────
@dataclass
class TagInfo:
    tag: str
    needs_confirm: int = 0
    manual_input: int = 0
    exclude_agg: int = 0
    examples: str = ""
    io_class: str = ""

    @property
    def example_list(self) -> list[str]:
        return [e.strip() for e in self.examples.splitlines() if e.strip()]


def load_tags() -> dict[str, TagInfo]:
    conn = connect()
    rows = conn.execute("SELECT * FROM account_tags ORDER BY tag").fetchall()
    conn.close()
    return {r["tag"]: TagInfo(**dict(r)) for r in rows}


def add_tag_example(tag: str, memo: str) -> None:
    """사용자가 확인한 적요를 사전에 누적 → 다음 달 자동분류율 상승."""
    if not memo.strip():
        return
    conn = connect()
    row = conn.execute("SELECT examples FROM account_tags WHERE tag=?", (tag,)).fetchone()
    if row is None:
        conn.close()
        return
    cur_examples = [e.strip() for e in (row["examples"] or "").splitlines() if e.strip()]
    if memo.strip() not in cur_examples:
        cur_examples.append(memo.strip())
        conn.execute("UPDATE account_tags SET examples=? WHERE tag=?",
                     ("\n".join(cur_examples), tag))
        conn.commit()
    conn.close()


def set_tag_examples(tag: str, examples: list[str]) -> None:
    """태그 사전 관리 화면(7번)에서 적요 예시 전체를 새로 저장합니다."""
    clean = []
    for e in examples:
        e = e.strip()
        if e and e not in clean:
            clean.append(e)
    conn = connect()
    conn.execute("UPDATE account_tags SET examples=? WHERE tag=?", ("\n".join(clean), tag))
    conn.commit()
    conn.close()


def set_tag_flags(tag: str, manual_input: int, exclude_agg: int) -> None:
    conn = connect()
    conn.execute("UPDATE account_tags SET manual_input=?, exclude_agg=? WHERE tag=?",
                 (manual_input, exclude_agg, tag))
    conn.commit()
    conn.close()


def tag_usage_counts() -> dict[str, int]:
    """계정과목별로 지금까지 저장된 거래 건수 — 사전 화면에서 참고용."""
    conn = connect()
    rows = conn.execute(
        "SELECT tag, COUNT(*) AS n FROM transactions WHERE excluded=0 GROUP BY tag"
    ).fetchall()
    conn.close()
    return {r["tag"]: r["n"] for r in rows if r["tag"]}


@dataclass
class RowMapEntry:
    tag: str
    direction: str
    currency: str
    row_key: str
    annual_group: str = ""
    annual_class: str = ""
    sort_order: int = 0


def load_row_map() -> list[RowMapEntry]:
    conn = connect()
    rows = conn.execute("SELECT * FROM tag_row_map").fetchall()
    conn.close()
    return [RowMapEntry(**dict(r)) for r in rows]


def find_row_map(tag: str, direction: str, currency: str) -> RowMapEntry | None:
    conn = connect()
    r = conn.execute(
        "SELECT * FROM tag_row_map WHERE tag=? AND direction=? AND currency IN (?, '*')"
        " ORDER BY CASE currency WHEN '*' THEN 1 ELSE 0 END LIMIT 1",
        (tag, direction, currency),
    ).fetchone()
    conn.close()
    return RowMapEntry(**dict(r)) if r else None


# ── 열 매핑 ───────────────────────────────────────────────────────
def save_column_mapping(dept: str, mapping: dict, signature: str) -> None:
    conn = connect()
    conn.execute(
        """INSERT INTO column_mappings (dept, mapping_json, signature)
           VALUES (?,?,?)
           ON CONFLICT(dept) DO UPDATE SET mapping_json=excluded.mapping_json,
             signature=excluded.signature, updated_at=datetime('now','localtime')""",
        (dept, json.dumps(mapping, ensure_ascii=False), signature),
    )
    conn.commit()
    conn.close()


def load_column_mapping(dept: str) -> tuple[dict, str] | None:
    conn = connect()
    r = conn.execute("SELECT * FROM column_mappings WHERE dept=?", (dept,)).fetchone()
    conn.close()
    if not r:
        return None
    return json.loads(r["mapping_json"]), r["signature"]


# ── 비고 라벨 ─────────────────────────────────────────────────────
def load_remark_labels() -> dict[tuple[str, str], str]:
    conn = connect()
    rows = conn.execute("SELECT * FROM remark_labels").fetchall()
    conn.close()
    return {(r["tag"], r["memo"]): r["label"] for r in rows}


def save_remark_label(tag: str, memo: str, label: str) -> None:
    conn = connect()
    conn.execute(
        """INSERT INTO remark_labels (tag, memo, label) VALUES (?,?,?)
           ON CONFLICT(tag, memo) DO UPDATE SET label=excluded.label""",
        (tag, memo, label),
    )
    conn.commit()
    conn.close()


# ── 운용자금 잔고 이력 (§11) ──────────────────────────────────────
def save_fund_balance(period: str, krw: float, usd: float, source: str = "auto") -> None:
    conn = connect()
    conn.execute(
        """INSERT INTO fund_balances (period, invest_end_krw, invest_end_usd, source)
           VALUES (?,?,?,?)
           ON CONFLICT(period) DO UPDATE SET invest_end_krw=excluded.invest_end_krw,
             invest_end_usd=excluded.invest_end_usd, source=excluded.source,
             updated_at=datetime('now','localtime')""",
        (period, krw, usd, source),
    )
    conn.commit()
    conn.close()


def load_fund_balance(period: str) -> dict | None:
    conn = connect()
    r = conn.execute("SELECT * FROM fund_balances WHERE period=?", (period,)).fetchone()
    conn.close()
    return dict(r) if r else None


def previous_period(period: str) -> str:
    y, m = (int(x) for x in period.split("-"))
    return f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"


# ── 월별 유동/운용 잔고 (일일자금계획에서 자동으로 읽은 값) ──────
def save_month_balance(period: str, **vals) -> None:
    cols = ["open_krw", "open_usd", "close_krw", "close_usd",
            "invest_krw", "invest_usd", "open_invest_krw", "open_invest_usd",
            "open_sheet", "close_sheet"]
    data = {c: vals.get(c, 0 if c.endswith(("krw", "usd")) else "") for c in cols}
    conn = connect()
    conn.execute(
        f"""INSERT INTO month_balances (period, {', '.join(cols)})
            VALUES (?, {', '.join('?' * len(cols))})
            ON CONFLICT(period) DO UPDATE SET
              {', '.join(f'{c}=excluded.{c}' for c in cols)},
              updated_at=datetime('now','localtime')""",
        [period, *[data[c] for c in cols]],
    )
    conn.commit()
    conn.close()


def load_month_balance(period: str) -> dict | None:
    conn = connect()
    r = conn.execute("SELECT * FROM month_balances WHERE period=?", (period,)).fetchone()
    conn.close()
    return dict(r) if r else None


def resolve_open_invest(period: str) -> tuple[float, float, str]:
    """이번 달 **기초 운용자금**을 어디서 가져올지 한 곳에서 정합니다.

    ① 일일자금계획 **전월 마지막 시트**(예: 0630-1)의 「3. 자금현황 > 3) 운용자금」
       — 은행 잔고 그대로라 가장 확실합니다
    ② 지난달 산출 때 앱이 저장해 둔 기말 운용자금
    ③ 지난달 일일자금계획에서 읽은 당월말 운용자금
    반환: (원화 백만원, 외화 천$, 어디서 왔는지 설명)
    """
    mb = load_month_balance(period)
    if mb and mb.get("open_invest_krw"):
        sheet = mb.get("open_sheet") or "전월 마지막"
        return (float(mb["open_invest_krw"]), float(mb.get("open_invest_usd") or 0.0),
                f"일일자금계획 `{sheet}` 시트에서 읽은 값")
    prev = previous_period(period)
    fb = load_fund_balance(prev)
    if fb:
        return (float(fb["invest_end_krw"]), float(fb["invest_end_usd"] or 0.0),
                f"전월({prev}) 기말 운용자금 — 저장된 이력")
    pmb = load_month_balance(prev)
    if pmb and pmb["invest_krw"]:
        return (float(pmb["invest_krw"]), float(pmb["invest_usd"] or 0.0),
                f"전월({prev}) 일일자금계획에서 읽은 값")
    return (0.0, 0.0, "전월 기록이 없습니다 — 직접 입력해주세요")


# ── 보고서 표의 비고 칸 ──────────────────────────────────────────
def clean_text(v) -> str:
    """비고 칸 값을 '사람이 읽는 글자'로 정리합니다.

    편집표(st.data_editor)에서 칸을 **지우면** 판다스가 빈 문자열이 아니라
    `NaN`(숫자 결측값)을 돌려줍니다. 이걸 그냥 `str()` 하면 **"nan"** 이라는
    글자가 되어 엑셀·PPT에 그대로 찍혔습니다. 여기서 빈칸으로 되돌립니다.
    """
    if v is None:
        return ""
    if isinstance(v, float):                 # NaN 포함
        return "" if v != v else str(v)
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "<na>", "nat") else str(v)


def save_report_remarks(period: str, table_key: str, texts: list) -> None:
    conn = connect()
    for i, t in enumerate(texts):
        conn.execute(
            """INSERT INTO report_remarks (period, table_key, row_idx, text)
               VALUES (?,?,?,?)
               ON CONFLICT(period, table_key, row_idx) DO UPDATE SET text=excluded.text""",
            (period, table_key, i, clean_text(t)))
    conn.commit()
    conn.close()


def load_report_remarks_raw(period: str, table_key: str, n: int) -> list[str | None]:
    """비고를 **세 가지 상태**로 구분해서 돌려줍니다.

        None  = 아직 저장한 적 없음      → 화면·산출물에서 **앱 초안**을 씁니다
        ""    = 저장했는데 **일부러 비움** → 빈칸 그대로 (초안을 되살리지 않습니다)
        "글"  = 사용자가 쓴 글

    예전에는 둘 다 "" 로 돌려줘서, **지우려고 비운 줄에 초안이 다시 살아났습니다.**
    """
    conn = connect()
    rows = conn.execute(
        "SELECT row_idx, text FROM report_remarks WHERE period=? AND table_key=?",
        (period, table_key)).fetchall()
    conn.close()
    # 예전 버전이 "nan" 을 저장해 둔 경우도 여기서 빈칸으로 걸러냅니다.
    got = {r["row_idx"]: clean_text(r["text"]) for r in rows}
    return [got.get(i) for i in range(n)]


def load_report_remarks(period: str, table_key: str, n: int) -> list[str]:
    """저장된 비고 (저장 안 한 줄은 빈 문자열). 세 가지 상태 구분이 필요하면
    `load_report_remarks_raw` 를 쓰세요."""
    return [x or "" for x in load_report_remarks_raw(period, table_key, n)]


def clear_report_remarks(period: str, table_key: str) -> None:
    """저장 기록을 **지웁니다** — 다시 '저장한 적 없음' 상태가 되어 초안이 살아납니다."""
    conn = connect()
    conn.execute("DELETE FROM report_remarks WHERE period=? AND table_key=?",
                 (period, table_key))
    conn.commit()
    conn.close()


# ── 재확인 항목 ───────────────────────────────────────────────────
def load_recheck_items() -> list[dict]:
    conn = connect()
    rows = conn.execute("SELECT * FROM recheck_items WHERE active=1").fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── 4번 '최종 점검' 화면: 검증 게이트 확인 기록 ────────────────────
def txn_fingerprint(txns: list[Txn]) -> str:
    """거래 목록의 '지문'. 한 건이라도 바뀌면 값이 달라집니다.

    사용자가 [확인함]을 누른 시점을 이 지문으로 기억해 두었다가,
    나중에 거래가 바뀌면 확인 표시를 자동으로 풀기 위해 씁니다.
    """
    import hashlib
    parts = sorted(
        f"{t.id}|{t.tag}|{t.currency}|{t.amount_in:.4f}|{t.amount_out:.4f}|{t.excluded}"
        for t in txns
    )
    return hashlib.sha1("\n".join(parts).encode("utf-8")).hexdigest()[:16]


def save_gate_ack(period: str, scope: str, gate_id: str, fingerprint: str) -> None:
    conn = connect()
    conn.execute(
        """INSERT INTO gate_acks (period, scope, gate_id, fingerprint)
           VALUES (?,?,?,?)
           ON CONFLICT(period, scope, gate_id) DO UPDATE SET
             fingerprint=excluded.fingerprint,
             updated_at=datetime('now','localtime')""",
        (period, scope, gate_id, fingerprint))
    conn.commit()
    conn.close()


def clear_gate_ack(period: str, scope: str, gate_id: str) -> None:
    conn = connect()
    conn.execute("DELETE FROM gate_acks WHERE period=? AND scope=? AND gate_id=?",
                 (period, scope, gate_id))
    conn.commit()
    conn.close()


def load_gate_acks(period: str, scope: str) -> dict[str, dict]:
    """{게이트ID: {'fingerprint':…, 'updated_at':…}}"""
    conn = connect()
    rows = conn.execute("SELECT * FROM gate_acks WHERE period=? AND scope=?",
                        (period, scope)).fetchall()
    conn.close()
    return {r["gate_id"]: {"fingerprint": r["fingerprint"],
                           "updated_at": r["updated_at"]} for r in rows}


# ── 비고 '쓰는 방식'(서식) 학습 ────────────────────────────────────
def load_remark_style(table_key: str) -> dict | None:
    conn = connect()
    r = conn.execute("SELECT * FROM remark_styles WHERE table_key=?", (table_key,)).fetchone()
    conn.close()
    if not r:
        return None
    try:
        return json.loads(r["style_json"])
    except Exception:
        return None


def save_remark_style(table_key: str, style: dict) -> None:
    conn = connect()
    conn.execute(
        """INSERT INTO remark_styles (table_key, style_json, samples)
           VALUES (?,?,1)
           ON CONFLICT(table_key) DO UPDATE SET
             style_json=excluded.style_json,
             samples=remark_styles.samples + 1,
             updated_at=datetime('now','localtime')""",
        (table_key, json.dumps(style, ensure_ascii=False)))
    conn.commit()
    conn.close()
