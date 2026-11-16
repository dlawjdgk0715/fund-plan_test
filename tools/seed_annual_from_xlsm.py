"""산출물(.xlsm)의 '연간' 시트에서 월별 값을 읽어 DB에 넣습니다.

왜 필요한가
    2026-11 이전 버전은 연간 시트의 지난 달 값을 **템플릿에서 이어받는** 구조였습니다.
    템플릿이 7월 원본에 고정돼 있어서 8·9월이 빈칸으로 남았습니다.
    새 구조(`annual_months` 표)로 넘어오면서, 이미 만들어 둔 산출물에서 과거 달을
    한 번만 긁어 넣으면 그 뒤로는 자동으로 유지됩니다.

쓰는 법
    python tools/seed_annual_from_xlsm.py 산출물.xlsm --months 1-9 --kind actual
    python tools/seed_annual_from_xlsm.py 산출물.xlsm --months 10,11 --year 2026

    --months 를 생략하면 값이 들어 있는 달을 전부 넣습니다.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
from openpyxl.utils import column_index_from_string

from core.data_store import annual_saved_months, init_db, save_annual_month
from core.excel_writer import ANNUAL_LEAF_ROW, ANNUAL_SHEET_COL, SRC_ANNUAL


def _months(spec: str | None) -> list[int] | None:
    if not spec:
        return None
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out += list(range(int(a), int(b) + 1))
        elif part:
            out.append(int(part))
    return sorted(set(out))


def read_annual(path: Path, months: list[int] | None) -> dict[int, tuple[dict, dict]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    if SRC_ANNUAL not in wb.sheetnames:
        raise SystemExit(f"'{SRC_ANNUAL}' 시트가 없습니다: {path}")
    ws = wb[SRC_ANNUAL]
    found: dict[int, tuple[dict, dict]] = {}
    for m in (months or range(1, 13)):
        base = ANNUAL_SHEET_COL.get(m)
        if not base:
            continue
        kc = column_index_from_string(base)
        uc = kc + 1
        krw, usd, filled = {}, {}, 0
        for key, row in ANNUAL_LEAF_ROW.items():
            kv = ws.cell(row=row, column=kc).value
            uv = ws.cell(row=row, column=uc).value
            krw[key] = float(kv) if isinstance(kv, (int, float)) else 0.0
            usd[key] = float(uv) if isinstance(uv, (int, float)) else 0.0
            if isinstance(kv, (int, float)):
                filled += 1
        if filled:                      # 한 칸도 없는 달은 건너뜁니다
            found[m] = (krw, usd)
    return found


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsm")
    ap.add_argument("--months", default=None, help="예: 1-9  또는  10,11")
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--kind", default="", help="actual | plan (기록용)")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    data = read_annual(Path(a.xlsm), _months(a.months))
    if not data:
        raise SystemExit("값이 들어 있는 달을 못 찾았습니다.")

    print(f"{Path(a.xlsm).name} 에서 찾은 달: "
          + ", ".join(f"{m}월" for m in sorted(data)))
    if a.dry_run:
        for m in sorted(data):
            krw = data[m][0]
            print(f"  {m:2}월  급여 {krw.get('expense.salary', 0):>8.2f} "
                  f"/ 매출 {krw.get('income.sales', 0):>8.2f}")
        return

    init_db()
    for m in sorted(data):
        krw, usd = data[m]
        save_annual_month(a.year, m, krw, usd, kind=a.kind)
    print(f"DB 저장 완료 — {a.year}년 보관된 달: "
          + ", ".join(f"{m}월" for m in annual_saved_months(a.year)))


if __name__ == "__main__":
    main()
