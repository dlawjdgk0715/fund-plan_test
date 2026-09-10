"""엑셀 산출 (M5) — 원본 워크북을 템플릿으로 쓰고 수식은 그대로 살립니다.

※ 아래 좌표는 업로드해주신 원본 `카탈리스터_7월_자금운용계획_260714_.xlsm`을
   직접 열어 확인한 실제 값입니다.

동작 순서
  1) '6월데이터'/'7월데이터' 시트에 이번 달 거래를 써 넣음 → 원본 SUMIFS가 알아서 계산
  2) 사람이 손으로 넣던 값 입력: 기초 유동잔고(3p!G6,H6), 기초 운용자금(2-1!E6,F6), 환율
  3) 1-1+1-2(위쪽 블록만) → 요약1,  2-1+2-2 → 요약2
  4) 3p → 유동자금세부 로 이름 변경
  5) 2026년 자금실적 시트에 이번 달 지출 그룹값 채우기
  6) N열 비고, 요약1 서술형 비고 채우기
  7) 불필요 시트 제거
"""
from __future__ import annotations

import json
import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path

import openpyxl
# (Translator는 더 이상 쓰지 않습니다 — _shift_local_refs_in_formula 사용)
from openpyxl.utils import get_column_letter

from core.config import MILLION, OUTPUT_DIR, TEMPLATE_PATH
from core.data_store import Txn, clean_text
from core.engine import (ANNUAL_GROUP, EXTRAORDINARY_INCOME_ACCOUNTS,
                         ROW_NO, Result, aggregate)

# ── 원본 시트 이름 (수식이 이 이름을 참조하므로 절대 바꾸지 않음) ──
SRC_ACTUAL_DATA = "6월데이터"
SRC_PLAN_DATA = "7월데이터"
SRC_3P = "3p"
SRC_YEAR_CUR = "2026년 자금실적"
SRC_ANNUAL = "연간"

# 산출 파일에서 제외할 시트 (v0.8 + v0.12 '요약' 추가)
DROP_SHEETS = [
    "2022년 자금실적", "2023년 자금실적", "2024년 자금실적",
    "카탈리스터-연결(미법포함)", "카탈리스터-별도",
    "PPT (2)", "ppt보고용-5", "ppt보고용-6",
    "요약",   # v0.12 사용자 결정 — 매크로 함수를 쓰는 레거시 시트
]
# 2025년 자금실적은 2026년 시트가 전년비교로 참조 → 남겨둠

# ── 데이터 시트 (헤더 3행, 데이터 4행부터) ────────────────────────
DATA_HEADER_ROW = 3
DATA_START_ROW = 4
DATA_COLS = ["구분", "업체명", "적요", "통화", "입금", "출금", "외화금액"]
DATA_RATE_CELL = "I1"      # 환율이 적혀 있는 칸 (2-1!G6가 참조)

# ── 3p: 사람이 직접 넣던 칸 ───────────────────────────────────────
CELL_OPEN_LIQ_KRW = "G6"   # 기초유동잔고 원화 (실적)
CELL_OPEN_LIQ_USD = "H6"   # 기초유동잔고 외화($)
CELL_FX_RATE = "P4"        # 3p 수식들이 참조하는 환율 칸 (원본에는 비어 있었음 — 아래 주의 참고)

# ── 2-1: 기초 운용자금 ────────────────────────────────────────────
CELL_OPEN_INV_KRW = "E6"
CELL_OPEN_INV_USD = "F6"

# ── 3p N열 비고가 들어갈 행 (원본에 이미 예시 문구가 있던 자리) ───
REMARK_COL = 14   # N열

# ── 2026년 자금실적: 세부그룹 → 행 번호 (원본 확인 완료) ──────────
ANNUAL_GROUP_ROW = {
    # 경상
    "인건비": 17,
    "S/W, H/W 매입": 18,
    "마케팅,지급이자,외주용역비": 19,   # ← tools/add_marketing_row.py 로 추가한 행
    "제세공과금": 21,
    "전산관련지출": 22,
    "경영, 업무지원, 운영용역비": 23,
    "복리후생비": 24,
    "식당, 카페, 매점 등 운영비": 25,
    "직원비용, 복지포인트 사용비": 26,
    "임차료, 관리비": 27,
    "지급수수료": 28,
    "통신비등": 29,
    # 비경상
    "지식재산권 양수": 31,
    "임대보증금": 32,
    "시설장치": 33,
    "투자금": 34,
    "기타": 35,
}
# 1월=H(8) … 12월=M(13)
ANNUAL_MONTH_COL = {m: get_column_letter(7 + m) for m in range(1, 13)}

# ── 연간 시트: 1월=G(7), 각 월 3열(원화/외화/환산계) ──────────────
ANNUAL_SHEET_COL = {m: get_column_letter(7 + (m - 1) * 3) for m in range(1, 13)}
# 연간 시트 leaf 행 번호 (원본 확인 완료)
ANNUAL_LEAF_ROW = {
    # 수입
    "income.sales": 9, "income.taxrefund": 10, "income.interest": 11, "income.etc": 12,
    # 지출 — 인건비 (14 = 15~18 합)
    "expense.salary": 15, "expense.bonus": 16,
    "expense.pension": 17, "expense.insurance": 18,
    "expense.marketing": 19,
    # 지출 — 제세공과금 (20 = 21~24 합)
    "expense.tax.corp": 21, "expense.tax.vat": 22,
    "expense.tax.wht": 23, "expense.tax.etc": 24,
    "expense.interest": 25, "expense.capex": 26, "expense.invest": 27,
    # 지출 — 기타 (28 = 29~32 합).  32'기타' = 3p 32행 소계(통신비등·복리후생 등)
    "expense.it": 29, "expense.outsource": 30,
    "expense.support": 31, "expense.etc32": 32,
    # 운용자금 / FX / 조달
    "fund.release": 34, "fund.deposit": 35,
    "fx.in": 37, "fx.out": 38, "finance.in": 40, "finance.out": 41,
}

# 표 두 개를 한 시트로 합칠 때, 위 표 마지막 줄과 아래 표 제목 사이 간격
TABLE_GAP_ROWS = 5

# 산출 파일의 시트 순서 (여기 없는 시트는 뒤에 원래 순서대로 붙습니다)
SHEET_ORDER = ["요약1", "요약2", "유동자금세부", "연간", "실적월", "계획월",
               "2025년 자금실적", "2026년 자금실적"]

# 요약1에서 왼쪽으로 옮길 '자금수지' 표 영역 (제목 B열은 그대로 둠)
SUMMARY1_BLOCK_ROWS = (3, 20)

# 1-2 시트에서 가져올 범위 (아래쪽 "6월 계획" 블록은 #REF! 다수 → 무시)
SHEET_1_2_MAX_ROW = 14


# ── 템플릿 가볍게 만들기 ──────────────────────────────────────────
# 원본 워크북은 엑셀에서 오래 쓰이며 **안 쓰는 셀 스타일이 6만 개 넘게** 쌓여 있습니다
# (styles.xml 15.9MB). 그래서 여는 데만 10초가 걸립니다.
# 안 쓰는 스타일만 걷어낸 사본을 한 번 만들어 두고 그걸 씁니다.
#   원본 2,522KB / 9.8초  →  정리본 391KB / 0.9초  (숫자서식·글꼴·수식은 그대로)
# 원본 파일이 바뀌면(크기·수정시각) 자동으로 다시 만듭니다.
FAST_TEMPLATE = TEMPLATE_PATH.with_name("_fast_" + TEMPLATE_PATH.name)
FAST_TEMPLATE_META = TEMPLATE_PATH.with_name("_fast_" + TEMPLATE_PATH.stem + ".json")


def _template_for_build(warnings: list) -> Path:
    src = TEMPLATE_PATH
    info = src.stat()
    key = {"size": info.st_size, "mtime": int(info.st_mtime)}
    if FAST_TEMPLATE.exists() and FAST_TEMPLATE_META.exists():
        try:
            if json.loads(FAST_TEMPLATE_META.read_text(encoding="utf-8")) == key:
                return FAST_TEMPLATE
        except Exception:
            pass
    try:
        from openpyxl.styles.named_styles import NamedStyleList
        wb = openpyxl.load_workbook(src, keep_vba=True)
        before = len(wb._named_styles)
        keep = [ns for ns in wb._named_styles
                if ns.name in ("Normal", "표준", "일반")] or list(wb._named_styles)[:1]
        wb._named_styles = NamedStyleList(keep[:1])
        wb.save(FAST_TEMPLATE)
        FAST_TEMPLATE_META.write_text(json.dumps(key), encoding="utf-8")
        warnings.append(
            f"템플릿을 처음 한 번 가볍게 정리했습니다 — 안 쓰는 셀 스타일 "
            f"{before:,}개 제거. 다음부터는 훨씬 빨리 만들어집니다.")
        return FAST_TEMPLATE
    except Exception as e:                                   # 실패하면 원본 그대로
        warnings.append(f"템플릿 정리에 실패해 원본을 그대로 씁니다 ({e}).")
        return src


@dataclass
class WriteReport:
    path: Path
    warnings: list[str]


# ── 시트 복사 유틸 ────────────────────────────────────────────────
def _copy_block(src, dst, row_offset=0, max_row=None):
    limit = max_row or src.max_row
    for row in src.iter_rows(min_row=1, max_row=limit):
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            r, c = cell.row + row_offset, cell.column
            new = dst.cell(row=r, column=c)
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and row_offset:
                # 같은 시트를 가리키는 참조만 밀어줍니다.
                # 다른 시트 참조('3p'!G49 등)는 건드리지 않고 _patch_cross_sheet_refs가 처리합니다.
                # (v0.6 버그: Translator를 그대로 쓰면 다른 시트 참조의 행 번호까지 밀려버림)
                v = _shift_local_refs_in_formula(v, row_offset)
            new.value = v
            if cell.has_style:
                new.font = copy(cell.font)
                new.border = copy(cell.border)
                new.fill = copy(cell.fill)
                new.number_format = cell.number_format
                new.alignment = copy(cell.alignment)
    for rng in src.merged_cells.ranges:
        if max_row and rng.min_row > max_row:
            continue
        try:
            dst.merge_cells(start_row=rng.min_row + row_offset, start_column=rng.min_col,
                            end_row=rng.max_row + row_offset, end_column=rng.max_col)
        except Exception:
            pass
    for key, dim in src.column_dimensions.items():
        dst.column_dimensions[key].width = dim.width


SHEET_REF = re.compile(r"'([^']+)'!(\$?)([A-Z]{1,3})(\$?)(\d+)")


# 시트 이름 뒤에 붙는 '!' 까지 한 덩어리로 잡습니다.
#   'ppt보고용-6'!  ·  연간!  ·  '6월데이터'!
SHEET_PREFIX = re.compile(r"('(?P<q>(?:[^']|'')+)'|(?<![A-Za-z0-9_.!$])(?P<b>[A-Za-z_][A-Za-z0-9_.]*))!")


def _patch_cross_sheet_refs(wb, old_sheet, new_sheet, row_offset=0):
    """다른 시트를 가리키는 참조를 새 이름(+행 보정)으로 고쳐 씁니다.

    ⚠️ 예전에는 `'시트'!A1` 처럼 **행 번호가 붙은 참조만** 잡았습니다.
       그런데 원본 3p의 SUMIFS는 `'6월데이터'!$E:$E` 처럼 **열 전체**를 참조합니다
       (행 번호가 없음). 그래서 시트 이름을 바꾸면 이 수식들이 없는 시트를
       가리킨 채 남아 파일 전체가 #REF!/#VALUE! 로 깨졌습니다.
       → 1차로 행 보정이 필요한 참조를 처리하고, 2차로 **남은 이름표를 전부** 바꿉니다.
    """
    def rename_prefix(text):
        def repl(m):
            name = m.group("q") or m.group("b")
            if name.replace("''", "'") != old_sheet:
                return m.group(0)
            return f"'{new_sheet}'!"
        return SHEET_PREFIX.sub(repl, text)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue

                def repl(m):
                    if m.group(1) != old_sheet:
                        return m.group(0)
                    return (f"'{new_sheet}'!{m.group(2)}{m.group(3)}{m.group(4)}"
                            f"{int(m.group(5)) + row_offset}")

                v = SHEET_REF.sub(repl, cell.value) if row_offset else cell.value
                cell.value = rename_prefix(v)


LOCAL_REF = re.compile(r"(?<![A-Z0-9_$!'])(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _shift_plain(text, row_offset):
    def repl(m):
        if m.group(3) == "$":
            return m.group(0)
        return f"{m.group(1)}{m.group(2)}{m.group(3)}{int(m.group(4)) + row_offset}"
    return LOCAL_REF.sub(repl, text)


def _shift_local_refs_in_formula(formula, row_offset):
    """v0.6 버그 수정 — 같은 시트 참조와 다른 시트 참조가 섞인 수식에서
    같은 시트 참조만 정확히 밀어줍니다."""
    out, last = [], 0
    for m in SHEET_REF.finditer(formula):
        out.append(_shift_plain(formula[last:m.start()], row_offset))
        out.append(m.group(0))
        last = m.end()
    out.append(_shift_plain(formula[last:], row_offset))
    return "".join(out)


def _shift_local_cols(formula: str, delta: int) -> str:
    """같은 시트 참조의 **열**만 옮깁니다. 다른 시트 참조('유동자금세부'!I8)는 그대로."""
    def bump(text):
        def repl(m):
            if m.group(1) == "$":
                return m.group(0)
            idx = openpyxl.utils.column_index_from_string(m.group(2)) + delta
            if idx < 1:
                return m.group(0)
            return f"{m.group(1)}{get_column_letter(idx)}{m.group(3)}{m.group(4)}"
        return LOCAL_REF.sub(repl, text)

    out, last = [], 0
    for m in SHEET_REF.finditer(formula):
        out.append(bump(formula[last:m.start()]))
        out.append(m.group(0))
        last = m.end()
    out.append(bump(formula[last:]))
    return "".join(out)


def _move_block_cols(ws, row_lo, row_hi, col_lo, col_hi, delta):
    """표 한 덩어리를 좌우로 옮깁니다 (값·서식·병합·수식 참조 모두 함께)."""
    if delta == 0:
        return
    cells = []
    for r in range(row_lo, row_hi + 1):
        for c in range(col_lo, col_hi + 1):
            src = ws.cell(row=r, column=c)
            style = None
            if src.has_style:
                style = (copy(src.font), copy(src.border), copy(src.fill),
                         src.number_format, copy(src.alignment))
            cells.append((r, c, src.value, style))
    merges = [rng for rng in list(ws.merged_cells.ranges)
              if row_lo <= rng.min_row and rng.max_row <= row_hi
              and col_lo <= rng.min_col and rng.max_col <= col_hi]
    for rng in merges:
        ws.unmerge_cells(str(rng))
    for r, c, _v, _st in cells:                      # 원래 자리를 비웁니다
        ws.cell(row=r, column=c).value = None
    for r, c, v, st in cells:
        if isinstance(v, str) and v.startswith("="):
            v = _shift_local_cols(v, delta)
        dst = ws.cell(row=r, column=c + delta)
        dst.value = v
        if st:
            dst.font, dst.border, dst.fill, dst.number_format, dst.alignment = st
    for rng in merges:
        ws.merge_cells(start_row=rng.min_row, start_column=rng.min_col + delta,
                       end_row=rng.max_row, end_column=rng.max_col + delta)


def _first_content_row(ws) -> int:
    """값이 실제로 들어 있는 첫 행."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                return cell.row
    return 1


def _order_sheets(wb, order):
    """지정한 순서대로 시트를 정렬합니다. 목록에 없는 시트는 뒤에 그대로 둡니다."""
    wanted = [n for n in order if n in wb.sheetnames]
    rest = [n for n in wb.sheetnames if n not in wanted]
    wb._sheets = [wb[n] for n in wanted + rest]


def _last_content_row(ws) -> int:
    """값이 실제로 들어 있는 마지막 행 (서식만 있는 빈 행은 안 셈)."""
    last = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value not in (None, ""):
                last = max(last, cell.row)
                break
    return last


# ── 원시 데이터 채우기 ────────────────────────────────────────────
def _fill_data_sheet(ws, txns, fx_rate):
    if ws.max_row >= DATA_START_ROW:
        ws.delete_rows(DATA_START_ROW, ws.max_row - DATA_START_ROW + 1)
    for j, h in enumerate(DATA_COLS, start=1):
        ws.cell(row=DATA_HEADER_ROW, column=j, value=h)
    for i, t in enumerate(txns, start=DATA_START_ROW):
        is_usd = (t.currency or "KRW").upper() != "KRW"
        ws.cell(row=i, column=1, value=(t.tag or "").strip())
        ws.cell(row=i, column=2, value=t.counterpart)
        ws.cell(row=i, column=3, value=t.memo)
        ws.cell(row=i, column=4, value=(t.currency or "KRW").upper())
        ws.cell(row=i, column=5, value=t.amount_in or None)
        ws.cell(row=i, column=6, value=t.amount_out or None)
        ws.cell(row=i, column=7, value=(t.amount if is_usd else None))
    if fx_rate:
        ws[DATA_RATE_CELL] = fx_rate   # 데이터 시트 환율 (2-1 시트가 참조)


# ── 메인 ──────────────────────────────────────────────────────────
def build_workbook(actual_txns, plan_txns, actual_result: Result,
                   actual_month: int, plan_month: int,
                   opening_liq_krw: float, opening_liq_usd: float,
                   opening_inv_krw: float, opening_inv_usd: float,
                   fx_rate: float,
                   leaf_remarks: dict[str, str] | None = None,
                   summary_drafts: dict[str, str] | None = None,
                   out_name: str | None = None,
                   draft: bool = False) -> WriteReport:
    warnings: list[str] = []
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"원본 템플릿을 못 찾았습니다: {TEMPLATE_PATH}\n"
            "원본 자금운용계획 워크북을 templates/자금운용계획_template.xlsm 로 복사해 넣어주세요."
        )
    # 원본이 매크로 포함(.xlsm)이라, 산출 파일도 .xlsm 으로 만듭니다.
    # (.xlsx 로 저장하면 엑셀이 "매크로 제외 파일인데 매크로 내용이 있다"고 경고합니다)
    wb = openpyxl.load_workbook(_template_for_build(warnings), keep_vba=True)

    # 1) 원시 데이터
    for name, txns in ((SRC_ACTUAL_DATA, actual_txns), (SRC_PLAN_DATA, plan_txns)):
        if name in wb.sheetnames:
            _fill_data_sheet(wb[name], txns, fx_rate)
        else:
            warnings.append(f"원시 데이터 시트 '{name}'을 템플릿에서 못 찾았습니다.")

    # 2) 손으로 넣던 값
    if SRC_3P in wb.sheetnames:
        ws3 = wb[SRC_3P]
        ws3[CELL_OPEN_LIQ_KRW] = opening_liq_krw
        ws3[CELL_OPEN_LIQ_USD] = opening_liq_usd
        ws3[CELL_FX_RATE] = fx_rate
    else:
        warnings.append(f"'{SRC_3P}' 시트가 없습니다.")
    if "2-1" in wb.sheetnames:
        wb["2-1"][CELL_OPEN_INV_KRW] = opening_inv_krw
        wb["2-1"][CELL_OPEN_INV_USD] = opening_inv_usd

    use_ref = STYLE_REF_PATH.exists()

    # 4-a) 3p 집계 수식을 태그 사전 기준으로 통일
    _rewrite_3p_formulas(wb, warnings)

    # 4-a2) 1-1(요약1)의 **실적월 기초잔고**도 우리 값으로. 템플릿에 지난달 숫자가
    #       그대로 박혀 있어서(2,488.757327) 요약1 기말만 다른 시트와 어긋났습니다.
    if "1-1" in wb.sheetnames and not use_ref:
        ws11 = wb["1-1"]
        base_row = base_col = None
        for r in range(1, min(ws11.max_row, 30) + 1):
            for c in range(1, 10):
                v = ws11.cell(row=r, column=c).value
                if isinstance(v, str) and "기초잔고" in v:
                    base_row, base_col = r, c
                    break
            if base_row:
                break
        if base_row:
            ws11.cell(row=base_row, column=base_col + 1,
                      value=opening_liq_krw + opening_inv_krw)
        else:
            warnings.append("요약1에서 '기초잔고' 행을 못 찾아 기초잔고를 못 채웠습니다.")

    # 5) 자금실적(당해)·연간 — 시트 합치기 전에 처리.
    #    실적월과 계획월 **둘 다** 채웁니다 (사용자 결정 2026-08-25).
    plan_result = aggregate(plan_txns, fx_rate=fx_rate or None)
    # 실적월 기초잔고 = 유동 + 운용 (전월말 실제 잔고)
    open_krw_total = opening_liq_krw + opening_inv_krw
    open_usd_total = opening_liq_usd + opening_inv_usd
    _fill_annual_actual(wb, actual_result, plan_result, actual_month, plan_month, warnings,
                        open_total=open_krw_total)
    _fill_annual_sheet(wb, actual_result, actual_month, warnings,
                       opening=(open_krw_total, open_usd_total))
    _fill_annual_sheet(wb, plan_result, plan_month, warnings)

    # 7) 월 표시 자동화
    _fix_month_labels(wb, actual_month, plan_month)

    # 8) 비고 — **월 표시 자동화 뒤에** 채웁니다.
    #    (먼저 채우면 비고 안의 '7월'까지 계획월로 바뀌어 버립니다)
    # 8-a) 3p N열 비고
    if SRC_3P in wb.sheetnames:
        ws3 = wb[SRC_3P]
        # 지난달 문구를 먼저 지웁니다. 안 지우면 거래가 없는 행에 **지난달 내용이
        # 이번 달 이름표를 달고** 그대로 남습니다
        # (예: 6월 정기예금 해지 문구가 '∙7월 : …'로 둔갑)
        for r in set(ROW_NO.values()):
            ws3.cell(row=r, column=REMARK_COL).value = None
        for row_key, text in (leaf_remarks or {}).items():
            r = ROW_NO.get(row_key)
            if r:
                # clean_text: 편집표에서 칸을 지우면 오던 NaN → 빈칸 ("nan" 방지)
                ws3.cell(row=r, column=REMARK_COL, value=(clean_text(text) or None))

    # 8-b) 1-1 서술형 비고 — 원본 규칙 그대로 (사용자 확인 2026-08-25)
    #      윗줄 = ● 실적월 / 아랫줄 = ○ 계획월,  수입·지출 각각 2줄
    #        I8  ● 7월 : 정기예금 만기해지 이자 6.4백만 등     (실적 수입)
    #        I9  ○ 8월 :                                     (계획 수입)
    #        I10 ● 7월 : 급여 15백만 등                        (실적 지출)
    #        I11 ○ 8월 : 급여 15백만, 부가세 납부 3.7백만 등     (계획 지출)
    if "1-1" in wb.sheetnames and not use_ref:
        ws11 = wb["1-1"]
        drafts = summary_drafts or {}
        for cell, key, mark, mon in (("I8", "actual_in", "●", actual_month),
                                     ("I9", "plan_in", "○", plan_month),
                                     ("I10", "actual_out", "●", actual_month),
                                     ("I11", "plan_out", "○", plan_month)):
            body = (drafts.get(key) or "").strip()
            # 초안에 이미 "7월 : " 이 들어 있으면 그대로, 없으면 만들어 붙입니다.
            if body and not re.match(r"^\s*[●○]", body):
                if not re.match(rf"^\s*{mon}월\s*:", body):
                    body = f"{mon}월 : {body}"
                body = f"{mark} {body}"
            elif not body:
                body = f"{mark} {mon}월 : "
            ws11[cell] = body


    # 3) 시트 합치기 (기준 파일이 있으면 그 시트를 그대로 쓰므로 건너뜁니다)
    if not use_ref:
        _merge_pair(wb, "1-1", "1-2", "요약1", warnings, max_row_b=SHEET_1_2_MAX_ROW)
        _merge_pair(wb, "2-1", "2-2", "요약2", warnings)

    # 3-b) 요약1: 위 표(자금수지)가 E열부터 시작해 아래 표(유동자금 계획, C열)와
    #      어긋나 보였습니다 → 위 표를 두 칸 왼쪽으로 옮겨 시작 열을 맞춥니다.
    if "요약1" in wb.sheetnames and not use_ref:
        _move_block_cols(wb["요약1"], SUMMARY1_BLOCK_ROWS[0], SUMMARY1_BLOCK_ROWS[1],
                         5, 9, -2)          # E~I → C~G

    # 4) 이름 변경
    if SRC_3P in wb.sheetnames:
        wb[SRC_3P].title = "유동자금세부"
        _patch_cross_sheet_refs(wb, SRC_3P, "유동자금세부", 0)

    # 4-b) 원시 데이터 시트 이름 — '6월데이터/7월데이터'는 헷갈리므로
    #      '실적월/계획월'로 바꿉니다 (수식 참조도 함께 고침).
    for src, new_title in ((SRC_ACTUAL_DATA, "실적월"), (SRC_PLAN_DATA, "계획월")):
        if src in wb.sheetnames:
            wb[src].title = new_title
            _patch_cross_sheet_refs(wb, src, new_title, 0)

    # 4-b2) 기준 파일 서식 적용 — 시트 이름을 바꾼 **뒤**에 해야 합니다.
    #        (기준 파일의 수식이 '유동자금세부'·'실적월'·'계획월'을 참조하므로)
    if use_ref:
        if not _apply_style_reference(wb, actual_month, plan_month,
                                      opening_liq_krw, opening_liq_usd,
                                      opening_inv_krw, opening_inv_usd,
                                      summary_drafts, warnings):
            _merge_pair(wb, "1-1", "1-2", "요약1", warnings, max_row_b=SHEET_1_2_MAX_ROW)
            _merge_pair(wb, "2-1", "2-2", "요약2", warnings)
            if "요약1" in wb.sheetnames:
                _move_block_cols(wb["요약1"], SUMMARY1_BLOCK_ROWS[0],
                                 SUMMARY1_BLOCK_ROWS[1], 5, 9, -2)

    for name in list(DROP_SHEETS) + ["1-1", "1-2", "2-1", "2-2"]:
        if name in wb.sheetnames:
            del wb[name]

    # 4-c) 시트 순서 (사용자 지정 2026-08-25)
    _order_sheets(wb, SHEET_ORDER)

    if draft:
        wb[wb.sheetnames[0]]["A1"] = "※ 검증 미완료 — 확정본 아님"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    name = out_name or f"자금운용계획_{actual_month}월실적_{plan_month}월계획.xlsm"
    if not name.lower().endswith((".xlsm", ".xlsx")):
        name += ".xlsm"
    path = OUTPUT_DIR / name
    wb.save(path)
    return WriteReport(path=path, warnings=warnings)


def _merge_pair(wb, a, b, new_name, warnings, max_row_b=None):
    if a not in wb.sheetnames:
        warnings.append(f"'{a}' 시트가 없어 {new_name} 합치기를 건너뜁니다.")
        return
    src_a = wb[a]
    dst = wb.create_sheet(new_name)
    _copy_block(src_a, dst, 0)
    # 표 사이 간격을 두 시트가 똑같도록 — '값이 있는 마지막 행' 기준으로 띄웁니다.
    # (예전엔 max_row 기준이라 서식만 있는 빈 행까지 세서 요약2가 46행이나 벌어졌음)
    offset = _last_content_row(src_a) + TABLE_GAP_ROWS
    if b in wb.sheetnames:
        # 아래 표가 원본 시트에서 몇 번째 행부터 시작하는지 빼줘야
        # 두 시트의 '빈 줄 개수'가 똑같아집니다.
        offset -= _first_content_row(wb[b]) - 1
        _copy_block(wb[b], dst, offset, max_row=max_row_b)
        _patch_cross_sheet_refs(wb, b, new_name, offset)
        del wb[b]
    _patch_cross_sheet_refs(wb, a, new_name, 0)
    del wb[a]


def _norm_label(v) -> str:
    return re.sub(r"\s+", "", str(v)) if v is not None else ""


def _label_rows(ws, min_row=1, max_row=None, cols=(2, 3, 4, 5)):
    """시트에서 '라벨 글자 → 행 번호' 표를 만듭니다 (왼쪽 열부터 처음 나오는 글자)."""
    out: dict[str, int] = {}
    for r in range(min_row, (max_row or ws.max_row) + 1):
        for c in cols:
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                out.setdefault(_norm_label(v), r)
                break
    return out


def _shift_rows_from(wb, sheet_name, at):
    """sheet_name 을 가리키는 참조 중 at 이상인 행을 전부 +1 (행 삽입 보정)."""
    def bump_plain(text):
        def repl(m):
            row = int(m.group(4))
            return f"{m.group(1)}{m.group(2)}{m.group(3)}{row + 1 if row >= at else row}"
        return LOCAL_REF.sub(repl, text)

    for ws in wb.worksheets:
        same = ws.title == sheet_name
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                out, last = [], 0
                for m in SHEET_REF.finditer(v):
                    seg = v[last:m.start()]
                    out.append(bump_plain(seg) if same else seg)
                    if m.group(1) == sheet_name:
                        r = int(m.group(5))
                        out.append(f"'{sheet_name}'!{m.group(2)}{m.group(3)}{m.group(4)}"
                                   f"{r + 1 if r >= at else r}")
                    else:
                        out.append(m.group(0))
                    last = m.end()
                seg = v[last:]
                out.append(bump_plain(seg) if same else seg)
                cell.value = "".join(out)


def _insert_marketing_row(wb, ws, after_row: int, warnings) -> int:
    """'마케팅,지급이자,외주용역비' 행을 S/W,H/W 매입 바로 아래에 끼워 넣습니다.

    원본 템플릿에는 이 행이 없어서, 산출할 때마다 앱이 직접 만들어 넣습니다
    (사용자 결정 2026-08-25 — 원본 워크북은 건드리지 않음).
    """
    at = after_row + 1
    ws.insert_rows(at)
    _shift_rows_from(wb, ws.title, at)
    # 윗 행(S/W, H/W 매입)의 서식·라벨 위치를 그대로 따릅니다.
    src_row = at - 1
    label_col = 4
    for c in range(2, 6):
        v = ws.cell(row=src_row, column=c).value
        if isinstance(v, str) and v.strip():
            label_col = c
            break
    for c in range(1, ws.max_column + 1):
        src, dst = ws.cell(row=src_row, column=c), ws.cell(row=at, column=c)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)
        dst.value = None
    ws.cell(row=at, column=label_col, value=MARKETING_GROUP)
    warnings.append(f"'{MARKETING_GROUP}' 행이 자금실적 시트에 없어 {at}행에 새로 넣었습니다.")
    return at


_MONTH_HEADER = re.compile(r"^\s*(\d{1,2})\s*월\s*(\(A\))?\s*$")


def _mark_actual_header(ws, actual_month: int, warnings):
    """자금실적 시트의 **실적월 머리글에 '(A)'** 를 붙입니다.

    이 시트는 확정된 달만 '1월(A) … 6월(A)' 처럼 (A)를 달고, 아직 예상치인 달은
    '7월'처럼 그냥 둡니다. 실적을 확정해 채워 넣었으니 그 달도 (A)가 돼야 합니다.
    (사용자 결정 2026-08-25.  계획월과 그 뒤 달은 손대지 않습니다.)
    """
    col = ANNUAL_MONTH_COL.get(actual_month)
    if not col:
        return
    idx = openpyxl.utils.column_index_from_string(col)
    for r in range(1, min(ws.max_row, 20) + 1):
        v = ws.cell(row=r, column=idx).value
        if not isinstance(v, str):
            continue
        m = _MONTH_HEADER.match(v)
        if not m or int(m.group(1)) != actual_month:
            continue
        if not m.group(2):
            ws.cell(row=r, column=idx).value = f"{actual_month}월(A)"
        return
    warnings.append(f"자금실적 시트에서 '{actual_month}월' 머리글을 못 찾아 (A)를 못 붙였습니다.")


MARKETING_GROUP = "마케팅,지급이자,외주용역비"
SW_HW_GROUP = "S/W, H/W 매입"


# ── 3p 집계 수식 다시 쓰기 ────────────────────────────────────────
# 원본 워크북의 SUMIFS 기준이 행마다 제각각이라 **엔진과 다른 숫자**가 나왔습니다.
#   · J29(전산관련지출 계획) 은 기준을 라벨 셀 `F29`("전산관련지출")로 잡는데
#     데이터의 계정과목은 "전산유지비" → 늘 0. 8월 지출이 0.684백만 덜 잡혔습니다.
#   · G27(투자금 실적)은 '지분투자'를 더하는데 J27(계획)은 안 더합니다.
# → 산출할 때 3p의 집계 행 수식을 **태그 사전 기준으로 통일해서** 다시 써넣습니다.
#   그래야 요약1·PPT·연간·자금실적이 전부 같은 숫자가 됩니다.
COL_ACTUAL_KRW, COL_ACTUAL_USD = 7, 8     # G, H
COL_PLAN_KRW, COL_PLAN_USD = 10, 11       # J, K
DATA_COL_IN, DATA_COL_OUT = "E", "F"      # 데이터 시트의 수입·지출 열


def _accounts_for(key: str, direction: str) -> list[str]:
    from core.engine import ACCOUNT_MAP
    return sorted({a for (a, d), k in ACCOUNT_MAP.items() if k == key and d == direction})


def _sumifs(sheet: str, amount_col: str, accounts: list[str], currency: str) -> str:
    div = "1000000" if currency == "KRW" else "1000"
    parts = [f"SUMIFS('{sheet}'!${amount_col}:${amount_col},"
             f"'{sheet}'!$A:$A,\"{a}\",'{sheet}'!$D:$D,\"{currency}\")/{div}"
             for a in accounts]
    return "=" + "+".join(parts) if parts else "=0"


def _rewrite_3p_formulas(wb, warnings):
    if SRC_3P not in wb.sheetnames:
        return
    ws = wb[SRC_3P]
    from core.engine import LEAF
    fixed = 0
    for key, row, label, direction, _accts in LEAF:
        accounts = _accounts_for(key, direction)
        if not accounts:
            continue
        amount_col = DATA_COL_IN if direction == "in" else DATA_COL_OUT
        for col, sheet, cur in ((COL_ACTUAL_KRW, SRC_ACTUAL_DATA, "KRW"),
                                (COL_ACTUAL_USD, SRC_ACTUAL_DATA, "USD"),
                                (COL_PLAN_KRW, SRC_PLAN_DATA, "KRW"),
                                (COL_PLAN_USD, SRC_PLAN_DATA, "USD")):
            new = _sumifs(sheet, amount_col, accounts, cur)
            if ws.cell(row=row, column=col).value != new:
                ws.cell(row=row, column=col).value = new
                fixed += 1
    if fixed:
        warnings.append(
            f"유동자금세부 집계 수식 {fixed}칸을 태그 사전 기준으로 통일했습니다 "
            "(원본 수식이 행마다 기준이 달라 일부 항목이 0으로 빠지고 있었습니다).")


# ── 사용자가 손본 서식 그대로 쓰기 ────────────────────────────────
# `templates/서식기준.xlsm` 에 **지난달 산출물을 직접 다듬은 파일**을 넣어두면,
# 요약1·요약2를 새로 조립하지 않고 그 파일의 시트를 통째로 가져다 씁니다.
# 이 두 시트는 거의 전부 수식(유동자금세부·요약2 참조)이라, 숫자는 알아서 새로 계산되고
# **열 너비·행 높이·글꼴·테두리·표 위치는 손보신 그대로** 유지됩니다.
# 우리가 갈아 끼우는 것은 ① 월 표기 ② 기초잔고 ③ 비고 문구, 세 가지뿐입니다.
STYLE_REF_PATH = TEMPLATE_PATH.with_name("서식기준.xlsm")
STYLE_REF_SHEETS = ("요약1", "요약2")


def _copy_sheet_whole(src_ws, dst_wb, title):
    """시트를 값·수식·서식·병합·열너비·행높이까지 통째로 복사합니다."""
    if title in dst_wb.sheetnames:
        del dst_wb[title]
    dst = dst_wb.create_sheet(title)
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            new = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new.font = copy(cell.font)
                new.border = copy(cell.border)
                new.fill = copy(cell.fill)
                new.number_format = cell.number_format
                new.alignment = copy(cell.alignment)
    for rng in src_ws.merged_cells.ranges:
        try:
            dst.merge_cells(str(rng))
        except Exception:
            pass
    for k, d in src_ws.column_dimensions.items():
        if d.width:
            dst.column_dimensions[k].width = d.width
        dst.column_dimensions[k].hidden = d.hidden
    for k, d in src_ws.row_dimensions.items():
        if d.height:
            dst.row_dimensions[k].height = d.height
    try:
        dst.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
    except Exception:
        pass
    return dst


_HDR_AB = re.compile(r"(\d{1,2})\s*월\s*\(\s*([ab])\s*\)")


def _detect_ref_months(ws) -> tuple[int, int] | None:
    """기준 파일이 몇 월 기준으로 만들어졌는지 'N월(a)' / 'M월(b)' 로 알아냅니다."""
    a = b = None
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20)):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            m = _HDR_AB.search(cell.value)
            if m:
                if m.group(2) == "a":
                    a = int(m.group(1))
                else:
                    b = int(m.group(1))
    return (a, b) if a and b else None


def _swap_months_in(ws, old_a, old_p, new_a, new_p):
    table = {str(old_a): new_a, str(old_p): new_p}
    pat = re.compile(rf"(?<!\d)({old_a}|{old_p})월")
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and "월" in v and not v.startswith("="):
                cell.value = pat.sub(lambda m: f"{table[m.group(1)]}월", v)


def _find_row_by_label(ws, text, max_col=8):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.column <= max_col \
                    and text in cell.value:
                return cell.row, cell.column
    return None, None


def _apply_style_reference(wb, actual_month, plan_month,
                           opening_liq_krw, opening_liq_usd,
                           opening_inv_krw, opening_inv_usd,
                           summary_drafts, warnings) -> bool:
    """기준 파일이 있으면 요약1·요약2를 그 서식으로 만듭니다. 썼으면 True."""
    if not STYLE_REF_PATH.exists():
        return False
    try:
        ref = openpyxl.load_workbook(STYLE_REF_PATH)
    except Exception as e:
        warnings.append(f"서식기준 파일을 못 읽어 기본 서식으로 만듭니다 ({e}).")
        return False
    missing = [n for n in STYLE_REF_SHEETS if n not in ref.sheetnames]
    if missing:
        warnings.append(f"서식기준 파일에 {', '.join(missing)} 시트가 없어 기본 서식으로 만듭니다.")
        return False

    for name in ("1-1", "1-2", "2-1", "2-2"):
        if name in wb.sheetnames:
            del wb[name]
    for name in STYLE_REF_SHEETS:
        _copy_sheet_whole(ref[name], wb, name)

    ws1, ws2 = wb["요약1"], wb["요약2"]
    months = _detect_ref_months(ws1)
    if months and months != (actual_month, plan_month):
        for ws in (ws1, ws2):
            _swap_months_in(ws, months[0], months[1], actual_month, plan_month)
    elif not months:
        warnings.append("서식기준 파일에서 'N월(a)' 머리글을 못 찾아 월 표기를 못 바꿨습니다.")

    # ① 요약1 기초잔고 (실적월 칸)
    r, c = _find_row_by_label(ws1, "기초잔고")
    if r:
        ws1.cell(row=r, column=c + 1, value=opening_liq_krw + opening_inv_krw)
    else:
        warnings.append("서식기준 요약1에서 '기초잔고' 행을 못 찾았습니다.")

    # ② 요약2 A. 기초자금 (실적월 원화·외화)
    r2, c2 = _find_row_by_label(ws2, "기초자금")
    if r2:
        ws2.cell(row=r2, column=c2 + 2, value=opening_inv_krw)
        ws2.cell(row=r2, column=c2 + 3, value=opening_inv_usd)

    # ③ 요약1 비고 — 수입·지출 각각 2줄(● 실적월 / ○ 계획월)
    hdr_r, hdr_c = _find_row_by_label(ws1, "비고")
    drafts = summary_drafts or {}
    if hdr_r:
        for label, keys in ((" B. 수입", ("actual_in", "plan_in")),
                            (" C. 지출", ("actual_out", "plan_out"))):
            rr, _ = _find_row_by_label(ws1, label.strip())
            if not rr:
                continue
            for i, (key, mark, mon) in enumerate(
                    ((keys[0], "●", actual_month), (keys[1], "○", plan_month))):
                body = (drafts.get(key) or "").strip()
                if body and not re.match(r"^\s*[●○]", body):
                    if not re.match(rf"^\s*{mon}월\s*:", body):
                        body = f"{mon}월 : {body}"
                    body = f"{mark} {body}"
                elif not body:
                    body = f"{mark} {mon}월 : "
                ws1.cell(row=rr + i, column=hdr_c, value=body)
    # ④ 나머지 시트(유동자금세부·연간·자금실적 등)의 열 너비·행 높이도 맞춥니다.
    for name in wb.sheetnames:
        if name in STYLE_REF_SHEETS or name not in ref.sheetnames:
            continue
        src, dst = ref[name], wb[name]
        for k, d in src.column_dimensions.items():
            if d.width:
                dst.column_dimensions[k].width = d.width
            dst.column_dimensions[k].hidden = d.hidden
        for k, d in src.row_dimensions.items():
            if d.height:
                dst.row_dimensions[k].height = d.height

    warnings.append("서식기준 파일의 요약1·요약2 서식(열 너비·표 위치·글꼴)을 그대로 적용했습니다.")
    return True


def _fill_annual_actual(wb, actual_result: Result, plan_result: Result | None,
                        actual_month: int, plan_month: int, warnings,
                        open_total: float | None = None):
    """2026년 자금실적 시트에 실적월·계획월 지출 세부그룹 합계를 채웁니다.

    행 번호를 코드에 박지 않고 **시트의 라벨 글자를 읽어** 찾습니다
    (사용자 결정 2026-08-25 — 행이 추가·삭제돼도 안 깨지도록).
    """
    if SRC_YEAR_CUR not in wb.sheetnames:
        warnings.append(f"'{SRC_YEAR_CUR}' 시트가 없어 자금실적 반영을 건너뜁니다.")
        return
    ws = wb[SRC_YEAR_CUR]

    def totals_of(result):
        # ⚠️ 0인 그룹도 **반드시 0으로 덮어써야** 합니다.
        #    안 그러면 템플릿(지난달 파일)에 남아 있던 값이 그대로 살아남아
        #    지출 합계가 부풀려집니다 (예: 식당·카페 0.097 / 지급수수료 0.3).
        out: dict[str, float] = {g: 0.0 for g, _ in ANNUAL_GROUP.values()}
        if result is None:
            return out
        for key, val in result.values.items():
            if key in ANNUAL_GROUP:
                grp = ANNUAL_GROUP[key][0]
                out[grp] = out.get(grp, 0.0) + val
        return out

    jobs = [(actual_month, totals_of(actual_result))]
    if plan_result is not None and plan_month:
        jobs.append((plan_month, totals_of(plan_result)))

    need_marketing = any(abs(t.get(MARKETING_GROUP, 0.0)) > 1e-12 for _, t in jobs)
    rows = _label_rows(ws, 11, 45)
    if need_marketing and _norm_label(MARKETING_GROUP) not in rows:
        anchor = rows.get(_norm_label(SW_HW_GROUP))
        if anchor:
            _insert_marketing_row(wb, ws, anchor, warnings)
            rows = _label_rows(ws, 11, 45)
        else:
            warnings.append(f"'{SW_HW_GROUP}' 행을 못 찾아 '{MARKETING_GROUP}' 행을 넣지 못했습니다.")

    # B. 수입 — 시트에는 '경상 / 비경상' 두 줄뿐입니다.
    # 확정 규칙: 자산매각대금·투자분배금·투자수익만 비경상, 나머지는 전부 경상.
    inc_row = rows.get(_norm_label("B. 수입"))
    inc_pair = None
    if inc_row and _norm_label(ws.cell(inc_row + 1, 3).value or
                              ws.cell(inc_row + 1, 4).value or
                              ws.cell(inc_row + 1, 5).value) == "경상":
        inc_pair = (inc_row + 1, inc_row + 2)

    def income_split(result):
        if result is None:
            return 0.0, 0.0
        extra = 0.0
        for key, txns in result.rows_by_key.items():
            if not key.startswith("income."):
                continue
            for t in txns:
                if (t.tag or "").strip() in EXTRAORDINARY_INCOME_ACCOUNTS:
                    extra += t.amount / MILLION
        total = result.values.get("income.total", 0.0)
        return total - extra, extra

    incomes = {actual_month: income_split(actual_result)}
    if plan_result is not None and plan_month:
        incomes[plan_month] = income_split(plan_result)

    _mark_actual_header(ws, actual_month, warnings)

    # A. 기초잔고 — 실적월 칸을 우리가 아는 값(전월말 실제 잔고)으로 맞춥니다.
    # 이 시트의 1~6월은 템플릿에 쌓여 온 값이라, 그 누적 결과와 일일자금계획에서
    # 읽은 실제 기초가 미세하게 달랐습니다(0.009~0.17백만). 그대로 두면
    # 이후 달 기말이 요약1과 계속 어긋납니다. 계획월 칸은 원본 수식(전월 기말)을 그대로 둡니다.
    if open_total is not None:
        base_row = rows.get(_norm_label("A. 기초잔고"))
        col_a = ANNUAL_MONTH_COL.get(actual_month)
        if base_row and col_a:
            ws[f"{col_a}{base_row}"] = open_total

    for month, totals in jobs:
        col = ANNUAL_MONTH_COL.get(month)
        if not col:
            continue
        if inc_pair and month in incomes:
            ordinary, extra = incomes[month]
            ws[f"{col}{inc_pair[0]}"] = ordinary
            ws[f"{col}{inc_pair[1]}"] = extra
        for grp, val in totals.items():
            row = rows.get(_norm_label(grp))
            if row is None:
                if abs(val) > 1e-12:
                    warnings.append(
                        f"'{grp}' 라벨을 자금실적 시트에서 못 찾아 {month}월 값 "
                        f"{val:,.3f}백만을 넣지 못했습니다.")
                continue
            ws[f"{col}{row}"] = val


def _fill_annual_sheet(wb, result: Result, month: int, warnings,
                       opening: tuple[float, float] | None = None):
    """연간 시트의 해당 월 컬럼(원화/외화)에 leaf 값을 채웁니다.
    소계·기초잔고 이월 행은 원본 수식이 알아서 재계산하도록 건드리지 않습니다."""
    if SRC_ANNUAL not in wb.sheetnames:
        warnings.append(f"'{SRC_ANNUAL}' 시트가 없어 연간 반영을 건너뜁니다.")
        return
    ws = wb[SRC_ANNUAL]
    base = ANNUAL_SHEET_COL.get(month)
    if not base:
        return
    krw_col = base
    usd_col = get_column_letter(openpyxl.utils.column_index_from_string(base) + 1)
    if opening is not None:
        # A. 기초잔고 — 자금실적 시트와 같은 이유로 실적월 칸을 맞춥니다(§16 참고)
        base = _label_rows(ws, 1, 12, cols=(1, 2, 3, 4, 5)).get(_norm_label("A. 기초잔고"))
        if base:
            ws[f"{krw_col}{base}"] = opening[0]
            ws[f"{usd_col}{base}"] = opening[1]

    from core.engine import SUB_ETC32
    for key, row in ANNUAL_LEAF_ROW.items():
        if key == "expense.etc32":       # 소계 — leaf 합으로 계산
            krw = sum(result.krw.get(k, 0.0) for k in SUB_ETC32)
            usd = sum(result.usd.get(k, 0.0) for k in SUB_ETC32)
        else:
            krw, usd = result.krw.get(key, 0.0), result.usd.get(key, 0.0)
        ws[f"{krw_col}{row}"] = krw
        ws[f"{usd_col}{row}"] = usd


# 템플릿에 박혀 있는 옛 달 (6월=실적, 7월=계획)
OLD_ACTUAL_MONTH, OLD_PLAN_MONTH = 6, 7
_MONTH_RE = re.compile(rf"(?<!\d)({OLD_ACTUAL_MONTH}|{OLD_PLAN_MONTH})월")


def _fix_month_labels(wb, actual_month, plan_month):
    """표 헤더에 옛날 달 숫자가 박혀 있는 것을 이번 달로 바꿉니다 (v0.8).

    ※ replace()를 두 번 이어 쓰면 '6월→7월→8월'로 연쇄 치환됩니다.
       (실적 7월·계획 8월일 때 제목이 전부 8월이 되던 버그)
       반드시 한 번의 정규식 치환으로 동시에 바꿉니다.
    """
    table = {str(OLD_ACTUAL_MONTH): actual_month, str(OLD_PLAN_MONTH): plan_month}
    for sheet in ("1-1", "1-2", "2-1", "2-2", SRC_3P):
        if sheet not in wb.sheetnames:
            continue
        for row in wb[sheet].iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and "월" in cell.value \
                        and not cell.value.startswith("="):
                    cell.value = _MONTH_RE.sub(
                        lambda m: f"{table[m.group(1)]}월", cell.value)
