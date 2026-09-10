"""파서 검증 — 외부 파일 없이 돌아갑니다.

실제 파일에서 겪었던 문제들이 다시 생기지 않는지 확인합니다.
실행:  python tests/test_parsers.py
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.parsers import (SECTION_ACTUAL, SECTION_PLAN, detect_plan_columns, guess_dept,
                          parse_daily_sheet, parse_plan_rows, read_cash_status,
                          sheet_matches, sheet_period)

ok = 0


def check(cond, label):
    global ok
    assert cond, f"✗ {label}"
    ok += 1
    print(f"✓ {label}")


def grid(rows_spec, width=18):
    rows = [[None] * width for _ in range(max(rows_spec) + 2)]
    for r, cells in rows_spec.items():
        for c, v in cells.items():
            rows[r][c] = v
    return rows


# ── 1. 일일자금계획: 실적/계획 표 분리, 숨김 A열, 소계 경계 ──────
daily = grid({
    5:  {1: "1. 일일자금수지 실적"},
    6:  {0: "계정과목", 1: "통화구분", 2: "거래처명", 3: "내 역", 5: "수입", 6: "지출"},
    7:  {0: "기타지출", 1: "1)원화", 2: "하나은행", 3: "잔액증명서발급수수료", 6: 1000},
    8:  {0: "원천세", 2: "금천세무서", 3: "근로소득세 원천징수", 6: 3563860},
    9:  {0: "기타입금", 2: "㈜비즈마켓", 3: "복지포인트 환급", 5: 78020},
    10: {4: "소계(₩)", 5: 78020, 6: 3564860},
    11: {1: "2) 외화(USD)"},
    12: {0: "전산유지비", 2: "AWS", 3: "클라우드 사용료", 6: 1500},
    13: {4: "소계(U$)", 6: 1500},
    15: {1: "2. 일일자금수지 계획"},
    16: {0: "계정과목", 1: "통화구분", 2: "거래처명", 3: "내 역", 5: "수입", 6: "지출"},
    17: {0: "퇴직연금", 1: "1) 원화", 2: "기업은행", 3: "임직원 퇴직연금", 6: 1466670},
    18: {4: "소계(₩)", 6: 1466670},
    # 표 밖 — 읽으면 안 되는 자금대체 영역
    20: {1: "3. 자금대체"},
    21: {2: "332-910052-92404", 3: "MMDA", 6: 827057},
})
a = parse_daily_sheet(daily, "2026-07", "0713", SECTION_ACTUAL)
p = parse_daily_sheet(daily, "2026-07", "0713", SECTION_PLAN)
check(len(a) == 4, f"일일자금계획 실적 표 4건 (실제 {len(a)})")
check(len(p) == 1, f"일일자금계획 계획 표 1건 (실제 {len(p)})")
check(a[0].tag == "기타지출", "숨김 A열의 계정과목을 읽음")
check(a[3].currency == "USD", "외화 구간을 USD로 인식")
check(not any("MMDA" in t.memo for t in a + p), "소계 아래 자금대체 표는 안 읽음")

# ── 2. 부서 계획서: '합계'(부가세 포함) 사용, 입/출금 구분 ────────
plan = [
    ["2026년 8월 자금 집행 계획서", None, None, None, None, None, None, None],
    ["부서명 : 내부회계 관리팀", None, None, None, None, None, None, None],
    ["업체명", "입/출금 구분", "사유", "통화", "금액", "부가가치세", "합계", "비고"],
    ["㈜컴투스", "출금", "인사 용역료", "KRW", 39718, 3971, 43689, "20일"],
    ["컴투스플랫폼", "입금", "클라우드 분담금", "KRW", 51208, 5120, 56328, "말일"],
    ["합 계", None, None, None, 90926, 9091, 100017, None],
]
hr, mp = detect_plan_columns(plan)
t = parse_plan_rows(plan, mp, hr, "2026-07", "회계팀")
check(mp.get("amount_total") == 6, "'합계' 열을 집계 금액으로 잡음")
check(len(t) == 2, f"합계 행을 거래로 세지 않음 (실제 {len(t)}건)")
check(t[0].amount_out == 43689, "출금 금액 = 합계(부가세 포함)")
check(t[1].amount_in == 56328, "'입금' 구분을 방향으로 반영")

# ── 3. 부서명 자동 추출 ──────────────────────────────────────────
check(guess_dept("내부회계관리팀_2026년8월_자금집행계획서.xlsx") == "내부회계관리팀",
      "파일명에서 부서명 추출")
check(guess_dept("2026.08 인프라실 자금집행계획.xlsx") == "인프라실",
      "날짜·공통단어를 걸러내고 부서명 추출")

# ── 4. 시트 이름 인식 (부서마다 표기가 다름) ─────────────────────
for name in ["2026년 8월", "2026-08", "26.08", "202608", "2608", "8월", "8월계획",
             "계획(8월)", "Aug", "August 2026"]:
    assert sheet_matches(name, (2026, 8)), f"✗ 시트 이름 인식: {name}"
check(True, "시트 이름 10가지 형태 인식")
check(not sheet_matches("2025년 8월", (2026, 8)), "다른 연도는 걸러냄")
check(sheet_period("표지") is None, "표지 같은 이름은 월로 안 봄")

# ── 5. 「3. 자금현황」 잔고 읽기 (행 위치가 달라도 동작) ──────────
def cash_grid(offset):
    spec = {
        0 + offset: {1: "3. 자금현황"},
        1 + offset: {1: "1) 원화 수시입출 예금"},
        2 + offset: {1: "은행", 2: "계좌번호", 4: "전일 잔액", 7: "당일잔액"},
        3 + offset: {1: "하나은행", 4: 115935, 7: 115935},
        4 + offset: {1: "기업은행", 4: 23643752, 7: 23643752},
        5 + offset: {3: "합계", 4: 23759687, 7: 23759687},
        7 + offset: {1: "2) 외화(USD) 수시입출 예금"},
        8 + offset: {1: "은행", 4: "전일 잔액", 7: "당일잔액"},
        9 + offset: {1: "하나은행", 7: 1234},
        10 + offset: {3: "합계", 7: 1234},
        12 + offset: {1: "3) 운용자금"},
        13 + offset: {1: "은행", 4: "전일 잔액", 7: "당일잔액"},
        14 + offset: {1: "하나은행", 7: 2100000000},
        15 + offset: {3: "합계", 7: 2100000000},
    }
    return grid(spec)


for off in (0, 1, 5):
    cs = read_cash_status(cash_grid(off))
    assert cs["krw"] == 23759687 and cs["usd"] == 1234 and cs["invest"] == 2100000000, \
        f"✗ 자금현황 읽기 (offset {off}): {cs}"
check(True, "「3. 자금현황」 — 행 위치가 밀려도 원화·외화·운용자금 정확히 읽음")

print(f"\n{ok}개 항목 전부 통과했습니다.")
