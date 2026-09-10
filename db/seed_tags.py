"""태그 사전 초기 데이터 — 원본 워크북의 실제 계정과목 이름을 그대로 씁니다.

원본 `3p` 시트의 SUMIFS 수식이 참조하는 계정과목 이름이 곧 표준 태그입니다.
그래서 목록은 core/engine.py 의 LEAF 정의에서 자동으로 가져옵니다.
적요 예시(EXAMPLES)만 여기서 관리하면 됩니다.

수정 후:  python db/seed_tags.py --reset
"""
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core.engine import ACCOUNT_MAP, ALL_ACCOUNTS, ANNUAL_GROUP, EXCLUDED_ACCOUNTS  # noqa: E402

DB_PATH = Path(__file__).resolve().parent / "tags.db"
SCHEMA = Path(__file__).resolve().parent / "schema.sql"

# 적요 예시 — 계정 자동분류의 근거가 됩니다. 쓰면서 계속 쌓입니다.
EXAMPLES = {
    "외상매출금": ["게임 매출 정산", "플랫폼 정산금"],
    "세금환급": ["부가가치세 환급", "법인세 환급"],
    "이자수익": ["정기예금 이자", "보통예금 이자", "정기예금 만기 해지 이자"],
    "기타입금": ["잡수입"],
    "수입임대료": ["사무실 전대 임대료"],
    "급여": ["임직원 급여"],
    "상여": ["상여금"],
    "퇴직연금": ["퇴직연금 부담금"],
    "4대보험": ["국민연금", "건강보험", "고용보험", "산재보험"],
    "광고판촉비": ["광고선전비", "판촉비", "마케팅 집행"],
    "법인세": ["법인세 중간예납", "법인세 납부"],
    "부가가치세": ["1기 확정 부가가치세 납부", "부가세 예정신고 납부"],
    "원천세": ["근로소득 원천세", "사업소득 원천세"],
    "기타제세공과금": ["주민세", "인지세"],
    "지급이자": ["차입금 이자"],
    "S/W": ["소프트웨어 라이선스 구매"],
    "H/W": ["서버 장비 구매", "PC 구매"],
    "시설장치": ["인테리어 공사"],
    "차량": ["업무용 차량 구입"],
    "투자금": ["지분 투자", "조합 출자"],
    "지분투자": ["지분 취득"],
    "전산유지비": ["서버 이용료", "클라우드 사용료"],
    "전산관련지출": ["전산 유지보수", "SaaS 구독료"],
    "외주용역비": ["개발 외주 용역"],
    "경영지원용역비": ["그룹사 공용비용", "경영지원 용역비", "업무지원 용역비"],
    "복리후생비": ["주택자금대출 이자지원", "경조사비", "건강검진비"],
    "식당운영비": ["구내식당 운영비", "카페 운영비", "매점 운영비"],
    "직원비용": ["직원경비", "법인카드 대금", "복지포인트 사용비"],
    "지급임차료": ["사무실 임차료", "관리비"],
    "지급수수료": ["위탁매매 수수료", "법률자문 수수료", "세무자문 수수료"],
    "기타지출": ["통신비", "우편료", "잡비"],
    "금융상품": ["정기예금 해지", "정기예금 예치", "정기예금 재예치"],
    "외환매도": ["달러 매도"],
    "외환매입": ["달러 매입"],
    "차입금": ["차입 실행"],
    "차입금상환": ["차입금 상환"],
    "보통예금": ["계좌간 자금대체"],
}

# ※ '매달 재확인 필요 항목' 기본 규칙은 core/data_store.py 한 곳에만 둡니다.
#    (예전에는 여기서도 넣어서, 조건이 서로 다른 규칙이 두 벌 생겼습니다)


def seed(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    accounts = sorted(set(ALL_ACCOUNTS) | set(EXCLUDED_ACCOUNTS) | set(EXAMPLES))
    for acct in accounts:
        manual = 1 if acct in ("복리후생비", "직원비용") else 0
        exclude = 1 if acct in EXCLUDED_ACCOUNTS else 0
        examples = "\n".join(EXAMPLES.get(acct, []))
        cur.execute(
            """INSERT INTO account_tags (tag, needs_confirm, manual_input, exclude_agg, examples, io_class)
               VALUES (?,?,?,?,?,'')
               ON CONFLICT(tag) DO UPDATE SET manual_input=excluded.manual_input,
                 exclude_agg=excluded.exclude_agg,
                 examples=CASE WHEN account_tags.examples='' THEN excluded.examples
                               ELSE account_tags.examples END""",
            (acct, 0, manual, exclude, examples),
        )
    for (acct, direction), row_key in ACCOUNT_MAP.items():
        grp, cls = ANNUAL_GROUP.get(row_key, ("", ""))
        cur.execute(
            """INSERT INTO tag_row_map (tag, direction, currency, row_key, annual_group, annual_class, sort_order)
               VALUES (?,?,'*',?,?,?,0)
               ON CONFLICT(tag, direction, currency) DO UPDATE SET
                 row_key=excluded.row_key, annual_group=excluded.annual_group,
                 annual_class=excluded.annual_class""",
            (acct, direction, row_key, grp, cls),
        )
    conn.commit()


if __name__ == "__main__":
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA.read_text(encoding="utf-8"))
    seed(conn)
    n = conn.execute("SELECT COUNT(*) FROM account_tags").fetchone()[0]
    m = conn.execute("SELECT COUNT(*) FROM tag_row_map").fetchone()[0]
    print(f"계정과목 {n}개, 행매핑 {m}개 저장 완료 → {DB_PATH}")
    conn.close()
