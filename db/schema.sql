-- 자금운용계획 자동화 웹앱 — 전체 스키마
-- 이 파일은 앱이 처음 실행될 때 자동으로 실행됩니다(core/data_store.py의 init_db).

-- ── M1: 태그 사전 ────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS account_tags (
    tag           TEXT PRIMARY KEY,   -- 표준 계정과목 이름(34개)
    needs_confirm INTEGER DEFAULT 0,  -- 1이면 자동분류돼도 사람이 반드시 확인
    manual_input  INTEGER DEFAULT 0,  -- 1이면 취합파일에 안 나타나 매달 직접 입력해야 하는 항목
    exclude_agg   INTEGER DEFAULT 0,  -- 1이면 집계에서 제외
    examples      TEXT DEFAULT '',    -- 적요 예시(줄바꿈 구분) — 유사도 비교용
    io_class      TEXT DEFAULT ''     -- 수입 태그의 경상/비경상 구분: 'ordinary' | 'extraordinary'
);

CREATE TABLE IF NOT EXISTS tag_row_map (
    tag          TEXT NOT NULL,
    direction    TEXT NOT NULL,       -- 'in'(입금) | 'out'(출금)
    currency     TEXT NOT NULL,       -- 'KRW' | 'USD' | '*'(통화 무관)
    row_key      TEXT NOT NULL,       -- 유동자금세부 표의 행 식별자
    annual_group TEXT DEFAULT '',     -- 자금실적(당해) 시트의 세부그룹 이름
    annual_class TEXT DEFAULT '',     -- '경상' | '비경상'
    sort_order   INTEGER DEFAULT 0,
    PRIMARY KEY (tag, direction, currency)
);

-- ── M2/M4: 거래 내역 ──────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS transactions (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    period      TEXT NOT NULL,        -- 'YYYY-MM'
    kind        TEXT NOT NULL,        -- 'actual'(실적) | 'plan'(계획)
    dept        TEXT DEFAULT '',      -- 부서명(계획 파일)
    txn_date    TEXT DEFAULT '',
    counterpart TEXT DEFAULT '',      -- 업체명
    memo        TEXT DEFAULT '',      -- 적요/사유
    tag         TEXT DEFAULT '',      -- 계정과목(태그)
    currency    TEXT DEFAULT 'KRW',
    amount_in   REAL DEFAULT 0,       -- 입금액(원 단위)
    amount_out  REAL DEFAULT 0,       -- 출금액(원 단위)
    source      TEXT DEFAULT '',      -- 'upload' | 'manual'
    excluded    INTEGER DEFAULT 0,    -- 1이면 중복 등으로 제외 처리
    created_at  TEXT DEFAULT (datetime('now','localtime'))
);
CREATE INDEX IF NOT EXISTS idx_txn_period ON transactions(period, kind);

-- ── M2: 부서별 열 매핑 기억 ──────────────────────────────────────
CREATE TABLE IF NOT EXISTS column_mappings (
    dept        TEXT PRIMARY KEY,
    mapping_json TEXT NOT NULL,       -- {"memo": 2, "amount": 5, ...} 열 번호
    signature   TEXT DEFAULT '',      -- 감지된 열 배치 지문(같으면 자동 재사용)
    updated_at  TEXT DEFAULT (datetime('now','localtime'))
);

-- ── M6/v0.11: 비고 라벨 축약표 ───────────────────────────────────
CREATE TABLE IF NOT EXISTS remark_labels (
    tag   TEXT NOT NULL,
    memo  TEXT NOT NULL,
    label TEXT NOT NULL,
    PRIMARY KEY (tag, memo)
);

-- ── §11: 운용자금 잔고 이력 ──────────────────────────────────────
CREATE TABLE IF NOT EXISTS fund_balances (
    period          TEXT PRIMARY KEY, -- 'YYYY-MM'
    invest_end_krw  REAL DEFAULT 0,   -- 기말 운용자금 원화(백만원)
    invest_end_usd  REAL DEFAULT 0,   -- 기말 운용자금 외화(천$)
    source          TEXT DEFAULT 'auto',
    updated_at      TEXT DEFAULT (datetime('now','localtime'))
);

-- ── M6: 매달 재확인 필요 항목 ────────────────────────────────────
CREATE TABLE IF NOT EXISTS recheck_items (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    tag         TEXT NOT NULL,        -- 계정과목 조건 (빈칸이면 계정과목은 안 봄)
    pattern     TEXT NOT NULL,        -- 적요에서 찾을 키워드
    label       TEXT NOT NULL,        -- 화면에 보여줄 이름
    counterpart TEXT DEFAULT '',      -- 거래처 조건 (빈칸이면 거래처는 안 봄)
    hint        TEXT DEFAULT '',      -- 수기 입력 안내 문구
    def_tag     TEXT DEFAULT '',      -- 수기 입력할 때 미리 채울 계정과목
    def_dir     TEXT DEFAULT 'out',   -- 'in'(입금) | 'out'(출금)
    active      INTEGER DEFAULT 1,
    last_seen   TEXT DEFAULT ''
);

-- ── 월별 유동/운용 잔고 (일일자금계획 「3. 자금현황」에서 자동으로 읽음) ──
CREATE TABLE IF NOT EXISTS month_balances (
    period        TEXT PRIMARY KEY,  -- 'YYYY-MM' (실적 월)
    open_krw      REAL DEFAULT 0,    -- 전월말 원화 유동잔고(백만원) = 이번 달 기초
    open_usd      REAL DEFAULT 0,    -- 전월말 외화(천$)
    close_krw     REAL DEFAULT 0,    -- 당월말 원화 유동잔고(백만원) = 다음 달 기초
    close_usd     REAL DEFAULT 0,
    invest_krw    REAL DEFAULT 0,    -- 당월말 운용자금(백만원)
    invest_usd    REAL DEFAULT 0,
    open_invest_krw REAL DEFAULT 0,  -- 전월말 운용자금(백만원) = 이번 달 기초 (0630-1 시트에서 읽음)
    open_invest_usd REAL DEFAULT 0,
    open_sheet    TEXT DEFAULT '',   -- 어느 시트에서 읽었는지 (예: 0630-1)
    close_sheet   TEXT DEFAULT '',
    updated_at    TEXT DEFAULT (datetime('now','localtime'))
);


-- ── 보고서 표의 '비고' 칸 (사람이 직접 쓰는 문장) ──────────────────
CREATE TABLE IF NOT EXISTS report_remarks (
    period    TEXT NOT NULL,
    table_key TEXT NOT NULL,   -- cash_flow / liquid_plan / fund_plan / stock
    row_idx   INTEGER NOT NULL,
    text      TEXT DEFAULT '',
    PRIMARY KEY (period, table_key, row_idx)
);

-- ── 4번 '최종 점검' 화면: 검증 게이트 확인 기록 ────────────────────
-- 사용자가 [확인함]을 누른 시점의 거래 내역 지문을 함께 저장합니다.
-- 나중에 거래를 고치면 지문이 달라져 확인이 자동으로 풀립니다(다시 빨간색).
CREATE TABLE IF NOT EXISTS gate_acks (
    period      TEXT NOT NULL,
    scope       TEXT NOT NULL,     -- 'actual'(실적) | 'plan'(계획)
    gate_id     TEXT NOT NULL,     -- V1 ~ V6
    fingerprint TEXT NOT NULL,     -- 확인 당시 거래 내역 지문(sha1)
    updated_at  TEXT DEFAULT (datetime('now','localtime')),
    PRIMARY KEY (period, scope, gate_id)
);

-- ── 비고 '쓰는 방식'(서식) 학습 ────────────────────────────────────
-- 문장 내용이 아니라 금액 단위·소수 자릿수·항목 개수·구분 기호 같은 '틀'만 기억합니다.
CREATE TABLE IF NOT EXISTS remark_styles (
    table_key  TEXT PRIMARY KEY,   -- cash_flow / detail / summary1
    style_json TEXT NOT NULL,
    samples    INTEGER DEFAULT 0,
    updated_at TEXT DEFAULT (datetime('now','localtime'))
);
