"""환율 조회 — 한국은행 ECOS API (통계표 731Y001, 매매기준율).

회사 PC에서 처음 실행할 때 방화벽에 막히지 않는지 꼭 확인해주세요.
API 키는 환경변수 ECOS_API_KEY 또는 core/config.py 에 넣습니다.
"""
from __future__ import annotations

import json
import urllib.request
from datetime import date

from core.config import ECOS_API_KEY

ITEM_CODE = {"USD": "0000001", "JPY": "0000002", "EUR": "0000003", "CNY": "0000053"}
BASE = "https://ecos.bok.or.kr/api/StatisticSearch"


def fetch_rate(currency: str, on: date | None = None, api_key: str | None = None) -> float | None:
    """해당 통화의 매매기준율(원). 조회 실패 시 None."""
    key = api_key or ECOS_API_KEY
    if not key:
        return None
    code = ITEM_CODE.get(currency.upper())
    if not code:
        return None
    d = on or date.today()
    ymd = d.strftime("%Y%m%d")
    # 주말·공휴일 대비로 최근 10일 구간을 조회하고 가장 마지막 값을 씁니다.
    start = ymd[:6] + "01"
    url = f"{BASE}/{key}/json/kr/1/100/731Y001/D/{start}/{ymd}/{code}"
    try:
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read().decode("utf-8"))
    except Exception:
        return None
    rows = data.get("StatisticSearch", {}).get("row", [])
    if not rows:
        return None
    try:
        return float(rows[-1]["DATA_VALUE"])
    except (KeyError, ValueError, IndexError):
        return None


def connection_test(api_key: str | None = None) -> tuple[bool, str]:
    """1·3번 화면에서 '환율 연결 테스트' 버튼용."""
    key = api_key or ECOS_API_KEY
    if not key:
        return False, "API 키가 설정되어 있지 않습니다 (환경변수 ECOS_API_KEY)."
    v = fetch_rate("USD", api_key=key)
    if v is None:
        return False, "연결 실패 — 방화벽 차단 또는 키 오류일 수 있습니다."
    return True, f"연결 성공 — USD 매매기준율 {v:,.2f}원"
