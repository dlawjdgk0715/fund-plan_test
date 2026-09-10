"""테스트 전체 실행:  python tests/run_all.py"""
import subprocess
import sys
from pathlib import Path

here = Path(__file__).resolve().parent
fails = []
for f in ("test_parsers.py", "test_engine.py", "test_report.py", "test_ui.py"):
    print(f"\n{'='*50}\n{f}\n{'='*50}")
    r = subprocess.run([sys.executable, str(here / f)])
    if r.returncode:
        fails.append(f)

print("\n" + "=" * 50)
if fails:
    print("실패:", ", ".join(fails))
    sys.exit(1)
print("전부 통과했습니다.")
