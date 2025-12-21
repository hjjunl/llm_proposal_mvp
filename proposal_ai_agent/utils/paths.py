# utils/paths.py
from pathlib import Path

# 프로젝트 루트 기준 경로
ROOT = Path(__file__).resolve().parents[1]
DB_DIR = ROOT / "DB"

RFP_DIR = DB_DIR / "RFP"
RESULT_DIR = DB_DIR / "proposal_result"
PROPOSAL_DIR = DB_DIR / "proposal"
SQLITE_PATH = DB_DIR / "clients.db"

# 필요한 디렉터리 미리 생성
for p in [DB_DIR, RFP_DIR, RESULT_DIR, PROPOSAL_DIR]:
    p.mkdir(parents=True, exist_ok=True)
