# run_pipeline_once.py
# -*- coding: utf-8 -*-
"""
단발 실행 러너:
- RFP 파일 하나를 입력으로 받아 build_flows_from_user_inputs 를 1회 실행
- 진행 로그를 콘솔에 출력
- 결과 DataFrame을 엑셀로 저장(DB/proposal_result/manual_run_YYYYmmdd_HHMMSS/auto_df.xlsx)

사용법:
  python run_pipeline_once.py --rfp "DB/RFP/sample.docx" \
      --client "EY 컨설팅" \
      --direction "전환율 개선과 CRM 연동을 중점"

필요 환경변수(.env 권장):
  OPENAI_API_KEY=...
  # Perplexity 키는 아래 중 아무 이름이나:
  # PERPLEXITY_API_KEY 또는 PPLX_API_KEY 또는 PEPLEXITY_API_KEY
  SERP_API_KEY=...        # (Perplexity 미설정/실패시 폴백)
"""

from __future__ import annotations

import os
import sys
import argparse
import traceback
from pathlib import Path
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from datetime import datetime
DATE_SUFFIX = os.getenv("FILE_DATE_SUFFIX", datetime.now().strftime("%Y%m%d"))
# -----------------------------
# 1) .env 로드 + 키 점검 출력
# -----------------------------
load_dotenv(override=True)

openai_api_key      = os.getenv("OPENAI_API_KEY")
anthropic_api_key   = os.getenv("ANTHROPIC_API_KEY")
google_api_key      = os.getenv("GOOGLE_API_KEY")
serp_api_key        = os.getenv("SERP_API_KEY")
# Perplexity는 여러 키 이름을 허용 (오타 포함)
perplexity_api_key  = (
    os.getenv("PERPLEXITY_API_KEY")
    or os.getenv("PPLX_API_KEY")
    or os.getenv("PEPLEXITY_API_KEY")
)

def _mask(k: str | None, n: int = 8) -> str:
    return (k[:n] if k else "")

print("=== .env / 환경키 점검 ===")
print(f"- OPENAI_API_KEY         : {'SET (' + _mask(openai_api_key) + '...)' if openai_api_key else 'NOT SET'}")
print(f"- ANTHROPIC_API_KEY      : {'SET (' + _mask(anthropic_api_key,7) + '...)' if anthropic_api_key else 'NOT SET'}")
print(f"- GOOGLE_API_KEY         : {'SET (' + _mask(google_api_key) + '...)' if google_api_key else 'NOT SET'}")
print(f"- SERP_API_KEY           : {'SET (' + _mask(serp_api_key) + '...)' if serp_api_key else 'NOT SET'}")
print(f"- PERPLEXITY_API_KEY/PPLX/PEPLEXITY : {'SET (' + _mask(perplexity_api_key) + '...)' if perplexity_api_key else 'NOT SET'}")

# (선택) OpenAI 클라이언트 선언이 필요한 경우를 대비해 미리 인스턴스 생성만 해둠
try:
    from openai import OpenAI
    if openai_api_key:
        openai_client = OpenAI(api_key=openai_api_key)
        MODEL = "gpt-5"  # 사용할 기본 모델명(내부 파이프라인에서 다른 모델을 지정할 수 있음)
    else:
        openai_client = None
        MODEL = "gpt-4o"
except Exception:
    openai_client = None
    MODEL = "gpt-4o"

# -----------------------------
# 2) 프로젝트 내부 모듈 import
# -----------------------------
# (프로젝트 루트에서 실행한다고 가정: proposal_ai_agent/pipeline/rfp2proposal.py)
sys.path.append(str(Path(__file__).parent))
from pipeline.rfp2proposal import (  # type: ignore
    build_flows_from_user_inputs,
)

# -----------------------------
# 3) 유틸
# -----------------------------
def ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

# -----------------------------
# 4) 메인
# -----------------------------
def main():
    default_rfp = "DB/RFP/25년 삼성전자 MX 미국 직영 매장 PMO_입찰공고문_F.docx"
    default_client = "EY 컨설팅"
    default_direction = "CRM 연동 및 전환율 개선 중심"

    ap = argparse.ArgumentParser()
    ap.add_argument("--rfp", default=default_rfp, help="RFP 파일 경로")
    ap.add_argument("--client", default=default_client, help="고객명")
    ap.add_argument("--direction", default=default_direction, help="고객 방향성")
    ap.add_argument("--model", default="gpt-4o", help="OpenAI 메인 모델")
    args = ap.parse_args()

    rfp_path = Path(args.rfp)
    if not rfp_path.exists():
        print(f"[에러] RFP 파일을 찾을 수 없습니다: {rfp_path}")
        sys.exit(1)

    def logf(msg: str):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    try:
        print("\n=== 파이프라인 시작 ===")
        auto_df: pd.DataFrame = build_flows_from_user_inputs(
            rfp_path=str(rfp_path),   # 문자열로 전달
            client_name=args.client,
            user_direction=args.direction,
            notes="",
            model_main=args.model,    # 예: "gpt-4o"
            logf=logf,                # 내부 진행 로그 콜백
        )

        out_dir = Path("DB/proposal_result") / f"manual_run_{ts()}"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_xlsx = out_dir / f"auto_df_{DATE_SUFFIX}.xlsx"   # → auto_df_20250930.xlsx
        auto_df.to_excel(out_xlsx, index=False)

        print(f"\n[완료] 결과 저장: {out_xlsx.resolve()}")
        print(f"[미리보기] 행/열: {auto_df.shape}")
        print(auto_df.head(8).to_string(index=False))

    except Exception as e:
        print("\n[실패] 파이프라인 실행 중 오류가 발생했습니다.")
        print("에러:", e)
        print("---- Traceback ----")
        traceback.print_exc()
        sys.exit(2)

if __name__ == "__main__":
    main()
