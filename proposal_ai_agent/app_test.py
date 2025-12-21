# app.py
import streamlit as st
from utils.db import init_db
from utils.paths import DB_DIR, RFP_DIR, RESULT_DIR, PROPOSAL_DIR

st.set_page_config(page_title="RFP → Proposal MVP", page_icon="🧠", layout="wide")

# DB 초기화 (최초 1회 실행)
init_db()

st.title("🧠 RFP → Proposal MVP")
st.markdown("""
이 앱은 **RFP 업로드 → 고객 방향성 입력 → 분석 파이프라인 연계**를 위한 최소기능 버전입니다.

**다음 페이지로 이동해 시작하세요:**
- 📄 **RFP 업로드 & 방향성 입력** (사이드바 Pages)
- 🗂️ **고객/프로젝트 히스토리**
""")

st.divider()
st.subheader("현재 경로")
st.code(f"""
DB_DIR        = {DB_DIR}
RFP_DIR       = {RFP_DIR}
RESULT_DIR    = {RESULT_DIR}
PROPOSAL_DIR  = {PROPOSAL_DIR}
""", language="text")
