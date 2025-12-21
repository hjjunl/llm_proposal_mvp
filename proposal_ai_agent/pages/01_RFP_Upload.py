# pages/01_RFP_Upload.py
import json
import streamlit as st
from pathlib import Path
from utils.db import (
    init_db, upsert_client, create_project, attach_rfp,
    fetch_client_names, fetch_client_id_by_name
)
from utils.paths import RFP_DIR, RESULT_DIR
from utils.io_utils import sanitize_filename, timestamp_utc, ensure_parent

init_db()
st.title("📄 RFP 업로드 & 고객 방향성 입력")

with st.expander("고객 선택/등록", expanded=True):
    mode = st.radio("고객 선택 방식", ["기존 고객 선택", "신규 고객 등록"], horizontal=True)

    if mode == "기존 고객 선택":
        names = ["-- 선택 --"] + fetch_client_names()
        client_name = st.selectbox("고객명", options=names, index=0)
        capability = headcount = past_projects = ""
        if client_name == "-- 선택 --":
            st.info("기존 고객을 선택하거나, 아래에서 신규 고객 등록으로 전환하세요.")
    else:
        client_name = st.text_input("고객명*", placeholder="예: LVMH P&C Korea")
        capability = st.text_area("고객 역량/특기")
        headcount = st.text_area("인원 정보")
        past_projects = st.text_area("전에 진행한 프로젝트")

with st.form("rfp_form"):
    col1, col2 = st.columns(2)
    with col1:
        project_title = st.text_input("프로젝트 제목*", placeholder="예: 2025 상반기 이커머스 고도화 제안")
        direction = st.text_area("고객 방향성/원하는 바*", height=180,
                                 placeholder="예: 전환율 +15%, CRM 연동, 보안 준수(ISO27001), 자동 리포팅…")
    with col2:
        rfp_file = st.file_uploader("RFP 파일 업로드 (PDF/DOCX/TXT)", type=["pdf", "docx", "txt"])
        notes = st.text_area("추가 메모(선택)")

    submitted = st.form_submit_button("등록하기")

if submitted:
    if not client_name or client_name == "-- 선택 --":
        st.error("고객명을 선택/입력해주세요.")
        st.stop()
    if not project_title or not direction:
        st.error("프로젝트 제목과 고객 방향성은 필수입니다.")
        st.stop()
    if not rfp_file:
        st.error("RFP 파일을 업로드해주세요.")
        st.stop()

    # 고객 저장/업데이트
    client_id = upsert_client(client_name, capability, headcount, past_projects)

    # 프로젝트 생성
    project_id = create_project(client_id, project_title, direction)

    # 파일 저장 경로 (타임스탬프 + 원본명 sanitize)
    ts = timestamp_utc()
    safe_name = sanitize_filename(rfp_file.name)
    proj_rfp_dir = RFP_DIR / str(project_id)
    stored_path = proj_rfp_dir / f"{ts}__{safe_name}"
    ensure_parent(stored_path)

    with open(stored_path, "wb") as f:
        f.write(rfp_file.getbuffer())

    attach_rfp(project_id, rfp_file.name, stored_path)

    # 분석 파이프라인용 config.json
    proj_out_dir = RESULT_DIR / str(project_id)
    ensure_parent(proj_out_dir / "config.json")
    config = {
        "project_id": project_id,
        "client_name": client_name,
        "project_title": project_title,
        "direction": direction,
        "rfp_path": str(stored_path),
        "notes": notes or "",
        "created_at": ts,
    }
    (proj_out_dir / "config.json").write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    st.success(f"등록 완료! 프로젝트 ID: {project_id}")
    st.markdown(f"- 원본 저장 위치: `{stored_path}`")
    st.markdown(f"- 분석 설정: `{proj_out_dir / 'config.json'}`")
    st.info("이제 파이프라인(주피터/CLI)에서 해당 config.json을 읽어 분석을 실행하세요.")
