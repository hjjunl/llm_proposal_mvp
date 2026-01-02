# 📱 Proposal AI Agent - Streamlit 앱 구조

## 🏗️ 전체 구조 개요

```
proposal_ai_agent/
│
├── app.py                          # 🎯 메인 Streamlit 앱 (단일 페이지, 탭 구성)
│
├── pages/                          # 📄 멀티페이지 구조 (현재는 app.py에서 탭으로 처리)
│   ├── 01_RFP_Upload.py           # RFP 업로드 페이지 (별도 실행 가능)
│   └── 02_Client_History.py        # 고객 히스토리 페이지 (비어있음)
│
├── pipeline/                       # 🔄 핵심 비즈니스 로직
│   ├── __init__.py
│   ├── rfp2proposal.py            # RFP → 제안서 변환 파이프라인 (메인)
│   ├── analyze_rfp.py             # RFP 분석 모듈
│   └── inputs2flows.py            # 입력 → 플로우 변환
│
├── utils/                          # 🛠️ 유틸리티 함수
│   ├── db.py                      # SQLite 데이터베이스 관리
│   ├── paths.py                   # 경로 관리
│   └── io_utils.py                # 파일 I/O 유틸리티
│
├── .streamlit/                     # ⚙️ Streamlit 설정
│   └── secrets.toml               # API 키 및 인증 정보
│
├── DB/                             # 💾 데이터 저장소
│   ├── clients.db                 # SQLite 데이터베이스
│   ├── RFP/                       # 업로드된 RFP 파일
│   ├── proposal_result/           # 분석 결과물
│   └── proposal/                  # 최종 제안서
│
└── requirements.txt                # 📦 패키지 의존성
```

---

## 🎯 app.py - 메인 Streamlit 앱 구조

### 앱 구성 방식
- **타입**: 단일 페이지 앱 (Single Page App)
- **레이아웃**: 탭(Tabs) 기반 구성
- **인증**: `.streamlit/secrets.toml` 기반 간단한 로그인 시스템

### 코드 구조

```python
# =====================================================================================
# 1. 임포트 및 초기화 (1-124행)
# =====================================================================================
- 표준 라이브러리 (os, json, sqlite3, pathlib 등)
- Streamlit
- PDF/PPT 생성 라이브러리 (reportlab, python-pptx)
- 파이프라인 모듈 (pipeline.rfp2proposal)
- 유틸리티 함수들

# =====================================================================================
# 2. 인증 시스템 (55-117행)
# =====================================================================================
- _get_users_from_secrets(): secrets.toml에서 사용자 정보 로드
- is_authed(): 인증 상태 확인
- login_form(): 로그인 폼 렌더링
- logout_button(): 로그아웃 버튼

# =====================================================================================
# 3. 전역 경로 및 DB 설정 (127-207행)
# =====================================================================================
- ROOT, DB_DIR, RFP_DIR, RESULT_DIR, PROPOSAL_DIR 경로 정의
- SQLite 데이터베이스 초기화 (init_db)
- 고객/프로젝트/RFP 파일 CRUD 함수들

# =====================================================================================
# 4. PDF/PPT/Excel 생성 함수 (872-1171행)
# =====================================================================================
- build_pdf(): DataFrame → PDF 변환
- build_ppt(): DataFrame → PPT 변환
- build_excel(): DataFrame → Excel 변환
- 한글 폰트 자동 탐지 및 등록

# =====================================================================================
# 5. UI 컴포넌트 함수 (246-493행)
# =====================================================================================
- render_excel_like_flow(): 옵션 선택 및 미리보기 공용 컴포넌트
- render_inline_preview(): 실시간 미리보기
- _show_project_full_result(): 프로젝트 결과물 표시

# =====================================================================================
# 6. Streamlit UI 메인 (1248-1728행)
# =====================================================================================
st.set_page_config(page_title="Proposal Builder", layout="wide")
init_db()
login_form()
logout_button()

st.title("제안 생성 · 미리보기 · 내보내기 (자동/엑셀)")

# 3개 탭 구성
tab1, tab2, tab3 = st.tabs([
    "📥 Excel 업로드", 
    "⚡ 자동 생성(LLM·RFP/방향성)", 
    "🗂️ 고객 히스토리"
])

# TAB 1: Excel 업로드 (1259-1383행)
with tab1:
    - 파일 업로드 (CSV/Excel)
    - 고객 정보 입력
    - 요청별 옵션 선택
    - 실시간 미리보기
    - PDF/Excel/PPT 내보내기

# TAB 2: 자동 생성 (1398-1670행)
with tab2:
    - 고객 선택/등록
    - 프로젝트 정보 입력 (Form)
    - RFP 파일 업로드
    - 자동 DF 생성 버튼
    - 진행 상태 표시 (진행바, 로그)
    - 옵션 선택 및 내보내기

# TAB 3: 고객 히스토리 (1674-1728행)
with tab3:
    - 고객 선택
    - 고객 정보 표시
    - 프로젝트 목록
    - 결과물 다운로드
    - 삭제 기능
```

---

## 📄 pages/ 디렉토리 구조

### 현재 상태
- `pages/` 디렉토리는 존재하지만, **현재 app.py에서 탭으로 모든 기능을 처리**
- Streamlit의 멀티페이지 기능을 활용하려면 각 파일을 독립적으로 실행 가능

### 01_RFP_Upload.py
```python
# 독립 실행 가능한 RFP 업로드 페이지
- 고객 선택/등록
- 프로젝트 정보 입력
- RFP 파일 업로드
- DB 저장 및 config.json 생성
```

### 02_Client_History.py
- 현재 비어있음 (빈 파일)

---

## 🔄 pipeline/ - 핵심 비즈니스 로직

### rfp2proposal.py (메인 파이프라인)

**역할**: RFP 텍스트를 제안서 DataFrame으로 변환

**주요 함수**:
```python
build_flows_from_user_inputs(
    rfp_path: str,
    client_name: str,
    user_direction: str,
    notes: str = "",
    model_main: str = "gpt-5",
    model_deck: str = "gpt-5",
    out_dir: Optional[str] = None,
    logf: Optional[Callable] = None
) -> pd.DataFrame
```

**실행 단계**:
1. `extract_text_from_file()`: RFP 텍스트 추출
2. `rfp_requirement_check()`: 요청사항/질문 생성 (OpenAI)
3. `generate_answer_dict()`: 외부 리서치 (Perplexity/SerpAPI)
4. `generate_updated_plans_per_question()`: 업데이트된 제안방안 생성 (OpenAI)
5. `_slim_from_updated_df()`: 슬라이드 흐름 설계 (OpenAI Responses)
6. DataFrame 반환 및 파일 저장

### analyze_rfp.py
- RFP 분석 관련 함수들 (현재 비어있을 수 있음)

### inputs2flows.py
- 입력 데이터를 플로우로 변환하는 함수들

---

## 🛠️ utils/ - 유틸리티 모듈

### db.py
**역할**: SQLite 데이터베이스 관리

**주요 함수**:
```python
- init_db(): 테이블 생성
- upsert_client(): 고객 정보 저장/업데이트
- create_project(): 프로젝트 생성
- attach_rfp(): RFP 파일 정보 저장
- fetch_client_names(): 고객 목록 조회
- list_projects(): 프로젝트 목록 조회
- list_rfp_files(): RFP 파일 목록 조회
```

**테이블 구조**:
- `clients`: 고객 정보
- `projects`: 프로젝트 정보
- `rfp_files`: RFP 파일 정보

### paths.py
**역할**: 프로젝트 내 경로 통합 관리

**정의된 경로**:
```python
ROOT = Path(__file__).resolve().parents[1]
DB_DIR = ROOT / "DB"
RFP_DIR = DB_DIR / "RFP"
RESULT_DIR = DB_DIR / "proposal_result"
PROPOSAL_DIR = DB_DIR / "proposal"
SQLITE_PATH = DB_DIR / "clients.db"
```

### io_utils.py
**역할**: 파일 I/O 유틸리티

**주요 함수**:
```python
- sanitize_filename(): 파일명 안전화
- timestamp_utc(): UTC 타임스탬프 생성
- ensure_parent(): 부모 디렉토리 생성
```

---

## ⚙️ .streamlit/ - Streamlit 설정

### secrets.toml
**역할**: API 키 및 인증 정보 저장

**구조 예시**:
```toml
# API 키
OPENAI_API_KEY = "sk-..."
PERPLEXITY_API_KEY = "pplx-..."

# 인증 설정
[auth]
enabled = true

[auth.users]
admin = "sha256:..."
viewer = "sha256:..."
```

**접근 방법**:
```python
# app.py에서
st.secrets.get("OPENAI_API_KEY")
st.secrets.get("auth", {}).get("enabled", False)
```

---

## 💾 DB/ - 데이터 저장소 구조

```
DB/
├── clients.db                      # SQLite 데이터베이스
│
├── RFP/                            # 업로드된 RFP 파일
│   ├── {project_id}/
│   │   └── {timestamp}__{filename}
│   └── ...
│
├── proposal_result/                # 분석 결과물
│   ├── {project_id}/
│   │   ├── config.json            # 프로젝트 설정
│   │   ├── auto_df_{project_id}.xlsx
│   │   ├── slim_master_slide_flows.json
│   │   └── slim_master_slide_flows.xlsx
│   └── ...
│
└── proposal/                       # 최종 제안서 (선택적)
    ├── json/
    ├── research_results/
    └── final_ppt/
```

---

## 🔀 데이터 흐름도

### 시나리오 1: Excel 업로드 (Tab 1)

```
사용자
  │
  ├─> Excel 파일 업로드
  │   └─> app.py (Tab 1)
  │       ├─> pd.read_excel()
  │       ├─> ensure_slim_schema()
  │       └─> compute_option_big_titles()
  │
  ├─> 옵션 선택
  │   └─> render_excel_like_flow()
  │       ├─> render_inline_preview()
  │       └─> selected_df 생성
  │
  └─> 내보내기
      ├─> build_pdf() → PDF 다운로드
      ├─> build_excel() → Excel 다운로드
      └─> build_ppt() → PPT 다운로드
```

### 시나리오 2: 자동 생성 (Tab 2)

```
사용자
  │
  ├─> 프로젝트 정보 입력
  │   └─> app.py (Tab 2)
  │       ├─> upsert_client() → utils/db.py
  │       ├─> create_project() → utils/db.py
  │       └─> RFP 파일 저장 → DB/RFP/
  │
  ├─> "자동 DF 생성" 버튼 클릭
  │   └─> build_flows_from_user_inputs()
  │       └─> pipeline/rfp2proposal.py
  │           ├─> extract_text_from_file()
  │           ├─> rfp_requirement_check() → OpenAI API
  │           ├─> generate_answer_dict() → Perplexity API
  │           ├─> generate_updated_plans_per_question() → OpenAI API
  │           └─> _slim_from_updated_df() → OpenAI Responses API
  │               └─> DataFrame 반환
  │
  ├─> 결과 저장
  │   └─> DB/proposal_result/{project_id}/
  │       ├─> auto_df_{project_id}.xlsx
  │       ├─> slim_master_slide_flows.json
  │       └─> slim_master_slide_flows.xlsx
  │
  └─> 옵션 선택 및 내보내기 (Tab 1과 동일)
```

### 시나리오 3: 고객 히스토리 (Tab 3)

```
사용자
  │
  ├─> 고객 선택
  │   └─> app.py (Tab 3)
  │       └─> fetch_client_id_by_name() → utils/db.py
  │
  ├─> 프로젝트 목록 조회
  │   └─> list_projects() → utils/db.py
  │
  └─> 결과물 다운로드
      └─> _show_project_full_result()
          ├─> DB/proposal_result/{project_id}/ 파일들
          └─> 다운로드 버튼 제공
```

---

## 🎨 UI 컴포넌트 구조

### 1. 사이드바 (Sidebar)
```python
with st.sidebar:
    - 로그인 폼 (login_form)
    - 로그아웃 버튼 (logout_button)
    - 내보내기 설정 (Tab 1에서만)
        - PDF 본문 글자 크기 슬라이더
        - PPT 본문 글자 크기 슬라이더
```

### 2. 메인 영역 (Main Area)

#### Tab 1: Excel 업로드
```
┌─────────────────────────────────────────┐
│ 제목: "제안 생성 · 미리보기 · 내보내기"    │
├─────────────────────────────────────────┤
│ [📥 Excel 업로드] [⚡ 자동 생성] [🗂️ 히스토리] │
├─────────────────────────────────────────┤
│ 1) 데이터 업로드                         │
│    [파일 업로드 버튼]                    │
│                                         │
│ 2) 고객 정보                            │
│    [고객사] [작성팀] [작성일]            │
│                                         │
│ 3) 요청별 옵션 선택                      │
│    [요청 ID 1]                          │
│    ○ 옵션 1  ○ 옵션 2  ○ 옵션 3        │
│    [미리보기 영역]                      │
│    ─────────────────────────────        │
│    [요청 ID 2]                          │
│    ...                                  │
│                                         │
│ 4) 내보내기                             │
│    [📄 PDF 생성] [📊 Excel 생성] [🖼️ PPT 생성] │
└─────────────────────────────────────────┘
```

#### Tab 2: 자동 생성
```
┌─────────────────────────────────────────┐
│ [고객 선택/등록] (Expander)              │
│   ○ 기존 고객 선택  ○ 신규 고객 등록     │
│                                         │
│ [프로젝트 정보 입력] (Form)              │
│   프로젝트 제목*                         │
│   고객 방향성*                           │
│   RFP 파일 업로드*                       │
│   추가 메모                              │
│   [등록하기] [입력칸 초기화]             │
│                                         │
│ [자동 DF 생성 섹션]                      │
│   [🚀 자동 DF 생성] 버튼                │
│   [진행 상태 박스]                       │
│   [진행바]                               │
│   [실시간 로그]                          │
│                                         │
│ [옵션 선택 및 내보내기]                  │
│   (Tab 1과 동일한 구조)                  │
└─────────────────────────────────────────┘
```

#### Tab 3: 고객 히스토리
```
┌─────────────────────────────────────────┐
│ [고객 선택] (Selectbox)                  │
│                                         │
│ [고객 정보]                              │
│   역량: ...                             │
│   인원 정보: ...                        │
│   이전 프로젝트: ...                     │
│                                         │
│ [프로젝트 목록]                          │
│   ▼ [#1] 프로젝트 제목 - NEW            │
│      방향성: ...                        │
│      [결과물 다운로드]                   │
│      [🗑 삭제]                           │
│                                         │
│   ▼ [#2] 프로젝트 제목 - NEW            │
│      ...                                │
└─────────────────────────────────────────┘
```

---

## 🔐 인증 시스템

### 구조
```python
# secrets.toml
[auth]
enabled = true  # false로 설정하면 인증 비활성화

[auth.users]
admin = "sha256:해시값"
viewer = "sha256:해시값"
```

### 동작 방식
1. **로그인 폼**: 사이드바에 표시
2. **세션 상태**: `st.session_state.auth_user`에 사용자 ID 저장
3. **인증 확인**: `is_authed()` 함수로 각 탭 접근 전 확인
4. **비활성화**: `enabled = false`로 설정하면 모든 사용자 접근 가능

### 비밀번호 해시 생성
```python
import hashlib
hashlib.sha256("비밀번호".encode()).hexdigest()
# 결과: "sha256:해시값"
```

---

## 📦 의존성 (requirements.txt)

### 핵심 라이브러리
```
streamlit==1.37.0          # Streamlit 프레임워크
pandas==2.3.2              # 데이터 처리
numpy==2.2.2               # 수치 연산
openai                     # OpenAI API (requirements.txt에 명시되지 않았지만 사용)
```

### 문서 생성
```
reportlab==4.2.5           # PDF 생성
python-pptx==1.0.2         # PPT 생성
XlsxWriter==3.2.5          # Excel 생성
openpyxl==3.1.5            # Excel 읽기
```

### 기타
```
rich==13.9.4               # Streamlit 의존성
Pillow==10.4.0             # 이미지 처리
lxml==6.0.1                # XML/HTML 파싱
requests==2.32.5           # HTTP 요청
```

---

## 🚀 실행 방법

### 1. 환경 설정
```bash
cd proposal_ai_agent
pip install -r requirements.txt
```

### 2. API 키 설정
`.streamlit/secrets.toml` 파일 생성:
```toml
OPENAI_API_KEY = "sk-..."
PERPLEXITY_API_KEY = "pplx-..."
```

### 3. 앱 실행
```bash
streamlit run app.py
```

또는

```bash
python -m streamlit run app.py
```

### 4. 접속
브라우저에서 `http://localhost:8501` 접속

---

## 🔄 Streamlit 세션 상태 관리

### 주요 세션 변수
```python
# 인증
st.session_state.auth_user          # 현재 로그인한 사용자

# 자동 생성 관련
st.session_state.last_project       # 최근 등록 프로젝트 메타
st.session_state.auto_df_payload    # 자동 생성 DF + 메타
st.session_state.AUTO_BUSY         # 자동 생성 진행 중 여부
st.session_state.autolog            # 실시간 로그 리스트
st.session_state.rfp_form_version  # 폼 버전 (입력칸 초기화용)

# 설정
st.session_state.pdf_body_size      # PDF 본문 글자 크기
st.session_state.ppt_body_size     # PPT 본문 글자 크기
```

---

## 📝 주요 특징

### 1. 단일 페이지 + 탭 구조
- Streamlit의 멀티페이지 기능 대신 탭으로 구성
- 모든 기능을 한 화면에서 접근 가능

### 2. 공용 컴포넌트 재사용
- `render_excel_like_flow()`: Tab 1과 Tab 2에서 공통 사용
- 코드 중복 최소화

### 3. 실시간 피드백
- 자동 생성 시 진행 상태 표시
- 실시간 로그 스트리밍
- 진행바 표시

### 4. 파일 기반 데이터 관리
- SQLite: 고객/프로젝트 메타데이터
- 파일 시스템: RFP 파일, 결과물 저장

### 5. 유연한 인증 시스템
- secrets.toml 기반 간단한 인증
- 비활성화 가능 (개발/테스트용)

---

## 🎯 향후 개선 방향

1. **멀티페이지 전환**: `pages/` 디렉토리 활용하여 진짜 멀티페이지 구조로 전환
2. **캐싱**: 동일 RFP 재실행 시 중간 결과 재사용
3. **에러 처리 강화**: 더 명확한 에러 메시지 및 복구 메커니즘
4. **테스트 코드**: 단위 테스트 및 통합 테스트 추가
5. **문서화**: API 문서 및 사용자 가이드 작성

---

**작성일**: 2025년 1월  
**버전**: 1.0

