# pipeline/inputs2flows.py
from __future__ import annotations
import os, json, re, time
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional

import pandas as pd

# ---------- 외부 의존 ----------
# - openai:  pip install openai
# - pypdf:   pip install pypdf
# - python-docx: pip install python-docx
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

# ------------------------------------------------------------------------------------
# 0) 공통 유틸: 파일에서 RFP 텍스트 추출
# ------------------------------------------------------------------------------------
def _read_txt(p: Path) -> str:
    return p.read_text(encoding="utf-8", errors="ignore")

def _read_pdf(p: Path) -> str:
    from pypdf import PdfReader
    r = PdfReader(str(p))
    return "\n".join([(pg.extract_text() or "") for pg in r.pages])

def _read_docx(p: Path) -> str:
    import docx
    d = docx.Document(str(p))
    return "\n".join([para.text for para in d.paragraphs])

def extract_rfp_text(rfp_path: str | Path) -> str:
    p = Path(rfp_path)
    suf = p.suffix.lower()
    if suf == ".pdf":
        return _read_pdf(p)
    if suf == ".docx":
        return _read_docx(p)
    if suf in [".txt", ".md"]:
        return _read_txt(p)
    raise ValueError(f"지원하지 않는 파일 형식: {suf}")

# ------------------------------------------------------------------------------------
# 1) 네가 제공한 기존 로직 (LLM 사용) — 최소 수정/래핑
# ------------------------------------------------------------------------------------
def _get_openai_client():
    if OpenAI is None:
        raise RuntimeError("openai 패키지가 필요합니다. `pip install openai`")
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("환경변수 OPENAI_API_KEY가 필요합니다.")
    return OpenAI(api_key=api_key)

def _chat(model: str, messages: List[Dict[str, str]], temperature: float = 0.1) -> str:
    cli = _get_openai_client()
    resp = cli.chat.completions.create(model=model, messages=messages, temperature=temperature)
    return resp.choices[0].message.content.strip()

# ---- (1) rfp_requirement_check ------------------------------------------------------
def rfp_requirement_check(user_inputs: dict, model: str = "gpt-4o", temperature: float = 0.1):
    rfp_text = user_inputs.get("rfp_text")
    style = user_inputs.get("style", "신뢰감 있는")
    keywords = user_inputs.get("keywords", [])
    client_name = user_inputs.get("client_name", "고객사명 미입력")
    user_direction = user_inputs.get("user_direction", "")

    if not rfp_text or len(rfp_text.strip()) < 30:
        raise ValueError("❗ RFP 원문이 비어 있거나 너무 짧습니다. 실제 RFP를 반드시 입력하세요.")

    system_prompt = f"""
    당신은 'EY·맥킨지 등 국내외 최상위 전략 컨설팅 회사의 파트너급 제안서 전문가 AI'입니다.
    <요청사항>
    - 제안 요청 사항을 파악하고, 각 요청 사항에 대해 상세한 제안 방안 제시해주세요.
    - 최대한 많이 제안 요청 사항을 세부적으로 나누어 주세요.
    - 제안 방안에는 RFP내 있는 세부적인 내용을 내포 할수 있도록 해주세요.
    - 마크다운으로 작성하지 마세요
    - 1. RFP에 있는 내용들을 관련이 있다면 제안 요청사항에 세부 항목으로 추가해서 정리하고 진행할 업무에 대해 각각을 제안한다고 생각하고 정리해주세요.
    - 2. 제안 방안을 뒷받침할 질문으로 만들어 주세요 추후 perplexity를 통해 리서치 후 벤치마킹을 하여 고객에게 제공할 예정입니다.
    - 3. 다른 제안을 하기위한 벤치마킹 질문을 만들어주세요.
    - 제안 방안을 실제로 실현 할 수 있기 위해 물어봐야할 질문 리스트를 만들어 주세요 최대한 (상세히). []
    - 질문의 내용은 최대한 상세하게 실제 제안에 유의미하게 작성해주세요.
    안좋은 예시: 프로젝트 관리에서 가장 중요하게 고려해야 할 요소는 무엇인가요?
    좋은 예시: 전자 상품을 파는 리테일 회사의 직영 매장 프로젝트 관리에서 가장 중요하게 고려해야 할 요소는 무엇인가요?
    """.strip()

    user_prompt = f"""
    [RFP 원문]
    {rfp_text}

    [고객 방향성/강조]
    {user_direction or '없음'}

    [강조 키워드]: {', '.join(keywords) if keywords else '없음'}

    위 정보를 바탕으로 제안 요청사항 및 제안 방안을 상세히 정리해주세요.
    """.strip()

    txt = _chat(model, [{"role":"system","content":system_prompt},{"role":"user","content":user_prompt}], temperature)
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    return lines

# ---- (2) 파싱: lines → dict --------------------------------------------------------
def parse_rfp_response_to_dict(response_lines: List[str]):
    results = {}
    current_section = {}
    current_key = ""
    section_index = 0
    for line in response_lines:
        line = line.strip()
        if line.startswith(tuple(str(i) + '.' for i in range(1, 100))):
            if current_section:
                results[f"요청사항_{section_index}"] = current_section
            section_index += 1
            current_section = {
                "제안요청사항": line.split(":", 1)[-1].strip(),
                "제안방안": "",
                "실현질문": [],
                "벤치마킹질문": []
            }
            current_key = ""
        elif line.startswith("- 제안 방안:"):
            current_section["제안방안"] = line.split(":", 1)[-1].strip()
        elif "실현 할 수 있기 위한 질문" in line:
            current_key = "실현질문"
        elif "벤치마킹" in line:
            current_key = "벤치마킹질문"
        elif line.startswith("(") and current_key:
            current_section[current_key].append(line)
    if current_section:
        results[f"요청사항_{section_index}"] = current_section
    return results

# ---- (3) 질문 컨텍스트 붙이기 ------------------------------------------------------
def make_contextual_question(req_title: str, question: str) -> str:
    req_title = (req_title or "").strip()
    question = (question or "").strip()
    if not req_title:
        return question
    return f"{req_title} 관점에서, {question}"

# ---- (4) DataFrame 변환 (질문/URL 포함) -------------------------------------------
def rfp_to_dataframe_complete(rfp_dict, answer_dict):
    rows = []
    for req_id, req in rfp_dict.items():
        req_title = req.get("제안요청사항", "제안요청사항 없음")
        # 실현질문
        SJ = req.get("실현질문", [])
        SA = (answer_dict.get(req_id, {}) or {}).get("실현질문_답변", [None]*len(SJ))
        SU = (answer_dict.get(req_id, {}) or {}).get("실현질문_url", [None]*len(SJ))
        for q, a, u in zip(SJ, SA, SU):
            rows.append({
                "요청 ID": req_id, "제안요청사항": req_title, "질문유형": "실현질문",
                "질문": make_contextual_question(req_title, q), "답변": a, "urls": u
            })
        # 벤치마킹질문
        BJ = req.get("벤치마킹질문", [])
        BA = (answer_dict.get(req_id, {}) or {}).get("벤치마킹질문_답변", [None]*len(BJ))
        BU = (answer_dict.get(req_id, {}) or {}).get("벤치마킹질문_url", [None]*len(BJ))
        for q, a, u in zip(BJ, BA, BU):
            rows.append({
                "요청 ID": req_id, "제안요청사항": req_title, "질문유형": "벤치마킹질문",
                "질문": make_contextual_question(req_title, q), "답변": a, "urls": u
            })
    return pd.DataFrame(rows)

# ---- (5) Q/A → 업데이트된 제안 방안 생성 -------------------------------------------
def _call_llm_simple(prompt: str, model: str = "gpt-4o", temperature: float = 0.3) -> str:
    cli = _get_openai_client()
    resp = cli.chat.completions.create(model=model, messages=[{"role":"user","content":prompt}], temperature=temperature)
    return resp.choices[0].message.content.strip()

def _make_prompt(mode: str, rfp_text: str, 요청사항: str,
                 기존제안: Optional[str] = None,
                 qas_text: Optional[str] = None,
                 single_bench: Optional[Tuple[str, str]] = None) -> str:
    header = "당신은 EY, 맥킨지 수준의 전략 컨설팅 전문가입니다.\n"
    common = f"[RFP 요약]\n{rfp_text}\n\n[제안요청사항]\n{요청사항}\n\n- markdown으로 작성하지 마세요\n"
    if mode == "실현":
        return (header + common +
                f"[기존 제안방안]\n{(기존제안 or '').strip()}\n\n[실현질문 및 답변]\n{qas_text or ''}").strip()
    if mode == "벤치" and single_bench:
        q,a = single_bench
        return (header + common +
                "[벤치마킹질문]\n" f"Q: {q}\nA: {a}").strip()
    return header + "입력 형식 오류"

def generate_updated_plans_per_question(rfp_dict: Dict[str, Dict[str, Any]],
                                        answer_dict: Dict[str, Dict[str, List[Any]]],
                                        rfp_text: str,
                                        model: str = "gpt-4o") -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for req_id, sec in rfp_dict.items():
        req_title = sec.get("제안요청사항","")
        base_plan = sec.get("제안방안","")
        # 실현질문 → 통합
        SJ = sec.get("실현질문",[]) or []
        SA = (answer_dict.get(req_id,{}) or {}).get("실현질문_답변",[]) or []
        if SJ and SA:
            qas = "\n".join([f"{i+1}. Q: {q}\n   A: {a}" for i,(q,a) in enumerate(zip(SJ,SA))])
            prompt = _make_prompt("실현", rfp_text, req_title, 기존제안=base_plan, qas_text=qas)
            result = _call_llm_simple(prompt, model=model, temperature=0.3)
            rows.append({"요청 ID":req_id,"제안요청사항":req_title,"질문유형":"실현질문",
                         "질문":"\n\n".join(SJ),"답변":"\n\n".join(SA),"urls":(answer_dict.get(req_id,{}).get("실현질문_url",[])),
                         "updated_제안방안":result})
        # 벤치마킹 → 개별
        BJ = sec.get("벤치마킹질문",[]) or []
        BA = (answer_dict.get(req_id,{}) or {}).get("벤치마킹질문_답변",[]) or []
        BU = (answer_dict.get(req_id,{}) or {}).get("벤치마킹질문_url",[]) or []
        for i,(q,a) in enumerate(zip(BJ,BA)):
            urls = BU[i] if i < len(BU) else []
            prompt = _make_prompt("벤치", rfp_text, req_title, single_bench=(q,a))
            result = _call_llm_simple(prompt, model=model, temperature=0.3)
            rows.append({"요청 ID":req_id,"제안요청사항":req_title,"질문유형":"벤치마킹질문",
                         "질문":q,"답변":a,"urls":urls,"updated_제안방안":result})
    return pd.DataFrame(rows)

# ---- (6) Slim Deck 생성 (요청별 4옵션 → 슬라이드 흐름) ----------------------------
def _flatten_urls_cell(cell) -> List[str]:
    if cell is None or (isinstance(cell,float) and pd.isna(cell)):
        return []
    if isinstance(cell,(list,tuple)):
        return [str(u) for u in cell if u]
    return [str(cell)]

def _resp_text(resp) -> str:
    # Responses API용 텍스트 추출 — 현재는 Chat API만 사용하므로 미사용. 남겨둠.
    return str(resp)

def _extract_json_block(text: str) -> Optional[Dict[str, Any]]:
    if not text: return None
    m = re.search(r"```json\s*(\{.*?\})\s*```", text, flags=re.S)
    if m:
        try: return json.loads(m.group(1))
        except: pass
    s,e = text.find("{"), text.rfind("}")
    if s!=-1 and e!=-1 and e>s:
        try: return json.loads(text[s:e+1])
        except: return None
    return None

def _build_request_deck_prompt(req_id: str, req_title: str, option_seeds: List[Dict[str, Any]], max_slides_per_option: int = 6) -> str:
    seeds_txt = []
    for seed in option_seeds:
        urls_block = "\n".join([f"  - {u}" for u in seed.get("urls", [])]) or "  - (없음)"
        seeds_txt.append(
            f"[옵션 {seed['option_no']} 시드]\n"
            f"Q: {seed.get('question','')}\n"
            f"A: {seed.get('answer','')}\n"
            f"Updated Plan: {seed.get('updated_plan','')}\n"
            f"URLs:\n{urls_block}\n"
        )
    seeds_joined = "\n".join(seeds_txt)
    schema = {
        "req_id": req_id,
        "req_title": req_title,
        "overview_slide": {"title":"요청사항별 제안 옵션 개요","subtitle":"4개 옵션 비교",
                           "purpose":"고객이 옵션을 빠르게 비교·선택하도록","table":{"columns":["옵션","핵심 전략","선택 기준","예상 기간(주)","주요 리스크"],"rows":[["옵션 1","","","",""]]}},
        "options":[{"option_no":1,"option_title":"", "why_choose":[],"fit_signals":[],"risks":[],"mitigations":[],
                    "timeline":[{"phase":"설계","duration_weeks":2},{"phase":"파일럿","duration_weeks":4}],
                    "slides":[{"slide_no":"1","title":"","subtitle":"","purpose":"","key_messages":[],"content_draft":"","paste_blocks":{"title_text":"","subtitle_text":"","body_bulleted":[]},"urls":[]}]}
                 ]
    }
    schema_json = json.dumps(schema, ensure_ascii=False, indent=2)
    rules = [
        "- 출력은 반드시 JSON 하나만.",
        "- 각 옵션은 6~12장 슬라이드.",
        "- 제목/부제목은 실제 PPT 표기처럼 명확하고 구체적으로.",
        "- seed URLs만 인용(없으면 일반론 금지하고 구조/실행 중심).",
        "- paste_blocks는 붙여넣기용 텍스트로 채우기.",
    ]
    return (
        "당신은 EY/맥킨지 스타일의 전략 컨설팅 문서 설계 전문가입니다.\n"
        f"요청사항 하나({req_id} - {req_title})에 대해, 아래 시드 4개를 각기 다른 제안 옵션으로 확장하고,\n"
        "각 옵션에 대한 슬라이드 흐름(6~12장)을 설계하세요. 고객이 선택하기 쉽게 옵션 메타도 포함하세요.\n\n"
        f"{seeds_joined}\n"
        "[규칙]\n- " + "\n- ".join(rules) + "\n\n"
        "[JSON 스키마(참고)]\n" + schema_json + "\n\n"
        "JSON만 반환하세요."
    )

def build_master_deck_slim(updated_df: pd.DataFrame, model: str = "gpt-4o",
                           max_slides_per_option: int = 6, options_per_req: int = 4) -> Dict[str, Any]:
    need = {"요청 ID","제안요청사항","질문유형","질문","답변","urls","updated_제안방안"}
    miss = need - set(updated_df.columns)
    if miss:
        raise ValueError(f"updated_df 컬럼 누락: {sorted(miss)}")
    cli = _get_openai_client()
    sections = []
    for req_id, g in updated_df.groupby("요청 ID", sort=False):
        req_title = str(g["제안요청사항"].iloc[0])
        seeds = []
        for i, (_, row) in enumerate(g.head(options_per_req).iterrows(), start=1):
            seeds.append({
                "option_no": i,
                "question": str(row["질문"]),
                "answer": str(row["답변"]) if pd.notna(row["답변"]) else "",
                "updated_plan": str(row["updated_제안방안"]) if pd.notna(row["updated_제안방안"]) else "",
                "urls": _flatten_urls_cell(row["urls"])
            })
        prompt = _build_request_deck_prompt(str(req_id), req_title, seeds, max_slides_per_option)
        # JSON 강제는 Chat API에선 보장 어렵기에 추출기로 안정화
        resp = cli.chat.completions.create(model=model, messages=[{"role":"user","content":prompt}], temperature=0.2)
        content = resp.choices[0].message.content.strip()
        data = _extract_json_block(content) or {}
        sections.append({
            "req_id": str(req_id),
            "req_title": req_title,
            "overview_slide": data.get("overview_slide", {}),
            "options": data.get("options", [])
        })
    master = {
        "cover":{"title":"제안서(슬림)","subtitle":"요청사항별 4개 제안 옵션과 슬라이드 흐름","fields":{"고객사":"", "작성일":"", "작성팀":""}},
        "sections": sections,
        "closing":{"title":"다음 단계","bullets":["옵션 선택 워크숍","데이터/사전조건 점검","파일럿 범위 합의 및 킥오프"]}
    }
    return master

def deck_to_dataframe_slim(master: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    rows.append({"요청 ID":"COVER","요청 제목":"(표지)","옵션번호":"","슬라이드번호":1,"제목":master["cover"]["title"],
                 "부제목":master["cover"]["subtitle"],"본문초안":json.dumps(master["cover"]["fields"], ensure_ascii=False),
                 "왜_이_옵션":"", "적합_시그널":"", "리스크":"", "완화책":"", "타임라인":"", "URL":""})
    for sec in master.get("sections", []):
        req_id, req_title = sec.get("req_id",""), sec.get("req_title","")
        ov = sec.get("overview_slide", {})
        if ov:
            rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":"OVERVIEW","슬라이드번호":0,
                         "제목":ov.get("title","옵션 개요"),"부제목":ov.get("subtitle",""),
                         "본문초안":ov.get("purpose",""),"왜_이_옵션":"", "적합_시그널":"", "리스크":"", "완화책":"",
                         "타임라인":json.dumps(ov.get("table",{}), ensure_ascii=False),"URL":""})
        for opt in sec.get("options", []):
            opt_no = opt.get("option_no","")
            why = "\n- " + "\n- ".join(opt.get("why_choose", [])) if opt.get("why_choose") else ""
            fit = "\n- " + "\n- ".join(opt.get("fit_signals", [])) if opt.get("fit_signals") else ""
            risk = "\n- " + "\n".join(["- "+x for x in (opt.get("risks", []) or [])]) if opt.get("risks") else ""
            miti = "\n- " + "\n- ".join(opt.get("mitigations", [])) if opt.get("mitigations") else ""
            timeline = json.dumps(opt.get("timeline", []), ensure_ascii=False)
            rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":opt_no,"슬라이드번호":"META","제목":opt.get("option_title",""),
                         "부제목":"(옵션 요약)","본문초안":"", "왜_이_옵션":why,"적합_시그널":fit,"리스크":risk,"완화책":miti,
                         "타임라인":timeline,"URL":""})
            for s in opt.get("slides", []):
                urls = "\n".join(s.get("urls", [])) if s.get("urls") else ""
                body = s.get("content_draft","")
                pb = s.get("paste_blocks", {})
                if pb:
                    bullets = pb.get("body_bulleted", [])
                    if bullets:
                        body += "\n\n[붙여넣기]\n- " + "\n- ".join(bullets)
                rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":opt_no,"슬라이드번호":s.get("slide_no"),
                             "제목":s.get("title"),"부제목":s.get("subtitle"),"본문초안":body,
                             "왜_이_옵션":"", "적합_시그널":"", "리스크":"", "완화책":"", "타임라인":"", "URL":urls})
    rows.append({"요청 ID":"CLOSING","요청 제목":"(마무리)","옵션번호":"","슬라이드번호":1,"제목":master["closing"]["title"],
                 "부제목":"","본문초안":"\n- " + "\n- ".join(master["closing"]["bullets"]),
                 "왜_이_옵션":"", "적합_시그널":"", "리스크":"", "완화책":"", "타임라인":"", "URL":""})
    return pd.DataFrame(rows)

# ------------------------------------------------------------------------------------
# 2) 오케스트레이션: 입력 → (LLM 파이프라인) → slim_master_slide_flows DF
# ------------------------------------------------------------------------------------
def build_flows_from_user_inputs(
    rfp_path: str | Path,
    client_name: str,
    user_direction: str,
    notes: str = "",
    model_main: str = "gpt-4o"
) -> pd.DataFrame:
    """엑셀 업로드 없이 바로 옵션 선택에 쓸 DF 생성"""
    # 1) 원문 추출
    rfp_text = extract_rfp_text(rfp_path)

    # 2) 요구/방안/질문 목록 생성 (LLM)
    response_lines = rfp_requirement_check(
        {"rfp_text": rfp_text, "client_name": client_name, "user_direction": user_direction},
        model=model_main, temperature=0.1
    )
    rfp_dict = parse_rfp_response_to_dict(response_lines)

    # 3) 외부 검색(Perplexity): 각 질문에 대해 Q→A, citations 수집
    answer_dict = {}
    for req_id, section in rfp_dict.items():
        req_title = section.get("제안요청사항", "")
        answer_dict[req_id] = {"실현질문_답변": [], "실현질문_url": [], "벤치마킹질문_답변": [], "벤치마킹질문_url": []}

        for q in section.get("실현질문", []):
            cq = make_contextual_question(req_title, q)
            a_text, urls = search_perplexity(cq)  # ← SONAR 아님
            answer_dict[req_id]["실현질문_답변"].append(a_text)
            answer_dict[req_id]["실현질문_url"].append(urls)

        for q in section.get("벤치마킹질문", []):
            cq = make_contextual_question(req_title, q)
            a_text, urls = search_perplexity(cq)  # ← SONAR 아님
            answer_dict[req_id]["벤치마킹질문_답변"].append(a_text)
            answer_dict[req_id]["벤치마킹질문_url"].append(urls)

    # 4) Q/A와 RFP로 'updated_제안방안' 작성
    updated_df = generate_updated_plans_per_question(rfp_dict, answer_dict, rfp_text, model=model_main)
    # 필수 컬럼 확보(빈 값은 채워줌)
    for col in ["요청 ID","제안요청사항","질문유형","질문","답변","urls","updated_제안방안"]:
        if col not in updated_df.columns:
            updated_df[col] = ""

    # 5) 요청별 4옵션 슬림 덱 구조 생성 → DF로 전개
    master = build_master_deck_slim(updated_df, model=model_main, max_slides_per_option=6, options_per_req=4)
    df_flows = deck_to_dataframe_slim(master)

    # 6) 앱의 스키마와 궁합 맞추기(미존재 컬럼 보강)
    for col in ["왜_이_옵션","적합_시그널","리스크","완화책","타임라인","URL","옵션대제목"]:
        if col not in df_flows.columns:
            df_flows[col] = ""

    # 7) 반환: 엑셀 대체 DF
    return df_flows
# --- NEW: Perplexity(SONAR) 검색 유틸 ---
# --- Perplexity 검색 (SONAR 미사용, pplx-* 전용) ---
import os, time, requests

def search_perplexity(
    query: str,
    api_key: str | None = None,
    model: str | None = None,
    max_retry: int = 2,
    timeout: int = 60,
) -> tuple[str, list[str]]:
    """
    Perplexity Chat Completions(OpenAI 호환) 호출 (SONAR 사용 안 함).
    - 기본 모델: pplx-70b-online  (없으면 pplx-7b-online로 폴백)
    - 반환: (answer_text, citations_urls[])
    """
    api_key = api_key or os.getenv("PERPLEXITY_API_KEY")
    if not api_key:
        raise RuntimeError("Perplexity API 키가 없습니다. 환경변수 PERPLEXITY_API_KEY 를 설정하세요.")

    model = model or os.getenv("PERPLEXITY_MODEL") or "pplx-70b-online"

    url = "https://api.perplexity.ai/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": query}],
    }

    last_err = None
    for attempt in range(max_retry + 1):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()

            # 본문
            content = data["choices"][0]["message"]["content"].strip()

            # citations 파싱(버전에 따라 위치가 다름)
            cites: list[str] = []
            # 1) top-level
            if isinstance(data.get("citations"), list):
                cites.extend([c for c in data["citations"] if isinstance(c, str)])
            # 2) message.citations
            msg = data["choices"][0].get("message", {})
            if isinstance(msg.get("citations"), list):
                cites.extend([c for c in msg["citations"] if isinstance(c, str)])
            # 3) message.context (일부 응답)
            if isinstance(msg.get("context"), list):
                cites.extend([c for c in msg["context"] if isinstance(c, str)])
            # 4) search_results (일부 응답 스키마)
            if isinstance(data.get("search_results"), list):
                for sr in data["search_results"]:
                    u = sr.get("url")
                    if isinstance(u, str):
                        cites.append(u)

            # 중복 제거 & 간단 필터
            uniq = []
            for u in cites:
                if u and u not in uniq:
                    uniq.append(u)

            return content, uniq

        except requests.HTTPError as e:
            # 모델 미지원/액세스 오류 시 pplx-7b-online로 1회 폴백
            if attempt == 0 and model != "pplx-7b-online":
                model = "pplx-7b-online"
                payload["model"] = model
                last_err = e
                time.sleep(0.7)
                continue
            last_err = e
        except Exception as e:
            last_err = e

        time.sleep(0.7)

    raise RuntimeError(f"Perplexity 검색 실패: {last_err}")

