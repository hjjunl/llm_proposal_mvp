# -*- coding: utf-8 -*-
# pipeline/rfp2proposal.py  (drop-in 교체)
from __future__ import annotations

import os, re, json, time, traceback, requests
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Callable
from dotenv import load_dotenv
load_dotenv()

import pandas as pd

__all__ = [
    "build_flows_from_user_inputs",  # Tab2에서 import 하는 엔트리포인트
    "extract_text_from_file",        # Tab2 진행 로그에서 호출
]

# =========================
# 공통 유틸 & 로거
# =========================
from datetime import datetime

# 파일명 날짜 접미사: 기본은 오늘(YYYYMMDD), .env에 FILE_DATE_SUFFIX 있으면 그 값 사용
DATE_SUFFIX = os.getenv("FILE_DATE_SUFFIX", datetime.now().strftime("%Y%m%d"))

def _with_suffix(path: str | Path, suffix: str | None = None) -> str:
    """파일 경로에 _YYYYMMDD 같은 접미사를 확장자 앞에 붙임"""
    p = Path(path)
    sfx = suffix if suffix is not None else f"_{DATE_SUFFIX}"
    return str(p.with_name(f"{p.stem}{sfx}{p.suffix}"))


def _get(key: str, default: str = "") -> str:
    val = os.getenv(key, default)
    if not val:
        # Streamlit secrets 지원(있을 때만)
        try:
            import streamlit as st
            val = st.secrets.get(key, default)  # type: ignore[attr-defined]
        except Exception:
            pass
    return val

def _now() -> str:
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _log(logf: Optional[Callable[[str], None]], msg: str):
    if logf:
        try:
            logf(msg)
        except Exception:
            pass

# =========================
# 파일 텍스트 추출
# =========================
def extract_text_from_file(path: str | Path) -> str:
    p = Path(path)
    ext = p.suffix.lower()
    text = ""
    if ext == ".pdf":
        try:
            from pypdf import PdfReader
        except Exception:
            from PyPDF2 import PdfReader
        r = PdfReader(str(p))
        text = "\n".join((pg.extract_text() or "") for pg in r.pages)
    elif ext == ".docx":
        import docx
        d = docx.Document(str(p))
        text = "\n".join(para.text for para in d.paragraphs)
    elif ext in (".txt", ".md"):
        text = p.read_text(encoding="utf-8", errors="ignore")
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}")

    # 후처리
    text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()

# =========================
# OpenAI (chat + responses)
# =========================
@dataclass
class OpenAIConf:
    model: str = "gpt-5"
    temperature: float = 1

def _openai_client():
    from openai import OpenAI
    api = _get("OPENAI_API_KEY")
    if not api:
        raise RuntimeError("OPENAI_API_KEY 없음")
    return OpenAI(api_key=api)

def _chat(model: str, messages: List[Dict[str, str]], temperature: float = 1) -> str:
    cli = _openai_client()
    resp = cli.chat.completions.create(model=model, messages=messages, temperature=temperature)
    return resp.choices[0].message.content.strip()

def _responses_client():
    from openai import OpenAI
    api = _get("OPENAI_API_KEY")
    if not api:
        raise RuntimeError("OPENAI_API_KEY 없음")
    return OpenAI(api_key=api)

def _resp_text(resp) -> str:
    # 다양한 포맷 방어 추출
    try:
        return resp.output_text
    except Exception:
        pass
    try:
        parts = resp.output[0].content
        for p in parts:
            if getattr(p, "text", None) and getattr(p.text, "value", None):
                return p.text.value
    except Exception:
        pass
    try:
        return str(resp)
    except Exception:
        return ""

def _extract_json_block(txt: str) -> Dict[str, Any]:
    m = re.search(r"```json\s*(\{.*?\})\s*```", txt, flags=re.S)
    if m:
        return json.loads(m.group(1))
    s, e = txt.find("{"), txt.rfind("}")
    if s != -1 and e != -1 and e > s:
        return json.loads(txt[s:e+1])
    raise ValueError("JSON 파싱 실패")

def _responses_json(prompt: str, model: str = "gpt-5") -> Dict[str, Any]:
    cli = _responses_client()
    r = cli.responses.create(model=model, input=prompt)
    raw = _resp_text(r)

    # (선택) 디버그 저장
    try:
        debug_dir = Path("DB/_debug"); debug_dir.mkdir(parents=True, exist_ok=True)
        (debug_dir / "last_responses_raw.txt").write_text(raw, encoding="utf-8")
    except Exception:
        pass

    # JSON 블록만 뽑기 (불가시 예외)
    try:
        return _extract_json_block(raw)
    except Exception:
        # 약식 폴백: 전체를 json.loads 시도
        try:
            return json.loads(raw)
        except Exception as e:
            # 최종 실패 → 상위에서 폴백 구조 생성
            raise ValueError(f"Responses JSON 파싱 실패: {e}")


# =========================
# Perplexity / SerpAPI
# =========================
def _get_pplx_key() -> str:
    for k in ["PERPLEXITY_API_KEY", "PPLX_API_KEY", "PEPLEXITY_API_KEY"]:
        v = _get(k, "")
        if v:
            return v
    return ""

# 허용 모델: sonar, sonar-small, sonar-medium, sonar-pro
_PPLX_DEFAULT_MODEL = "sonar-pro"

def search_perplexity(query: str, model: str = _PPLX_DEFAULT_MODEL, temperature: float = 1) -> Tuple[str, List[str]]:
    api_key = _get_pplx_key()
    if not api_key:
        raise RuntimeError("Perplexity API Key 없음")

    url = "https://api.perplexity.ai/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model,
        "temperature": temperature,
        "return_citations": True,   # <<< 중요: 인용 반환
        "messages": [
            {"role": "system", "content": "You are a precise research assistant. Always cite sources as links."},
            {"role": "user", "content": query},
        ],
    }
    r = requests.post(url, headers=headers, json=payload, timeout=90)
    if r.status_code >= 400:
        raise RuntimeError(f"Perplexity {r.status_code}: {r.text}")
    data = r.json()

    content = data["choices"][0]["message"]["content"]
    urls: List[str] = []
    try:
        # 모델/버전에 따라 위치가 다를 수 있어 모두 시도
        cand = data.get("citations") or data["choices"][0]["message"].get("citations") or []
        if isinstance(cand, list):
            urls = [u for u in cand if isinstance(u, str)]
    except Exception:
        pass

    if not urls:
        urls = _extract_urls_from_text(content)  # 본문에서 폴백 추출

    return content, urls


def search_serpapi(query: str, num_results: int = 5) -> Tuple[str, List[str]]:
    key = _get("SERP_API_KEY") or _get("SERPAPI_API_KEY") or _get("serp_api_key") or _get("SERP_API")
    if not key:
        raise RuntimeError("SerpAPI Key 없음")
    url = "https://serpapi.com/search"
    params = {"q": query, "api_key": key, "num": num_results, "engine": "google"}
    r = requests.get(url, params=params, timeout=60)
    r.raise_for_status()
    js = r.json()
    snippets, urls = [], []
    for it in js.get("organic_results", []):
        if it.get("snippet"): snippets.append(it["snippet"])
        if it.get("link"): urls.append(it["link"])
    content = " ".join(snippets) if snippets else "No relevant content."
    return content, urls

# =========================
# 1) RFP → 요청/질문 생성 (OpenAI)
# =========================
_REQ_SYS = (
    "산출물은 실제 제안서 본문에 즉시 삽입 가능한 수준의 전문 한국어 문체로 작성합니다. "
    "위 정보를 바탕으로 제안 요청사항 및 제안 방안을 상세히 정리해주세요."
)


def rfp_requirement_check(user_inputs: dict, conf: OpenAIConf) -> List[str]:
    rfp_text = user_inputs["rfp_text"]
    user_direction = user_inputs.get("user_direction", "")
    keywords = user_inputs.get("keywords", [])

    system_prompt = f"""
    [RFP]
    {rfp_text}

    [고객 방향성]
    {user_direction or '없음'}

    [강조 키워드] {", ".join(keywords) if keywords else "없음"}
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
        
        예시 포멧. (질문을 각각 3개씩 만들어야함)
        1. 제안 요청사항: 삼성전자 직영 매장의 성공적인 구축 및 운영 모델확보
        - 제안 방안: 삼성전자의 직영 매장 구축 및 운영을 위한 성공 모델을 확보하기 위해, 제조업과 리테일 업계의 차이점을 분석하고, 내/외부 전문가와의 협업을 통해 빠른 내부 역량 배양을 추진합니다. 또한, 미국 MX 직영 매장의 사례를 벤치마킹하여 글로벌 직영 매장 구축 및 운영 표준 가이드라인을 수립합니다.
        - 제안 방안을 실제로 실현 할 수 있기 위한 질문 3개:
        (1)전자 산업의 제조업과 리테일 업계의 차이점을 분석을 어떻게 하고 차이점은 무엇인가요?
        (2) ...etc
        (3) ...etc
        - 제안 요청사항의 추가 제안 방안 예시를 알기 위한 질문 3개 (벤치마킹 질문): 
        (1)전자 리테일 직영 매장 운영의 모든 영역을 매뉴얼화하기 위해 필요한 세부 항목은 무엇이며, 이를 어떻게 구체적으로 정의할 수 있을까요?
        (2) ...etc
        (3) ...etc
    """
    txt = _chat(conf.model, [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": _REQ_SYS},
    ], temperature=conf.temperature)
    return [ln.strip() for ln in txt.splitlines() if ln.strip()]

def parse_rfp_response_to_dict(lines: List[str]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    sec, idx, curkey = {}, 0, ""
    for ln in lines:
        if re.match(r"^\d+\.\s", ln):
            if sec: out[f"요청사항_{idx}"] = sec
            idx += 1
            sec = {"제안요청사항":"", "제안방안":"", "실현질문":[], "벤치마킹질문":[]}
            sec["제안요청사항"] = ln.split(":",1)[-1].strip()
            curkey = ""
            continue
        if ln.startswith("- 제안 방안:"):
            sec["제안방안"] = ln.split(":",1)[-1].strip(); continue
        if "실현 할 수 있기 위한 질문" in ln: curkey = "실현질문"; continue
        if "벤치마킹" in ln: curkey = "벤치마킹질문"; continue
        if ln.startswith("(") and curkey:
            sec[curkey].append(ln)
    if sec: out[f"요청사항_{idx}"] = sec
    return out

# =========================
# 2) Perplexity/SerpAPI 리서치
# =========================
def _ctx_q(req_title: str, q: str) -> str:
    req_title, q = (req_title or "").strip(), (q or "").strip()
    return f"{req_title} 관점에서, {q}"

def generate_answer_dict(rfp_dict: Dict[str, Any], prefer: str = "perplexity",
                         logf: Optional[Callable[[str], None]] = None) -> Dict[str, Any]:
    ans: Dict[str, Any] = {}
    pplx_ok = bool(_get_pplx_key())
    total_q = sum(len(sec.get("실현질문", [])) + len(sec.get("벤치마킹질문", [])) for sec in rfp_dict.values())
    _log(logf, f"예상 리서치 호출 수: {total_q} (prefer={prefer}, pplx_ok={pplx_ok})")

    for req_id, sec in rfp_dict.items():
        req_title = sec.get("제안요청사항","")
        ans[req_id] = {"실현질문_답변":[], "실현질문_url":[], "벤치마킹질문_답변":[], "벤치마킹질문_url":[]}

        # 실현질문
        for q in sec.get("실현질문", []):
            cq = _ctx_q(req_title, q)
            try:
                if prefer == "perplexity" and pplx_ok:
                    _log(logf, f"[{req_id}] PPLX 질의: {cq}")
                    a, urls = search_perplexity(cq, model=_PPLX_DEFAULT_MODEL)
                else:
                    _log(logf, f"[{req_id}] SerpAPI 질의: {cq}")
                    a, urls = search_serpapi(cq)
            except Exception as e:
                _log(logf, f"[{req_id}] 리서치 실패 → 휴리스틱 대체: {e}")
                a, urls = ("리서치 실패. RFP 문맥 기반 추정 답변.", [])
            ans[req_id]["실현질문_답변"].append(a)
            ans[req_id]["실현질문_url"].append(urls)

        # 벤치마킹질문
        for q in sec.get("벤치마킹질문", []):
            cq = _ctx_q(req_title, q)
            try:
                if prefer == "perplexity" and pplx_ok:
                    _log(logf, f"[{req_id}] PPLX 질의(벤치): {cq}")
                    a, urls = search_perplexity(cq, model=_PPLX_DEFAULT_MODEL)
                else:
                    _log(logf, f"[{req_id}] SerpAPI 질의(벤치): {cq}")
                    a, urls = search_serpapi(cq)
            except Exception as e:
                _log(logf, f"[{req_id}] 리서치 실패(벤치) → 휴리스틱 대체: {e}")
                a, urls = ("리서치 실패. 일반적 벤치마킹 방향 제시.", [])
            ans[req_id]["벤치마킹질문_답변"].append(a)
            ans[req_id]["벤치마킹질문_url"].append(urls)
    return ans

# === URL helpers (상단 import 근처) ===
_URL_RE = re.compile(r'https?://[^\s)>\]"}]+')

def _dedupe_preserve_order(seq: List[str]) -> List[str]:
    seen, out = set(), []
    for x in seq:
        if not x or x in seen:
            continue
        seen.add(x); out.append(x)
    return out

def _extract_urls_from_text(text: str) -> List[str]:
    return _dedupe_preserve_order(_URL_RE.findall(text or ""))

def _force_url_list(value: Any) -> List[str]:
    """중첩(list/tuple/set)·문자열(리스트 문자열 포함) 등을 모두 URL 리스트로 평탄화"""
    out: List[str] = []
    def rec(v: Any):
        if v is None:
            return
        if isinstance(v, (list, tuple, set)):
            for it in v: rec(it)
        elif isinstance(v, str):
            # "['http..','http..']" 같은 문자열도 정규식으로 추출
            out.extend(_URL_RE.findall(v))
        else:
            out.extend(_URL_RE.findall(str(v)))
    rec(value)
    return _dedupe_preserve_order(out)

# =========================
# 3) Q/A → updated_제안방안 (OpenAI)
# =========================
def _mk_prompt(mode: str, rfp_text: str, req: str,
               base_plan: str, qas: Optional[str] = None,
               single_bench: Optional[Tuple[str,str]] = None) -> str:
    head = (
        "당신은 EY/맥킨지 수준의 전략 컨설팅 전문가입니다.\n"
        "문체: 컨설턴트 장표와 같이, 마크다운 금지\n" # 간결·권위적·전문 용어 사용 
        "다음 제안요청사항과 RFP 본문, 하나의 벤치마킹 질문과 그에 대한 답변을 바탕으로, "
        "참조 사례와 보다 실행 가능하고 설득력 있는 'updated_제안방안'을 작성해 주세요.\n"
        "참조 사례와 실행방안을 포함한 'updated_제안방안'을 작성해 주세요.\n"
        "제안 방안은 최대한 상세히 작성 (실제 action itemd으로)\n"
    )
    base = f"[RFP]\n{rfp_text}\n\n[제안요청사항]\n{req}\n\n[기존 제안방안]\n{base_plan}\n\n"
    if mode == "실현":
        return head + base + "[실현질문·답변]\n" + (qas or "")
    if mode == "벤치" and single_bench:
        q, a = single_bench
        return head + base + f"[벤치마킹 질문]\nQ: {q}\nA: {a}"
    return head + base


def generate_updated_plans_per_question(rfp_dict: Dict[str, Any],
                                        answer: Dict[str, Any],
                                        rfp_text: str,
                                        conf: OpenAIConf,
                                        logf: Optional[Callable[[str], None]] = None) -> pd.DataFrame:
    rows: List[Dict[str,Any]] = []
    for req_id, sec in rfp_dict.items():
        req = sec.get("제안요청사항",""); base = sec.get("제안방안","")

        # 실현(모아쓰기)
        qs = sec.get("실현질문", []); ans = answer.get(req_id, {}).get("실현질문_답변", [])
        urls_list = answer.get(req_id, {}).get("실현질문_url", [])
        if qs and ans:
            combined = "\n".join([f"{i+1}. Q: {q}\n   A: {a}" for i,(q,a) in enumerate(zip(qs,ans))])
            prompt = _mk_prompt("실현", rfp_text, req, base, qas=combined)
            _log(logf, f"[{req_id}] updated_제안방안(실현) 생성 호출")
            up = _chat(conf.model, [{"role":"user","content":prompt}], temperature=1)
            rows.append({
                "요청 ID": req_id, "제안요청사항": req, "질문유형":"실현질문",
                "질문":"\n\n".join(qs), "답변":"\n\n".join(ans),
                "urls": _force_url_list(urls_list), "updated_제안방안": up
            })

        # 벤치(개별)
        bq = sec.get("벤치마킹질문", []); ba = answer.get(req_id, {}).get("벤치마킹질문_답변", [])
        burl = answer.get(req_id, {}).get("벤치마킹질문_url", [])
        for i,(q,a) in enumerate(zip(bq, ba)):
            prompt = _mk_prompt("벤치", rfp_text, req, base, single_bench=(q,a))
            _log(logf, f"[{req_id}] updated_제안방안(벤치) 생성 호출 - Q{i+1}")
            up = _chat(conf.model, [{"role":"user","content":prompt}], temperature=1)
            rows.append({
                "요청 ID": req_id, "제안요청사항": req, "질문유형":"벤치마킹질문",
                "질문": q, "답변": a,
                "urls": _force_url_list(burl[i] if i < len(burl) else []),  # <<< 정규화
                "updated_제안방안": up
            })
    return pd.DataFrame(rows)

# =========================S
# 4) updated_df → 슬림 옵션/슬라이드 (OpenAI Responses)
# =========================
def _flatten_urls(cell) -> List[str]:
    if cell is None: return []
    if isinstance(cell, (list, tuple)): return [str(x) for x in cell if x]
    return [str(cell)]

def _build_deck_prompt(req_id: str, req_title: str, seeds: List[Dict[str,Any]]) -> str:
    seeds_txt = []
    for s in seeds:
        seeds_txt.append(
            f"[옵션 {s['option_no']} 시드]\n"
            f"Q: {s['question']}\nA: {s['answer']}\n"
            f"Updated: {s['updated_plan']}\nURLs:\n" +
            ("\n".join(f"  - {u}" for u in s.get("urls", [])) if s.get("urls") else "  - (없음)")
        )

    schema = {
        "req_id": req_id, "req_title": req_title,
        "overview_slide": {
            "title": "요청별 옵션 개요",
            "subtitle": "비교",
            "table": {"columns": ["옵션","핵심전략","기간(주)","주요KPI","리스크"],
                      "rows": [["1","","","",""]]}
        },
        "options": [{
            "option_no": 1,
            "option_title": "",
            "why_choose": [],
            "fit_signals": [],
            "risks": [],
            "mitigations": [],
            "timeline": [ {"phase":"","duration_weeks":""} ],
            "kpis": [ {"name":"","baseline":"","target":"","cadence":""} ],
            "slides": [{
                "slide_no": 1,
                "title": "",
                "subtitle": "",
                "purpose": "이 슬라이드로 설득하려는 핵심",
                "key_messages": ["","",""],
                "content_draft": "상세히 작성",
                "paste_blocks": {"body_bulleted":[]},
                "urls": []
            }]
        }]
    }

    # 튜플 → 문자열로 조인 (버그 픽스)
    rules = (
        "- 출력은 반드시 JSON 하나만.",
        "- 각 옵션은 5~10장 슬라이드.",
        "- 각 제안별 슬라이드는 1~3장 이상 차이가 나지 않게 구성.",
        "- 제목/부제목은 실제 PPT 표기처럼 명확하고 구체적으로, 자세하게 작성.",
        "- 정량/사례가 필요하면 seed URLs에서만 인용. 없으면 일반론 금지하고 실행/구조 중심으로.",
        "- 각 slide 객체의 'urls' 배열에는 해당 슬라이드의 근거 링크를 seed URLs 중에서 0~3개 기입(원문 링크 그대로).",
        "- paste_blocks는 최대한 상세하고 실현 가능하게 작성.",
    )
    rules_str = "\n".join(rules)

    seeds_block = "\n\n".join(seeds_txt) if seeds_txt else "[시드 없음]"

    return (
        "당신은 컨설팅 문서 설계 전문가입니다.\n"
        f"요청({req_id} {req_title})에 대해 아래 시드 4개를 각각 다른 옵션으로 확장하고 "
        "각 옵션에 대한 슬라이드 흐름(5~10장)을 설계하세요. 고객이 선택하기 쉽게 옵션 메타도 포함하세요\n\n"
        "각 슬라이드 흐름, 의도, 설득력이 있도록 상세히 작성해주세요.\n"
        "본 문 초안을 상세하게 고객이 충분히 설득 될 수 있도록 작성해주세요.(3문장 이상)\n"
        "본문 초안과 제안 포인트는 최대한 구체적이고 실행 가능하게 작성해주세요"
        f"{seeds_block}"
        "\n\n[규칙]\n" + rules_str +
        "\n\n[스키마]\n" + json.dumps(schema, ensure_ascii=False)
    )



# === [PATCH] 추가/수정 부분 시작 ===

__all__ = [
    "build_flows_from_user_inputs",
    "extract_text_from_file",
    "run_slim_pipeline",                 # ← (추가) 기존 노트북 진입점과 동일 역할
]

def _slim_from_updated_df(
    updated_df: pd.DataFrame,
    model: str = "gpt-5",
    logf: Optional[Callable[[str], None]] = None,
    collect_sections: Optional[List[Dict[str, Any]]] = None,   # ← (추가) 섹션 원본 수집
) -> pd.DataFrame:
    rows: List[Dict[str,Any]] = []
    # COVER
    rows.append({"요청 ID":"COVER","요청 제목":"(표지)","옵션번호":"","슬라이드번호":"1",
                 "제목":"제안서(슬림)","부제목":"옵션/슬라이드 흐름","본문초안":"",
                 "왜_이_옵션":"","적합_시그널":"","리스크":"","완화책":"",
                 "타임라인":"","URL":"","옵션대제목":""})

    for req_id, g in updated_df.groupby("요청 ID", sort=False):
        req_title = str(g["제안요청사항"].iloc[0])
        seeds = []
        for i, (_, r) in enumerate(g.head(4).iterrows(), start=1):
            seed_urls = _force_url_list(r.get("urls", []))   # <<< 여기
            seeds.append({
                "option_no": i,
                "question": str(r["질문"]),
                "answer": str(r["답변"]),
                "updated_plan": str(r.get("updated_제안방안","")),
                "urls": seed_urls
            })

        prompt = _build_deck_prompt(str(req_id), req_title, seeds)
        _log(logf, f"[{req_id}] 옵션/슬라이드 설계 호출 (Responses)")
        try:
            js = _responses_json(prompt, model=model)
        except Exception as e:
            _log(logf, f"[{req_id}] Responses 파싱 실패 → 폴백 사용: {e}")
            js = {
                "req_id": str(req_id),
                "req_title": req_title,
                "overview_slide": {"title": "옵션 개요", "subtitle": "", "table": {"columns": [], "rows": []}},
                "options": []
            }
        # (추가) 원본 섹션 수집 → 나중에 master JSON 생성/저장 용
        if collect_sections is not None:
            # js 안에 req_id/req_title/overview_slide/options가 이미 포함되어 있음
            collect_sections.append({
                "req_id": js.get("req_id", str(req_id)),
                "req_title": js.get("req_title", req_title),
                "overview_slide": js.get("overview_slide", {}),
                "options": js.get("options", []),
            })

        ov = js.get("overview_slide", {})
        rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":"OVERVIEW","슬라이드번호":"0",
                     "제목":ov.get("title","옵션 개요"),"부제목":ov.get("subtitle",""),
                     "본문초안":ov.get("purpose",""),"왜_이_옵션":"","적합_시그널":"",
                     "리스크":"","완화책":"",
                     "타임라인":json.dumps(ov.get("table",{}),ensure_ascii=False),
                     "URL":"","옵션대제목":""})
        for opt in js.get("options", []):
            opt_no = str(opt.get("option_no",""))
            rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":opt_no,"슬라이드번호":"META",
                         "제목":opt.get("option_title",""),"부제목":"(요약)","본문초안":"",
                         "왜_이_옵션":"\n- "+"\n- ".join(opt.get("why_choose",[])) if opt.get("why_choose") else "",
                         "적합_시그널":"\n- "+"\n- ".join(opt.get("fit_signals",[])) if opt.get("fit_signals") else "",
                         "리스크":"\n- "+"\n- ".join(opt.get("risks",[])) if opt.get("risks") else "",
                         "완화책":"\n- "+"\n- ".join(opt.get("mitigations",[])) if opt.get("mitigations") else "",
                         "타임라인":json.dumps(opt.get("timeline",[]),ensure_ascii=False),
                         "URL":"","옵션대제목":opt.get("option_title","")})
            for s in opt.get("slides", []):
                body = s.get("content_draft","")
                pb = s.get("paste_blocks",{})
                if pb and isinstance(pb.get("body_bulleted", []), list) and pb["body_bulleted"]:
                    body += "\n\n[제안 포인트]\n- " + "\n- ".join(pb["body_bulleted"])
                rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":opt_no,
                             "슬라이드번호":str(s.get("slide_no","1")),
                             "제목":s.get("title",""),"부제목":s.get("subtitle",""),
                             "본문초안": body,
                             "왜_이_옵션":"","적합_시그널":"","리스크":"","완화책":"",
                             "타임라인":"","URL":"\n".join(s.get("urls",[])) if s.get("urls") else "",
                             "옵션대제목":opt.get("option_title","")})

    # CLOSING
    rows.append({"요청 ID":"CLOSING","요청 제목":"(마무리)","옵션번호":"",
                 "슬라이드번호":"1","제목":"다음 단계","부제목":"",
                 "본문초안":"- 옵션 선택 워크숍\n- 데이터/사전조건 점검\n- 파일럿 범위 합의",
                 "왜_이_옵션":"","적합_시그널":"","리스크":"","완화책":"",
                 "타임라인":"","URL":"","옵션대제목":""})
    Path("DB").mkdir(parents=True, exist_ok=True)
    last_base = "DB/last_updated_df.xlsx"
    last_path = _with_suffix(last_base)  # → DB/last_updated_df_20250930.xlsx
    updated_df.to_excel(last_path, index=False)
    df = pd.DataFrame(rows)
    for c in ["옵션번호","슬라이드번호"]:
        df[c] = df[c].astype(str)
    return df

def _make_master_from_sections(sections: List[Dict[str, Any]]) -> Dict[str, Any]:
    """수집된 섹션 목록 → master JSON 구성(노트북 버전 호환)."""
    return {
        "cover": {
            "title": "제안서(슬림)",
            "subtitle": "요청사항별 4개 제안 옵션과 슬라이드 흐름",
            "fields": {"고객사":"", "작성일":"", "작성팀":""}
        },
        "sections": sections,
        "closing": {
            "title": "다음 단계",
            "bullets": ["옵션 선택 워크숍", "데이터/사전조건 점검", "파일럿 범위 합의 및 킥오프"]
        }
    }
# rfp2proposal.py 상단 유틸 근처에 추가
def _json_default(o):
    # numpy, pandas, set 등 비직렬화 타입 방어
    try:
        import numpy as np
        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return float(o)
        if isinstance(o, (np.ndarray,)):
            return o.tolist()
    except Exception:
        pass
    try:
        import pandas as pd
        if isinstance(o, (pd.Timestamp,)):
            return o.isoformat()
        if hasattr(o, "to_dict"):
            return o.to_dict()
    except Exception:
        pass
    if isinstance(o, set):
        return list(o)
    # 최후 보루
    return str(o)

def run_slim_pipeline(
    updated_df: pd.DataFrame,
    out_dir: str = "DB/proposal_result",
    model: str = "gpt-5",
    logf: Optional[Callable[[str], None]] = None,
) -> pd.DataFrame:
    """노트북의 run_slim_pipeline과 동일한 동작: df 반환 + 파일 저장."""
    os.makedirs(out_dir, exist_ok=True)
    sections: List[Dict[str, Any]] = []
    df = _slim_from_updated_df(updated_df, model=model, logf=logf, collect_sections=sections)
    master = _make_master_from_sections(sections)

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        master = _make_master_from_sections(sections)
        json_path = _with_suffix(os.path.join(out_dir, "slim_master_slide_flows.json"))   # ..._20250930.json
        xlsx_path = _with_suffix(os.path.join(out_dir, "slim_master_slide_flows.xlsx"))   # ..._20250930.xlsx

        # 1차 저장 시도
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(master, f, ensure_ascii=False, indent=2, default=_json_default)
        except TypeError as e:
            # 비직렬화 항목이 있으면 한 번 더 정규화해서 저장
            _log(logf, f"[경고] JSON 비직렬화 요소 감지 → 정규화 후 저장: {e}")
            safe_master = json.loads(json.dumps(master, ensure_ascii=False, default=_json_default))
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(safe_master, f, ensure_ascii=False, indent=2)
    
    df.to_excel(xlsx_path, index=False)
    _log(logf, f"[저장 완료] {json_path}")
    _log(logf, f"[저장 완료] {xlsx_path}")
    return df

# ---- 리서치 상한 (env로 즉시 조정 가능) ----
MAX_REQS = int(os.getenv("RFP_MAX_REQS", "4"))              # 선택할 '제안 요청사항' 최대 개수
MAX_Q_REALIZE = int(os.getenv("RFP_MAX_Q_REALIZE", "2"))    # 요청별 실현질문 최대 개수
MAX_Q_BENCH = int(os.getenv("RFP_MAX_Q_BENCH", "1"))        # 요청별 벤치질문 최대 개수

def _cap_rfp_dict(rfp_dict: Dict[str, Dict[str, Any]],
                  max_reqs: int = MAX_REQS,
                  max_realize: int = MAX_Q_REALIZE,
                  max_bench: int = MAX_Q_BENCH,
                  logf: Optional[Callable[[str], None]] = None) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for i, (req_id, sec) in enumerate(rfp_dict.items()):
        if i >= max_reqs:
            break
        trimmed = dict(sec)
        trimmed["실현질문"] = sec.get("실현질문", [])[:max_realize]
        trimmed["벤치마킹질문"] = sec.get("벤치마킹질문", [])[:max_bench]
        out[req_id] = trimmed
    if logf:
        total_q = sum(len(s.get("실현질문", [])) + len(s.get("벤치마킹질문", [])) for s in out.values())
        _log(logf, f"요청/질문 상한 적용 → 요청 {len(out)}개, 총 질문 {total_q}개 (실현≤{max_realize}, 벤치≤{max_bench})")
    return out

def build_flows_from_user_inputs(
    rfp_path: str,
    client_name: str,
    user_direction: str,
    notes: str = "",
    model_main: str = "gpt-5",
    logf: Optional[Callable[[str], None]] = None,
    out_dir: Optional[str] = "DB/proposal_result",
    model_deck: str = "gpt-5",
) -> pd.DataFrame:
    _log(logf, "텍스트 추출 시작")
    rfp_text = extract_text_from_file(rfp_path)
    _log(logf, f"텍스트 추출 완료 ({len(rfp_text)} chars)")

    # (2) 요청/질문
    _log(logf, "요청/질문 생성 (OpenAI)")
    user_inputs = {
        "rfp_text": rfp_text,
        "style": "신뢰감 있는",
        "keywords": [],
        "client_name": client_name,
        "proposal_title": "",
        "user_direction": user_direction
    }
    req_lines = rfp_requirement_check(user_inputs, OpenAIConf(model=model_main, temperature=1))
    rfp_dict_raw = parse_rfp_response_to_dict(req_lines)
    _log(logf, f"요청 개수(원본): {len(rfp_dict_raw)}")
    rfp_dict = _cap_rfp_dict(rfp_dict_raw, logf=logf)   # ← 하드-컷 적용
    _log(logf, f"요청 개수(상한 적용 후): {len(rfp_dict)}")

    # (3) 리서치
    _log(logf, "Perplexity/SerpAPI 리서치 시작")
    answers = generate_answer_dict(rfp_dict, prefer="perplexity", logf=logf)
    _log(logf, "리서치 완료")

    # (4) 업데이트 제안
    _log(logf, "updated_제안방안 생성 (OpenAI)")
    updated_df = generate_updated_plans_per_question(
        rfp_dict, answers, rfp_text, OpenAIConf(model=model_main, temperature=1), logf=logf
    )
    dbg_base = os.path.join(out_dir or ".", "updated_plans_debug.xlsx")
    dbg_path = _with_suffix(dbg_base)  # → .../updated_plans_debug_20250930.xlsx
    updated_df.to_excel(dbg_path, index=False)
    # 필수 컬럼 보강
    required = {"요청 ID","제안요청사항","질문유형","질문","답변","urls","updated_제안방안"}
    for c in required - set(updated_df.columns):
        updated_df[c] = ""
    _log(logf, f"updated_df 행수: {len(updated_df)}")

    # (5) 슬림 플로우 → df + (옵션) 저장
    _log(logf, "슬라이드 흐름 설계 (Responses)")
    sections: List[Dict[str, Any]] = []
    df = _slim_from_updated_df(updated_df, model=model_deck, logf=logf, collect_sections=sections)

    # 앱 호환 스키마 보강
    need = ["요청 ID","요청 제목","옵션번호","슬라이드번호","제목","부제목","본문초안",
            "왜_이_옵션","적합_시그널","리스크","완화책","타임라인","URL","옵션대제목"]
    for c in need:
        if c not in df.columns: df[c] = ""
    df["슬라이드번호"] = df["슬라이드번호"].astype(str)
    df["옵션번호"] = df["옵션번호"].astype(str)

    # 저장 옵션
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        master = _make_master_from_sections(sections)
        json_path = os.path.join(out_dir, "slim_master_slide_flows.json")
        xlsx_path = os.path.join(out_dir, "slim_master_slide_flows.xlsx")

        # 1차 저장 시도
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(master, f, ensure_ascii=False, indent=2, default=_json_default)
        except TypeError as e:
            _log(logf, f"[경고] JSON 비직렬화 요소 감지 → 정규화 후 저장: {e}")
            safe_master = json.loads(json.dumps(master, ensure_ascii=False, default=_json_default))
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(safe_master, f, ensure_ascii=False, indent=2)

        df.to_excel(xlsx_path, index=False)
        _log(logf, f"[저장 완료] {json_path}")
        _log(logf, f"[저장 완료] {xlsx_path}")

    _log(logf, "파이프라인 완료")
    return df

