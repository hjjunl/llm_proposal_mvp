# pipeline/rfp2proposal.py  (drop-in 교체)
# -*- coding: utf-8 -*-
from __future__ import annotations
import os, re, json, time, traceback, requests
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

__all__ = [
    # Tab2에서 import 하는 엔트리포인트
    "build_flows_from_user_inputs",
    # 유틸
    "extract_text_from_file",
]

# ========== 공통 유틸 ==========

def _get(key: str, default: str = "") -> str:
    # env / st.secrets 둘 다 지원 (streamlit 미의존)
    val = os.getenv(key, default)
    if not val and "streamlit.runtime" in globals():  # 방어
        try:
            import streamlit as st
            val = st.secrets.get(key, default)
        except Exception:
            pass
    return val

def _now() -> str:
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def extract_text_from_file(path: str | Path) -> str:
    p = Path(path); ext = p.suffix.lower()
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
    # 줄바꿈 정리
    text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()

# ========== OpenAI 호출 ==========

@dataclass
class OpenAIConf:
    model: str = "gpt-4o"
    temperature: float = 0.2

def _openai_client():
    # v1 new
    from openai import OpenAI
    api = _get("OPENAI_API_KEY")
    if not api:
        raise RuntimeError("OPENAI_API_KEY 없음")
    return OpenAI(api_key=api)

def _chat(model: str, messages: List[Dict[str, str]], temperature: float = 0.2) -> str:
    cli = _openai_client()
    resp = cli.chat.completions.create(model=model, messages=messages, temperature=temperature)
    return resp.choices[0].message.content.strip()

# ========== Perplexity 호출 ==========

def _get_pplx_key() -> str:
    # 다양한 키 이름을 허용
    for k in ["PERPLEXITY_API_KEY", "PPLX_API_KEY", "PEPLEXITY_API_KEY"]:
        v = _get(k, "")
        if v:
            return v
    return ""

def search_perplexity(query: str, model: str = "sonar", temperature: float = 0.2) -> Tuple[str, List[str]]:
    """
    Perplexity /chat/completions
    - 모델: sonar / sonar-small / sonar-medium / sonar-pro 중 하나
    - 반환: (답변 텍스트, citations URL 리스트)
    """
    api_key = _get_pplx_key()
    if not api_key:
        raise RuntimeError("Perplexity API Key 없음")

    url = "https://api.perplexity.ai/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": model,
        "temperature": temperature,
        "messages": [
            {"role": "system", "content": "You are a precise research assistant. Always cite sources."},
            {"role": "user", "content": query},
        ],
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(f"Perplexity {r.status_code}: {r.text}")
    data = r.json()
    txt = data["choices"][0]["message"]["content"]
    # citations는 응답 상단 또는 message 내에 존재
    urls = []
    if "citations" in data and isinstance(data["citations"], list):
        urls = data["citations"]
    else:
        # 일부 응답 포맷 변형 대비
        try:
            urls = data["choices"][0]["message"].get("citations", []) or []
        except Exception:
            urls = []
    return txt, urls

def search_serpapi(query: str, num_results: int = 5) -> Tuple[str, List[str]]:
    key = _get("SERP_API_KEY") or _get("SERPAPI_API_KEY") or _get("serp_api_key")
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
    return (" ".join(snippets) or "No relevant content."), urls

# ========== 1) RFP → 요청/질문 생성 (OpenAI) ==========

_REQ_SYS = (
    "당신은 EY·맥킨지 수준의 제안서 설계 전문가입니다. "
    "RFP로부터 '제안 요청사항'을 도출하고 각 요청에 대해 "
    "실행질문(실현을 위한 질문) 3개, 벤치마킹질문 3개를 만들어 주세요. "
    "마크다운 금지, 순수 텍스트. 예시 형식을 따르세요."
)
def rfp_requirement_check(user_inputs: dict, conf: OpenAIConf) -> List[str]:
    rfp_text = user_inputs["rfp_text"]
    user_direction = user_inputs.get("user_direction", "")
    keywords = user_inputs.get("keywords", [])
    user_prompt = f"""
[RFP]
{rfp_text}

[고객 방향성]
{user_direction or '없음'}

[강조 키워드] {", ".join(keywords) if keywords else "없음"}

예시 형식
1. 제안 요청사항: (요청 제목)
- 제안 방안: (간략)
- 제안 방안을 실제로 실현 할 수 있기 위한 질문 3개:
(1) ...
(2) ...
(3) ...
- 제안 요청사항의 추가 제안 방안 예시를 알기 위한 질문 3개 (벤치마킹 질문):
(1) ...
(2) ...
(3) ...
"""
    txt = _chat(conf.model, [
        {"role": "system", "content": _REQ_SYS},
        {"role": "user", "content": user_prompt},
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

# ========== 2) Perplexity/SerpAPI로 질문 답변 수집 ==========

def _ctx_q(req_title: str, q: str) -> str:
    req_title, q = (req_title or "").strip(), (q or "").strip()
    return f"{req_title} 관점에서, {q}"

def generate_answer_dict(rfp_dict: Dict[str, Any], prefer: str = "perplexity") -> Dict[str, Any]:
    ans: Dict[str, Any] = {}
    pplx_ok = bool(_get_pplx_key())
    for req_id, sec in rfp_dict.items():
        req_title = sec.get("제안요청사항","")
        ans[req_id] = {"실현질문_답변":[], "실현질문_url":[], "벤치마킹질문_답변":[], "벤치마킹질문_url":[]}

        # 실현질문
        for q in sec.get("실현질문", []):
            cq = _ctx_q(req_title, q)
            if prefer == "perplexity" and pplx_ok:
                a, urls = search_perplexity(cq, model="sonar")
            else:
                a, urls = search_serpapi(cq)
            ans[req_id]["실현질문_답변"].append(a)
            ans[req_id]["실현질문_url"].append(urls)

        # 벤치마킹질문
        for q in sec.get("벤치마킹질문", []):
            cq = _ctx_q(req_title, q)
            if prefer == "perplexity" and pplx_ok:
                a, urls = search_perplexity(cq, model="sonar")
            else:
                a, urls = search_serpapi(cq)
            ans[req_id]["벤치마킹질문_답변"].append(a)
            ans[req_id]["벤치마킹질문_url"].append(urls)
    return ans

# ========== 3) Q/A 기반으로 updated_제안방안 생성 (OpenAI) ==========

def _mk_prompt(mode: str, rfp_text: str, req: str,
               base_plan: str, qas: Optional[str] = None,
               single_bench: Optional[Tuple[str,str]] = None) -> str:
    head = "당신은 EY/맥킨지 수준의 전략 컨설팅 전문가입니다.\n- markdown 금지\n"
    base = f"[RFP 요약]\n{rfp_text[:3000]}\n\n[제안요청사항]\n{req}\n\n[기존 제안방안]\n{base_plan}\n\n"
    if mode == "실현":
        return head + base + "[실현질문·답변]\n" + (qas or "")
    if mode == "벤치" and single_bench:
        q, a = single_bench
        return head + base + f"[벤치마킹 질문]\nQ: {q}\nA: {a}"
    return head + base

def generate_updated_plans_per_question(rfp_dict: Dict[str, Any],
                                        answer: Dict[str, Any],
                                        rfp_text: str,
                                        conf: OpenAIConf) -> pd.DataFrame:
    rows: List[Dict[str,Any]] = []
    for req_id, sec in rfp_dict.items():
        req = sec.get("제안요청사항",""); base = sec.get("제안방안","")
        # 실현 (모아쓰기)
        qs = sec.get("실현질문", []); ans = answer.get(req_id, {}).get("실현질문_답변", [])
        urls_list = answer.get(req_id, {}).get("실현질문_url", [])
        if qs and ans:
            qas = "\n".join([f"{i+1}. Q: {q}\n   A: {a}" for i,(q,a) in enumerate(zip(qs,ans))])
            prompt = _mk_prompt("실현", rfp_text, req, base, qas=qas)
            up = _chat(conf.model, [{"role":"user","content":prompt}], temperature=0.3)
            rows.append({
                "요청 ID": req_id, "제안요청사항": req, "질문유형":"실현질문",
                "질문":"\n\n".join(qs), "답변":"\n\n".join(ans),
                "urls": urls_list, "updated_제안방안": up
            })
        # 벤치(개별)
        bq = sec.get("벤치마킹질문", []); ba = answer.get(req_id, {}).get("벤치마킹질문_답변", [])
        burl = answer.get(req_id, {}).get("벤치마킹질문_url", [])
        for i,(q,a) in enumerate(zip(bq, ba)):
            prompt = _mk_prompt("벤치", rfp_text, req, base, single_bench=(q,a))
            up = _chat(conf.model, [{"role":"user","content":prompt}], temperature=0.3)
            rows.append({
                "요청 ID": req_id, "제안요청사항": req, "질문유형":"벤치마킹질문",
                "질문": q, "답변": a, "urls": burl[i] if i < len(burl) else [],
                "updated_제안방안": up
            })
    return pd.DataFrame(rows)

# ========== 4) updated_df → 슬림 옵션/슬라이드 (OpenAI Responses) ==========

def _responses_client():
    from openai import OpenAI
    api = _get("OPENAI_API_KEY")
    return OpenAI(api_key=api)

def _resp_text(resp) -> str:
    # 다양한 포맷 방어적으로 추출
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
    return json.dumps(resp, ensure_ascii=False)

def _extract_json_block(txt: str) -> Dict[str,Any]:
    m = re.search(r"```json\s*(\{.*?\})\s*```", txt, flags=re.S)
    if m:
        return json.loads(m.group(1))
    s, e = txt.find("{"), txt.rfind("}")
    if s != -1 and e != -1 and e > s:
        return json.loads(txt[s:e+1])
    raise ValueError("JSON 파싱 실패")

def _responses_json(prompt: str, model: str = "gpt-4o-mini") -> Dict[str,Any]:
    cli = _responses_client()
    r = cli.responses.create(model=model, input=prompt)
    return _extract_json_block(_resp_text(r))

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
            f"Updated: {s['updated_plan']}\nURLs:\n" + "\n".join(f"  - {u}" for u in s.get("urls", []))
        )
    schema = {
        "req_id": req_id, "req_title": req_title,
        "overview_slide":{"title":"요청별 옵션 개요","subtitle":"비교",
                          "table":{"columns":["옵션","핵심전략","기간(주)","리스크"],"rows":[["1","","",""]]}},
        "options":[{"option_no":1,"option_title":"", "why_choose":[], "fit_signals":[],
                    "risks":[], "mitigations":[], "timeline":[],
                    "slides":[{"slide_no":1,"title":"","subtitle":"","purpose":"",
                               "key_messages":[],"content_draft":"","paste_blocks":{"body_bulleted":[]},
                               "urls":[]}] }]
    }
    rules = (
        "- JSON만 반환\n"
        "- 각 옵션 6~10장 슬라이드\n"
        "- 수치/사례는 seed URL에서만 인용(없으면 실행 구조 중심)\n"
        "- 붙여넣기용 텍스트 블록(paste_blocks) 채우기"
    )
    return (
        "당신은 컨설팅 문서 설계 전문가입니다.\n"
        f"요청({req_id} {req_title})에 대해 아래 시드 4개를 각각 다른 옵션으로 확장하고 "
        "슬라이드 흐름을 만들어 주세요.\n\n" +
        "\n\n".join(seeds_txt) +
        "\n\n[규칙]\n" + rules + "\n\n[스키마]\n" + json.dumps(schema, ensure_ascii=False)
    )

def _slim_from_updated_df(updated_df: pd.DataFrame, model: str = "gpt-4o-mini") -> pd.DataFrame:
    req_groups = list(updated_df.groupby("요청 ID", sort=False))
    rows: List[Dict[str,Any]] = []
    # 표지
    rows.append({"요청 ID":"COVER","요청 제목":"(표지)","옵션번호":"","슬라이드번호":"1",
                 "제목":"제안서(슬림)","부제목":"옵션/슬라이드 흐름","본문초안":"",
                 "왜_이_옵션":"","적합_시그널":"","리스크":"","완화책":"","타임라인":"","URL":"","옵션대제목":""})
    for req_id, g in req_groups:
        req_title = str(g["제안요청사항"].iloc[0])
        seeds = []
        for i, (_, r) in enumerate(g.head(4).iterrows(), start=1):
            seeds.append({
                "option_no": i,
                "question": str(r["질문"]),
                "answer": str(r["답변"]),
                "updated_plan": str(r.get("updated_제안방안","")),
                "urls": _flatten_urls(r.get("urls", []))
            })
        prompt = _build_deck_prompt(str(req_id), req_title, seeds)
        js = _responses_json(prompt, model=model)

        ov = js.get("overview_slide", {})
        rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":"OVERVIEW","슬라이드번호":"0",
                     "제목":ov.get("title","옵션 개요"),"부제목":ov.get("subtitle",""),
                     "본문초안":ov.get("purpose",""),"왜_이_옵션":"","적합_시그널":"",
                     "리스크":"","완화책":"","타임라인":json.dumps(ov.get("table",{}),ensure_ascii=False),
                     "URL":"","옵션대제목":""})
        for opt in js.get("options", []):
            opt_no = str(opt.get("option_no",""))
            rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":opt_no,"슬라이드번호":"META",
                         "제목":opt.get("option_title",""),"부제목":"(요약)","본문초안":"",
                         "왜_이_옵션":"\n- "+"\n- ".join(opt.get("why_choose",[])),
                         "적합_시그널":"\n- "+"\n- ".join(opt.get("fit_signals",[])),
                         "리스크":"\n- "+"\n- ".join(opt.get("risks",[])),
                         "완화책":"\n- "+"\n- ".join(opt.get("mitigations",[])),
                         "타임라인":json.dumps(opt.get("timeline",[]),ensure_ascii=False),
                         "URL":"","옵션대제목":opt.get("option_title","")})
            for s in opt.get("slides", []):
                rows.append({"요청 ID":req_id,"요청 제목":req_title,"옵션번호":opt_no,
                             "슬라이드번호":str(s.get("slide_no","1")),
                             "제목":s.get("title",""),"부제목":s.get("subtitle",""),
                             "본문초안": (s.get("content_draft","") +
                                        (("\n\n[붙여넣기]\n- "+"\n- ".join(s.get("paste_blocks",{}).get("body_bulleted",[])))
                                         if s.get("paste_blocks") else "")),
                             "왜_이_옵션":"","적합_시그널":"","리스크":"","완화책":"",
                             "타임라인":"","URL":"\n".join(s.get("urls",[])),"옵션대제목":opt.get("option_title","")})
    rows.append({"요청 ID":"CLOSING","요청 제목":"(마무리)","옵션번호":"",
                 "슬라이드번호":"1","제목":"다음 단계","부제목":"",
                 "본문초안":"- 옵션 선택 워크숍\n- 데이터/사전조건 점검\n- 파일럿 범위 합의",
                 "왜_이_옵션":"","적합_시그널":"","리스크":"","완화책":"","타임라인":"","URL":"","옵션대제목":""})
    df = pd.DataFrame(rows)
    for c in ["옵션번호","슬라이드번호"]:
        df[c] = df[c].astype(str)
    return df

# ========== 공개 엔트리: Tab2 버튼이 호출하는 전체 파이프라인 ==========

def build_flows_from_user_inputs(
    rfp_path: str,
    client_name: str,
    user_direction: str,
    notes: str = "",
    model_main: str = "gpt-4o",
) -> pd.DataFrame:
    """
    1) 파일 텍스트 추출
    2) 요청/질문 생성 (OpenAI)
    3) Perplexity(있으면) 또는 SerpAPI로 각 질문 리서치 → Q/A + URL
    4) Q/A로 updated_제안방안 생성 (OpenAI)
    5) updated_df → 슬림 옵션/슬라이드 DF (OpenAI Responses)
    """
    rfp_text = extract_text_from_file(rfp_path)
    # (2) 요청/질문
    user_inputs = {
        "rfp_text": rfp_text,
        "style": "신뢰감 있는",
        "keywords": [],
        "client_name": client_name,
        "proposal_title": "",
        "user_direction": user_direction
    }
    req_lines = rfp_requirement_check(user_inputs, OpenAIConf(model=model_main, temperature=0.2))
    rfp_dict = parse_rfp_response_to_dict(req_lines)

    # (3) 리서치
    answers = generate_answer_dict(rfp_dict, prefer="perplexity")

    # (4) 업데이트 제안
    updated_df = generate_updated_plans_per_question(
        rfp_dict, answers, rfp_text, OpenAIConf(model=model_main, temperature=0.3)
    )
    required = {"요청 ID","제안요청사항","질문유형","질문","답변","urls","updated_제안방안"}
    missing = required - set(updated_df.columns)
    if missing:
        # 이론상 발생X, 안전장치
        for c in missing: updated_df[c] = ""

    # (5) 슬림 플로우
    auto_df = _slim_from_updated_df(updated_df, model="gpt-4o-mini")

    # 스키마 보강(앱 호환)
    need = ["요청 ID","요청 제목","옵션번호","슬라이드번호","제목","부제목","본문초안",
            "왜_이_옵션","적합_시그널","리스크","완화책","타임라인","URL","옵션대제목"]
    for c in need:
        if c not in auto_df.columns: auto_df[c] = ""
    auto_df["슬라이드번호"] = auto_df["슬라이드번호"].astype(str)
    auto_df["옵션번호"] = auto_df["옵션번호"].astype(str)
    return auto_df
