import os, math
from typing import List, Dict, Tuple
from dotenv import load_dotenv
load_dotenv()

import chromadb
from chromadb.config import Settings
from sentence_transformers import SentenceTransformer, CrossEncoder

from rank_bm25 import BM25Okapi
import argparse

from openai import OpenAI

# ---------- 설정 ----------
DB_DIR = "vectordb"
COLLECTION_NAME = "rag_collection"
EMBEDDING_MODEL = os.getenv("EMBEDDING_MODEL", "BAAI/bge-m3")
TOP_K = int(os.getenv("TOP_K", "6"))
USE_MMR = os.getenv("MMR", "False").lower() == "true"
USE_RERANK = os.getenv("RERANK", "False").lower() == "true"

OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

# ---------- 임베딩 & DB ----------
def get_embedder():
    model = SentenceTransformer(EMBEDDING_MODEL)
    model.max_seq_length = 512
    return model

def get_collection():
    client = chromadb.PersistentClient(path=DB_DIR, settings=Settings(allow_reset=False))
    return client.get_collection(COLLECTION_NAME)

# ---------- 검색 ----------
def retrieve(query: str, k: int = TOP_K, mmr: bool = USE_MMR) -> List[Dict]:
    col = get_collection()
    embedder = get_embedder()
    q_emb = embedder.encode([query], normalize_embeddings=True).tolist()[0]

    res = col.query(query_embeddings=[q_emb], n_results=max(k*3 if mmr else k, k))
    docs = res["documents"][0]
    metas = res["metadatas"][0]
    dists = res["distances"][0] if "distances" in res else [0.0]*len(docs)

    items = [{"doc": d, "meta": m, "dist": dist} for d, m, dist in zip(docs, metas, dists)]

    # 간단 MMR(유사도 다양성) — 임베딩 기반 재선택
    if mmr:
        emb_chunks = embedder.encode([it["doc"] for it in items], normalize_embeddings=True)
        selected = []
        cand_idx = set(range(len(items)))
        # greedy
        while cand_idx and len(selected) < k:
            if not selected:
                # 가장 가까운 것부터
                best = min(cand_idx, key=lambda i: items[i]["dist"])
                selected.append(best)
                cand_idx.remove(best)
                continue
            # 다양성 점수 계산: 후보와 이미 선택된 것들 간 최대 유사도 최소화
            def mmr_score(i):
                import numpy as np
                qsim = 1 - items[i]["dist"]  # cosine distance -> similarity approx
                diversity = max(np.dot(emb_chunks[i], emb_chunks[j]) for j in selected)
                lam = 0.75
                return lam*qsim - (1-lam)*diversity
            best = max(cand_idx, key=mmr_score)
            selected.append(best)
            cand_idx.remove(best)
        items = [items[i] for i in selected]
    else:
        items = sorted(items, key=lambda x: x["dist"])[:k]

    return items

# ---------- (옵션) Lexical + Re-Rank ----------
def bm25_mix(query: str, hits: List[Dict], alpha: float = 0.2) -> List[Dict]:
    """ 벡터 히트에서 BM25 점수와 혼합 (alpha는 BM25 가중치) """
    corpus = [h["doc"] for h in hits]
    bm25 = BM25Okapi([c.split() for c in corpus])
    bm_scores = bm25.get_scores(query.split())
    # 거리(dist)는 낮을수록 유리 → 유사도처럼 바꿔서 합산
    import numpy as np
    dist = np.array([h["dist"] for h in hits])
    sim = 1 - (dist / (dist.max() + 1e-9))
    mixed = alpha*bm_scores + (1-alpha)*sim
    order = mixed.argsort()[::-1]
    return [hits[i] for i in order]

def rerank_cross_encoder(query: str, hits: List[Dict], model_name: str = "BAAI/bge-reranker-v2-m3") -> List[Dict]:
    ce = CrossEncoder(model_name)
    pairs = [[query, h["doc"]] for h in hits]
    scores = ce.predict(pairs)
    ranked = sorted(zip(hits, scores), key=lambda x: x[1], reverse=True)
    return [h for h, _ in ranked]

# ---------- 생성 ----------
def build_prompt(query: str, contexts: List[Dict]) -> List[Dict]:
    context_block = "\n\n".join(
        [f"[{i+1}] SOURCE: {c['meta'].get('source')} (chunk {c['meta'].get('chunk_idx')})\n{c['doc']}"
         for i, c in enumerate(contexts)]
    )
    system = (
        "당신은 정확한 RAG 비서입니다. 제공된 '컨텍스트'만 근거로 한국어로 답하세요. "
        "모르면 모른다고 말하세요. 반드시 근거가 된 출처를 인덱스 번호로 함께 표기하세요."
    )
    user = (
        f"질문:\n{query}\n\n"
        f"컨텍스트(참조용):\n{context_block}\n\n"
        "요구사항:\n- 컨텍스트 범위를 벗어난 추측 금지\n- 핵심 요약 → 근거 표기 [1], [2]...\n"
    )
    return [
        {"role":"system", "content": system},
        {"role":"user", "content": user}
    ]

def generate_answer(query: str, contexts: List[Dict]) -> str:
    client = OpenAI()
    messages = build_prompt(query, contexts)
    resp = client.chat.completions.create(
        model=OPENAI_MODEL,
        messages=messages,
        temperature=0.2,
    )
    return resp.choices[0].message.content

# ---------- CLI ----------
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--q", required=True, help="질문 텍스트")
    parser.add_argument("--k", type=int, default=TOP_K, help="검색 개수")
    parser.add_argument("--lexical", action="store_true", help="BM25 가중 혼합 사용")
    parser.add_argument("--rerank", action="store_true", help="Cross-Encoder 재랭킹 사용")
    args = parser.parse_args()

    hits = retrieve(args.q, k=args.k, mmr=USE_MMR)
    if args.lexical:
        hits = bm25_mix(args.q, hits)
    if args.rerank or USE_RERANK:
        hits = rerank_cross_encoder(args.q, hits)

    answer = generate_answer(args.q, hits[:args.k])

    print("\n=== ANSWER ===\n")
    print(answer.strip())
    print("\n=== SOURCES ===")
    for i, h in enumerate(hits[:args.k], 1):
        src = h["meta"].get("source"); idx = h["meta"].get("chunk_idx")
        print(f"[{i}] {src} (chunk {idx})")

if __name__ == "__main__":
    main()
