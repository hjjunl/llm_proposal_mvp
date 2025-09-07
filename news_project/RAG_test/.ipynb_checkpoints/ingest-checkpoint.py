import os, glob, uuid
from typing import List, Dict, Tuple
from dotenv import load_dotenv
load_dotenv()

import chromadb
from chromadb.config import Settings
from sentence_transformers import SentenceTransformer
from pypdf import PdfReader
import tiktoken

# ---------- 설정 ----------
DATA_DIR = "data"
DB_DIR = "vectordb"
COLLECTION_NAME = "rag_collection"

EMBEDDING_MODEL = os.getenv("EMBEDDING_MODEL", "BAAI/bge-m3")
CHUNK_TOKENS = 400
CHUNK_OVERLAP = 80

# ---------- 유틸 ----------
def read_text_file(path: str) -> str:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()

def read_pdf_file(path: str) -> str:
    reader = PdfReader(path)
    texts = []
    for page in reader.pages:
        txt = page.extract_text() or ""
        texts.append(txt)
    return "\n".join(texts)

def load_documents(data_dir: str) -> List[Tuple[str, str]]:
    docs = []
    patterns = ["**/*.txt", "**/*.md", "**/*.pdf"]
    for pat in patterns:
        for p in glob.glob(os.path.join(data_dir, pat), recursive=True):
            ext = os.path.splitext(p)[1].lower()
            if ext in [".txt", ".md"]:
                text = read_text_file(p)
            elif ext == ".pdf":
                text = read_pdf_file(p)
            else:
                continue
            if text.strip():
                docs.append((p, text))
    return docs

def token_chunk(text: str, max_tokens=CHUNK_TOKENS, overlap=CHUNK_OVERLAP) -> List[str]:
    enc = tiktoken.get_encoding("cl100k_base")
    toks = enc.encode(text)
    chunks = []
    start = 0
    while start < len(toks):
        end = min(start + max_tokens, len(toks))
        chunk = enc.decode(toks[start:end])
        chunks.append(chunk)
        if end == len(toks):
            break
        start = max(0, end - overlap)
    return chunks

# ---------- 임베딩 ----------
def get_embedder():
    model = SentenceTransformer(EMBEDDING_MODEL)
    # bge/e5 류는 보통 normalize 추천
    model.max_seq_length = 512
    return model

# ---------- 메인 ----------
def main():
    os.makedirs(DB_DIR, exist_ok=True)
    client = chromadb.PersistentClient(path=DB_DIR, settings=Settings(allow_reset=False))
    colls = [c.name for c in client.list_collections()]
    if COLLECTION_NAME in colls:
        collection = client.get_collection(COLLECTION_NAME)
    else:
        collection = client.create_collection(COLLECTION_NAME, metadata={"hnsw:space": "cosine"})

    embedder = get_embedder()

    docs = load_documents(DATA_DIR)
    if not docs:
        print(f"[INGEST] No documents found in ./{DATA_DIR}. Add files and rerun.")
        return

    ids, texts, metadatas = [], [], []
    for path, full_text in docs:
        chunks = token_chunk(full_text)
        for i, ch in enumerate(chunks):
            ids.append(str(uuid.uuid4()))
            texts.append(ch)
            metadatas.append({"source": path, "chunk_idx": i})

    print(f"[INGEST] Embedding {len(texts)} chunks with {EMBEDDING_MODEL} ...")
    embs = embedder.encode(texts, normalize_embeddings=True, show_progress_bar=True).tolist()

    print(f"[INGEST] Upserting to Chroma ({COLLECTION_NAME}) ...")
    collection.upsert(ids=ids, embeddings=embs, metadatas=metadatas, documents=texts)

    print(f"[DONE] {len(docs)} files, {len(texts)} chunks indexed at {DB_DIR}/")

if __name__ == "__main__":
    main()
