# 📄 Proposal AI Agent  
### RFP → Research → PPT Outline Automation System

Proposal AI Agent is an AI-driven system that reads RFP documents (PDF, PPT, Word), analyzes their requirements, conducts supporting research, and generates a **logic-first, evidence-based PPT outline** for proposal creation.

This project focuses on automating **thinking and structure**, not slide design.  
It supports consultants, agencies, and strategy teams by accelerating the most time-consuming intellectual steps of proposal work.

---

## 🎯 What This Project Does

This system answers a core question:

> “Given this RFP,  
> what questions must the proposal answer,  
> what evidence should support those answers,  
> and how should the PPT be structured?”

### ✅ What it DOES
- Reads real-world RFP documents (PDF / PPT / DOCX)
- Extracts and analyzes RFP requirements
- Generates key proposal questions
- Conducts external research per question
- Produces a **PPT structure outline** with evidence-based reasoning

### ❌ What it does NOT do
- Does NOT generate final PPT files
- Does NOT design slides or visuals
- Does NOT replace human review or judgment

---

## 📥 Supported Input Formats

- 📄 PDF  
- 🖼️ PowerPoint (PPT / PPTX)  
- 📝 Word (DOCX)  

No pre-formatting or manual structuring is required.

---

## 🧠 Core Concept

> **This is not a “proposal writing AI”.  
It is a “proposal thinking automation system”.**

The system mirrors how experienced consultants work:

1. Read and understand the RFP  
2. Identify what must be answered  
3. Research facts, trends, and benchmarks  
4. Decide what slides are needed and why  

---

## 🔄 End-to-End Workflow


---

## 🧩 Step-by-Step Breakdown

### 1️⃣ RFP Ingestion & Text Extraction
- Extracts raw text from uploaded documents
- Handles long, unstructured enterprise RFPs
- Preserves contextual structure (sections, clauses)

---

### 2️⃣ Requirement Decomposition
The system identifies:
- Explicit requirements
- Implicit expectations
- Evaluation criteria
- Mandatory response areas

**Output:**  
A structured list of questions the proposal must answer.

---

### 3️⃣ Question-Based Research
For each question, the system:
- Generates targeted research queries
- Performs external research via Perplexity APIs
- Collects facts, statistics, trends, and examples

---

### 4️⃣ Evidence-Centered Reasoning
Instead of writing slides directly, the system:
- Links research results to specific questions
- Filters for relevance and logical support
- Organizes insights into reasoning blocks

---

### 5️⃣ PPT Outline Generation (Final Output)
The system produces a **PPT outline blueprint**, including:
- Recommended slide sequence
- Purpose of each slide
- Key messages per slide
- Supporting evidence references
- Logical flow across slides

This outline is designed to be reviewed and finalized by humans.

---

## 📦 Output Format

The output is a **structured, human-readable outline**, typically provided as:

- 📊 Table / DataFrame
- 🧩 Slide-by-slide structure
- 🔗 Each slide mapped to:
  - RFP requirement
  - Research-backed evidence
  - Intended message

This format supports easy review and downstream use.

---

## 🏗️ System Architecture (Conceptual)

Each layer is separated to ensure:
- Maintainability
- Explainability
- Extensibility

---

## 💡 Why This Project Matters

Most AI proposal tools attempt to skip reasoning and jump directly to writing.

This project takes the opposite approach:

> **Automate thinking first,  
so humans can write, design, and decide better.**

It demonstrates how LLMs can be used as **structured reasoning systems**, not just text generators.

---

## 🚧 Project Status

- ✅ RFP ingestion (PDF / PPT / DOCX)
- ✅ Requirement analysis & question generation
- ✅ Research-backed reasoning pipeline
- ✅ PPT outline generation
- ⏳ Agent orchestration & workflow expansion planned

---

## 🚀 Future Extensions

- LangGraph-based agent orchestration
- Reviewer / approval workflow
- Slide template mapping
- SaaS multi-tenant deployment
- Cost-aware research caching

---

## 👤 Author

**HyunJun Lee**  
Technology Consultant & AI Automation Builder  

📫 Email: hyunjun960214@gmail.com  
🌐 LinkedIn: https://www.linkedin.com/in/hyunjun-lee-a37448212/

