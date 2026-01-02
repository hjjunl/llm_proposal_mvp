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

### Demo Screen
<img width="1613" height="841" alt="화면1" src="https://github.com/user-attachments/assets/ba39b05f-2888-4868-9500-e5a76d0016cb" />

<img width="1202" height="801" alt="화면2" src="https://github.com/user-attachments/assets/86557dc9-8759-4c56-ab24-4c78f800a443" />

### Result
<img width="1186" height="787" alt="image" src="https://github.com/user-attachments/assets/09316642-b9ac-466c-8e65-9ca1ffa96661" />
<img width="1136" height="836" alt="image" src="https://github.com/user-attachments/assets/3ccac58b-ffbf-43e7-928e-85aee5e93213" />

---

## 📥 Supported Input Formats

- 📄 PDF  
- 🖼️ PowerPoint (PPT / PPTX)  
- 📝 Word (DOCX)  

No pre-formatting or manual structuring is required.

---

## 🚀 Getting Started

### Prerequisites

- **Python**: 3.11 or higher (3.13 recommended)
- **API Keys**: OpenAI API key (required), Perplexity API key (required)

### Installation

#### Option 1: Using pip (Recommended)

1. **Clone the repository**
   ```bash
   git clone https://github.com/hjjunl/llm_proposal_mvp.git
   cd llm_proposal_mvp
   ```

2. **Install dependencies**
   ```bash
   cd proposal_ai_agent
   pip install -r requirements.txt
   ```

#### Option 2: Using Conda

1. **Create conda environment**
   ```bash
   conda env create -f environment.yml
   conda activate llms
   ```

2. **Install additional dependencies**
   ```bash
   cd proposal_ai_agent
   pip install -r requirements.txt
   ```

### Environment Setup

1. **Create `.env` file** in the project root directory:
   ```bash
   # In the root directory (llm_proposal_mvp/)
   touch .env
   ```

2. **Add your API keys** to the `.env` file:
   ```env
   # Required
   OPENAI_API_KEY=sk-your-openai-api-key-here

   # Recommended (for research functionality)
   PERPLEXITY_API_KEY=pplx-your-perplexity-api-key-here
   # Alternative names also supported:
   # PPLX_API_KEY=pplx-your-key
   # PEPLEXITY_API_KEY=pplx-your-key

   # Optional (fallback for research)
   SERP_API_KEY=your-serp-api-key-here
   ANTHROPIC_API_KEY=sk-ant-your-anthropic-key
   GOOGLE_API_KEY=AIzaSy-your-google-key
   ```

   > **Note**: At minimum, you need `OPENAI_API_KEY`. `PERPLEXITY_API_KEY` is highly recommended for research features.

### Running the Application

#### Method 1: Streamlit Web App (Recommended)

1. **Navigate to the proposal_ai_agent directory**
   ```bash
   cd proposal_ai_agent
   ```

2. **Run Streamlit**
   ```bash
   streamlit run app.py
   ```

3. **Access the app**
   - The app will open automatically in your browser
   - Default URL: `http://localhost:8501`

4. **Usage**
   - **Page 1 (RFP Upload)**: Upload RFP documents and enter client information
   - **Page 2 (Client History)**: View past proposals and project history

#### Method 2: Command Line Interface

For batch processing or automation:

```bash
cd proposal_ai_agent
python run_pipeline_once.py \
    --rfp "DB/RFP/sample.docx" \
    --client "Client Name" \
    --direction "Project direction and focus areas"
```

**Arguments:**
- `--rfp`: Path to RFP file (PDF, DOCX, or PPTX)
- `--client`: Client/company name
- `--direction`: Project direction or focus areas (optional)

**Output:**
- Results saved to: `DB/proposal_result/manual_run_YYYYmmdd_HHMMSS/auto_df.xlsx`

### Project Structure

```
llm_proposal_mvp/
├── proposal_ai_agent/          # Main application directory
│   ├── app.py                  # Streamlit main app
│   ├── pages/                  # Streamlit pages
│   │   ├── 01_RFP_Upload.py   # RFP upload page
│   │   └── 02_Client_History.py # Client history page
│   ├── pipeline/               # Core processing modules
│   │   ├── analyze_rfp.py     # RFP analysis
│   │   ├── inputs2flows.py    # Flow generation
│   │   └── rfp2proposal.py    # Main pipeline
│   ├── utils/                 # Utility functions
│   ├── DB/                    # Data storage
│   │   ├── RFP/               # Uploaded RFP files
│   │   ├── proposal_result/   # Generated proposals
│   │   └── clients.db         # SQLite database
│   └── requirements.txt       # Python dependencies
├── .env                       # Environment variables (create this)
├── .gitignore
├── README.md
└── requirements.txt          # Root-level dependencies
```

### Troubleshooting

#### Common Issues

1. **ModuleNotFoundError**
   - Ensure you're in the `proposal_ai_agent` directory when running commands
   - Try: `pip install -r requirements.txt` again

2. **API Key Errors**
   - Verify your `.env` file is in the project root
   - Check that API keys are correctly formatted (no quotes, no spaces)
   - Restart your terminal/IDE after creating `.env`

3. **Streamlit Port Already in Use**
   ```bash
   streamlit run app.py --server.port 8502
   ```

4. **Database Errors**
   - The SQLite database (`DB/clients.db`) is created automatically on first run
   - Ensure the `DB/` directory exists and is writable

### Next Steps

1. Upload your first RFP document through the Streamlit interface
2. Enter client information and project direction
3. Review the generated PPT outline
4. Export results as Excel or PDF

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

