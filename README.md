# Smart-QA

Smart-QA is a full-stack QA automation application designed to help QA teams convert business requirements into professionally formatted, Excel-ready test cases. The platform supports uploading requirement files, configuring an LLM endpoint and API key from the UI, generating structured test cases using predefined QA rules, reviewing results automatically, and downloading the final workbook.

## Overview

This repository contains both the backend services and the web frontend for the Smart-QA application:

- Python/ FastAPI backend for generation, validation, and review workflows
- React + Vite frontend for user interaction and report management
- Prompt pack files used to enforce QA prompt rules and output structure

## Features

- Upload requirements in Excel, DOCX, or image formats
- Configure LLM endpoint, API key, and model from the UI
- Generate structured, Excel-ready test cases from uploaded requirement artifacts
- Apply review rules and validation logic to produced test cases
- Download generated output and review reports
- Use predefined QA prompt packs for naming, formatting, and review standards

## Repository Structure

```text
Smart-QA/
├── .gitignore
├── QA_Agent_Prompt_Pack/
│   ├── examples.md
│   ├── excel_rules.md
│   ├── naming_rules.md
│   ├── output_schema.json
│   ├── review_rules.md
│   ├── system_prompt.md
│   ├── testcase_rules.md
│   └── QA_Agent_Prompt_Pack.zip
├── qa_agent/
│   ├── Smart_QA/
│   │   ├── src/
│   │   ├── public/
│   │   ├── package.json
│   │   ├── package-lock.json
│   │   ├── vite.config.js
│   │   └── README.md
│   ├── qa_api.py
│   ├── qa_agent_2.py
│   ├── test_case_reviewer.py
│   ├── watch_excel_trigger.py
│   ├── watch_excel_watchdog.py
│   ├── qa_agent_config.example.json
│   ├── TestCaseReview_CheckList.xlsx
│   └── temp_uploads/ (runtime-generated)
└── README.md
```

## Tech Stack

- Backend: Python, FastAPI
- Frontend: React, Vite, Bootstrap
- Data processing: Excel/Document workflows and QA review automation
- LLM integration: configurable external LLM endpoint

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/rsaiteja-ai/Smart-QA.git
cd Smart-QA
```

### 2. Set up the backend

```bash
cd qa_agent
python -m venv .venv
source .venv/bin/activate  # On Windows use .venv\Scripts\activate
```

Install the required Python dependencies for the project in your environment, then start the API:

```bash
uvicorn qa_api:app --reload --host 0.0.0.0 --port 8000
```

The API documentation will be available at:

- http://localhost:8000/docs

### 3. Set up the frontend

```bash
cd qa_agent/Smart_QA
npm install
npm run dev
```

Open the frontend in the browser at:

- http://localhost:5173

## Configuration

The project includes an example config file at:

- `qa_agent/qa_agent_config.example.json`

Use this as a template to configure your LLM connection details such as endpoint, API key, and model.

## Usage

1. Launch the backend and frontend.
2. Open the Smart-QA web application.
3. Upload a requirement file or artifact.
4. Enter or configure the LLM endpoint and API key.
5. Generate the test cases.
6. Review the generated quality checks and download the final Excel output.

## Prompt Pack

The repository contains a prompt pack in `QA_Agent_Prompt_Pack/` to guide the test-case generation and review process. These files define:

- naming conventions
- output schema
- Excel formatting rules
- review checklist behavior
- system prompt and example patterns

## Notes

- The backend writes generated files and review artifacts into the runtime `qa_agent/temp_uploads/` folder.
- The nested frontend project has its own `README.md` in `qa_agent/Smart_QA/` if you want a frontend-specific quick start.

## License

This repository does not currently declare a license. If you plan to share or distribute the project publicly, add a license file that matches your intended usage.

## Contact

For questions, enhancements, or collaboration, please reach out through the repository owner or project maintainers on GitHub.
