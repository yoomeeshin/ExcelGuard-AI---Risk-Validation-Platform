# ExcelGuard AI

![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)

**AI-powered validation platform for financial risk management** - automates Excel workbook validation to identify budget overruns, cash flow issues, and resource allocation problems. Reduces analysis time from 40 hours to 2 minutes.

**Built for Consulting Advisory** where analysts manually validated project financials across 200+ engagements. Due to data confidentiality requirements, cloud-based solutions (OpenAI, external APIs) could not be used as sensitive client financial data cannot be transmitted to external servers. This necessitated a fully local solution with all processing occurring on-premise.

---

## Technical Overview

**Local Processing**: Flask 3.1 REST API + Pandas 3.0 DataFrames + openpyxl 3.1 for cell-level Excel manipulation. Zero external API calls.

**Multi-Agent Architecture**: Specialized agents (SupervisorAgent, RuleInterpreterAgent, SmartRuleInterpreter, RowValidatorAgent) coordinated via ThreadPoolExecutor for parallel execution - validates 500K+ cells in <20 seconds.

**Statistical Rule Suggestions**: IQR outlier detection, regex pattern matching (email/phone/SSN), and cardinality analysis to auto-recommend validation rules.  

**Smart Validation**: Cross-sheet logic (gap detection, conditional sums, temporal patterns), 10+ operators (>, >=, regex, contains, date_future), accounting format handling `(123.45)` → `-123.45`.

**Source Attribution**: Every violation includes cell address, rule ID, and suggested fix for audit trails.

---

## Quick Start

```bash
# Local
git clone https://github.com/yourusername/excelguard-ai.git
cd excelguard-ai
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
python app.py  # http://localhost:5001

# Docker
docker-compose up -d
```

---

## Usage

**Web Interface**: Upload Excel workbook + validation rules → Click "Start Validation" → Download report with violations

**API**:
```bash
curl -X POST http://localhost:5001/api/validate \
  -H "Content-Type: application/json" \
  -d '{"data_filename": "data.xlsx", "rules_filename": "rules.xlsx"}'
```
---

## Tech Stack

```python
Flask==3.1.0           # REST API
pandas==3.0.0          # DataFrame processing
openpyxl==3.1.2        # Excel manipulation
flask-cors==4.0.0      # CORS support
```

---
