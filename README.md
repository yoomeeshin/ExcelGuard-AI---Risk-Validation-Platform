# ExcelGuard AI

![Python Version](https://img.shields.io/badge/python-3.10+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

An **AI-powered validation platform** for financial risk management. Automates Excel workbook validation to identify budget overruns, cash flow issues, and resource allocation problems - reducing analysis time from 40 hours to 2 minutes.

**Built for Big 4 Risk Advisory** where analysts manually validated project financials across 200+ engagements. Due to data confidentiality requirements, cloud-based solutions (OpenAI, external APIs) could not be used as sensitive client financial data cannot be transmitted to external servers. This necessitated a fully local solution with all processing occurring on-premise.

---

## Key Features

- **Multi-Agent Architecture**: Specialized validation agents (Rule Interpreter, Smart Validator, Anomaly Detector) working in parallel
- **AI Rule Suggestions**: Analyzes data patterns and recommends validation rules automatically using statistical methods (IQR, pattern recognition) - no external API calls required
- **100% Local Processing**: All validation runs on-premise; no data transmitted to cloud services or external APIs (OpenAI, etc.) ensuring client confidentiality
- **Fast**: Validates 500K+ cells in under 20 seconds using parallel processing
- **Smart Validation**: Cross-sheet logic, temporal patterns, conditional sums, gap detection

---

## Technology Stack

| Component | Technology |
|-----------|-----------|
| **Backend** | Flask 3.1+ |
| **Data Processing** | Pandas 3.0+, openpyxl 3.1+ |
| **Concurrency** | ThreadPoolExecutor (5-8x speedup) |
| **AI Engine** | Statistical analysis (IQR), Pattern recognition |
| **Deployment** | Docker, docker-compose |

---

## Quick Start

```bash
# Clone and install
git clone https://github.com/yourusername/excelguard-ai.git
cd excelguard-ai
python3 -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Run
python app.py
# Visit http://localhost:5001
```

**Docker:**
```bash
docker-compose up -d
```

---

## Usage

1. Upload your Excel workbook (data file)
2. Upload validation rules file
3. Click "Start Validation"
4. Download detailed report with violations and suggested fixes

**API Integration:**
```bash
curl -X POST http://localhost:5001/api/validate \
  -H "Content-Type: application/json" \
  -d '{"data_filename": "/path/to/data.xlsx", "rules_filename": "/path/to/rules.xlsx"}'
```
---
