# ExcelGuard AI - Quick Start Guide

## 🚀 Getting Started in 5 Minutes

### Option 1: Local Development (Recommended for Testing)

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/excelguard-ai.git
cd excelguard-ai

# 2. Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the application
python app.py

# 5. Open browser
# Navigate to: http://localhost:5001
```

### Option 2: Docker (Recommended for Production)

```bash
# 1. Clone the repository
git clone https://github.com/yourusername/excelguard-ai.git
cd excelguard-ai

# 2. Build and run with Docker Compose
docker-compose up -d

# 3. Access the application
# Navigate to: http://localhost:5001

# 4. View logs
docker-compose logs -f

# 5. Stop the application
docker-compose down
```

## 📝 Basic Usage

### 1. Download Rule Template
- Click "Download Rule Template" button
- Opens Excel file with example rules
- Customize rules for your use case

### 2. Upload Files
- **Data File**: Your Excel workbook to validate
- **Rules File**: Excel file with validation rules

### 3. Start Validation
- Click "Start Validation"
- Wait for processing (typically 1-10 seconds)
- Download detailed report with violations

### 4. AI Rule Suggestions (New!)
- Upload your data file
- Click "Suggest Rules with AI"
- System analyzes your data and suggests validation rules
- Download suggested rules as Excel file
- Review and customize before using

## 🎯 Example Use Cases

### Financial Budget Validation
```
1. Download "Finance - Budget Reconciliation" template
2. Upload your budget Excel file
3. System validates:
   - Revenue/expense logic
   - Sum reconciliations
   - Date continuity
   - Negative value checks
```

### Healthcare Claims Processing
```
1. Download "Healthcare - Claims" template
2. Upload claims workbook
3. System validates:
   - ICD code formats
   - Date ranges
   - Amount validations
   - Required field checks
```

## 🔧 Troubleshooting

### Common Issues

**Issue**: "ModuleNotFoundError: No module named 'openpyxl'"
```bash
# Solution: Reinstall dependencies
pip install -r requirements.txt
```

**Issue**: "Port 5001 already in use"
```bash
# Solution: Change port in app.py
# Line 365: app.run(debug=True, host='0.0.0.0', port=5002)
```

**Issue**: Docker container won't start
```bash
# Solution: Check logs
docker-compose logs

# Rebuild image
docker-compose down
docker-compose build --no-cache
docker-compose up -d
```

## 📊 Performance Tips

### For Large Files (100K+ cells)
- Use Docker for better memory management
- Close unnecessary applications
- Consider splitting into multiple workbooks

### For Complex Rules (50+ rules)
- System automatically uses parallel processing
- Typical processing time: 5-20 seconds

## 🤝 Getting Help

- **Issues**: Open issue on GitHub
- **Questions**: Check [FAQ](docs/FAQ.md)
- **Email**: your.email@example.com

## 🎓 Next Steps

1. **Read the full documentation**: [README.md](README.md)
2. **Explore example rules**: See `templates/` folder
3. **Try AI rule suggestion**: Upload a sample file
4. **Customize for your workflow**: Create industry-specific templates

---

**Note**: This is a development tool. Always verify automated validations with domain expertise before making critical business decisions.
