# Contract Extraction & Calculation Tool

A Flask-based web application for extracting contract terms and performing incentive calculations from Excel data and Word contracts.

## Features

- **Contract Term Extraction**: Extract CHD rules, BIR tables, NACV tiers, and general terms from Word documents
- **CV Tier Calculations**: Calculate incentives based on CV Tier/Lo-ROC with BIR basis points
- **NACV Calculations**: Calculate NACV from Bus/AR/Writeoff data with optional CHD performance adjustments
- **Auto-Detection**: Automatically detects calculation type from uploaded files
- **Enhanced Processing**: Uses fuzzy matching, NLP, and table structure recognition

## Tech Stack

- **Backend**: Flask (Python)
- **Document Processing**: python-docx, camelot-py, tabula-py
- **Data Analysis**: pandas, numpy, openpyxl
- **NLP**: spaCy, NLTK, fuzzywuzzy
- **Production Server**: Gunicorn

## Local Development

### Prerequisites

- Python 3.11+
- pip

### Installation

1. Clone the repository:
```bash
git clone <your-repo-url>
cd "Contracts Extraction Calculation Tool - Copy"
```

2. Create virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
python -m spacy download en_core_web_sm
```

4. Download NLTK data:
```bash
python -c "import nltk; nltk.download('punkt'); nltk.download('stopwords'); nltk.download('averaged_perceptron_tagger')"
```

5. Run the application:
```bash
python app.py
```

6. Open browser to `http://localhost:5000`

## Deployment to Render

### Prerequisites

- GitHub account
- Render account (free tier available)

### Steps

1. **Push code to GitHub**:
```bash
git init
git add .
git commit -m "Initial commit - Contract Calculator"
git branch -M main
git remote add origin <your-github-repo-url>
git push -u origin main
```

2. **Deploy on Render**:
   - Go to [Render Dashboard](https://dashboard.render.com/)
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Render will auto-detect the `render.yaml` configuration
   - Click "Apply" to deploy

3. **Configuration** (handled automatically by render.yaml):
   - Build Command: Installs dependencies and downloads NLP models
   - Start Command: `gunicorn app:app`
   - Environment: Python 3.11
   - Persistent Disk: 1GB for uploads

4. **Environment Variables** (auto-generated):
   - `FLASK_SECRET_KEY`: Automatically generated secure key
   - `FLASK_ENV`: Set to `production`

### Manual Render Setup (Alternative)

If you prefer manual setup instead of render.yaml:

1. Create new Web Service on Render
2. Connect GitHub repository
3. Configure:
   - **Build Command**:
     ```
     pip install -r requirements.txt && python -m spacy download en_core_web_sm && python -c 'import nltk; nltk.download("punkt"); nltk.download("stopwords"); nltk.download("averaged_perceptron_tagger")'
     ```
   - **Start Command**: `gunicorn app:app`
   - **Environment Variables**:
     - `FLASK_ENV` = `production`
     - `FLASK_SECRET_KEY` = (generate random string)

## Usage

### Extract Contract Terms

1. Select "Extract Terms" mode
2. Upload a Word contract (.docx)
3. View extracted CHD rules, BIR tables, and other terms

### Perform Calculations

1. Select "Perform Calculations" mode
2. Upload Excel data file and Word contract
3. Click "Analyze Files" for auto-detection
4. Review detected calculation type (CV Tier or NACV)
5. Optionally enter CHD adjustment value
6. Click "Process" to see results

## File Formats

### Excel Files (.xlsx)
- **CV Tier**: Should contain rows with "CV Tier" or "Lo-ROC" labels
- **NACV**: Should contain "Bus Unadjusted", "AR 180", "Writeoff" rows

### Contract Files (.docx)
- Should contain BIR tables with NACV ranges and basis points
- CHD rules with threshold days and adjustment rates
- NACV tier tables with ranges and rates

## Project Structure

```
.
├── app.py                              # Main Flask application
├── enhanced_document_processor.py      # Advanced document processing
├── requirements.txt                    # Python dependencies
├── render.yaml                         # Render deployment config
├── templates/
│   ├── index.html                     # Main upload interface
│   ├── result_cv_tier.html            # CV Tier results
│   ├── result_nacv.html               # NACV results
│   └── result_extract_terms.html      # Term extraction results
├── static/
│   └── style.css                      # Styling
└── uploads/                           # File upload directory
```

## Security Notes

- Max file size: 16MB
- Allowed file types: .xlsx, .xls, .xlsm, .xlsb, .docx
- Files are securely named using `secure_filename()`
- Uploaded files stored in `uploads/` directory

## Troubleshooting

### Build Fails on Render

- Check that all dependencies in `requirements.txt` are compatible
- Verify Python version (3.11) is specified in render.yaml
- Check build logs for specific error messages

### NLP Models Not Found

- Ensure build command downloads spaCy and NLTK data
- Check that `en_core_web_sm` is being downloaded during build

### File Upload Errors

- Verify persistent disk is mounted at `/opt/render/project/src/uploads`
- Check file size limits (16MB default)

## License

MIT

## Support

For issues or questions, please open an issue in the GitHub repository.
