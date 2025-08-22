# CV Incentive Tool

A small Flask web app that:
1) Reads your *calculation Excel* and totals each **CV Tier** and **Lo‑ROC** row (no hardcoding).
2) Parses your *contract .docx* to extract the **BIR (basis points)** table (no hardcoding).
3) Applies the respective BIR to each tier total and reports **Total Gross Incentive**.

> This is built fresh for your workflow. We used patterns similar to your prior reference app where it helped, but the internals are adapted for your current spec.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000 and upload:
- **Excel**: your input calculation workbook (e.g., `input file 2.xlsx`)
- **Word**: your contract (e.g., `Sample 2 contract.docx`)

## How it works

### Excel parsing (nothing hardcoded)
- Loads the **first sheet** as raw rows (no header inference).
- Finds rows whose first column contains **“CV Tier”** or **“Lo‑ROC”** (case/spacing/variants tolerated).
- Sums **all numeric cells** across each found row. Any month layout/merged cells are handled by coercing to numeric.

### Contract parsing (BIR table)
- Looks for any table whose header mentions **NACV/CC‑NACV/Charge Volume** and **BIR/bp/basis points**.
- Reads each row as a **range in USD millions** (e.g., `0–9.9`, `10.0–24.9`, `75.0+`) and a **bps** value.
- It then chooses the band for a given tier total (after converting the total to **USD millions**) and computes:
  \n
  `adjustment = total × (bps / 10,000)`

### Output
- **Table 1**: Totals for each CV Tier / Lo‑ROC (Excel → sums)
- **Table 2**: BIR table (from contract)
- **Table 3**: Per‑tier adjustments + **Total Gross Incentive** (sum of adjustments)

## Notes
- Currency is treated as USD for band selection. If your contract uses another currency, adjust upstream or extend the parser.
- This phase computes **client‑level gross incentive** by band *per category* as requested. CHD/MR/Credit Loss integration can be added next.
- Robust to metadata rows and merged headers.


### CHD (new in this version)
- The app extracts **CHD Benchmark** from the contract (.docx) automatically.
- You can optionally enter your **CHD Performance** value on the upload form.
- The results page displays **Benchmark**, **Your CHD**, and **Difference (CHD − Benchmark)** only (no incentive calc yet).
