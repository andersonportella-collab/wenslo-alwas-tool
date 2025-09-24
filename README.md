# wenslo-alwas-tool
Multicriteria tool based on WENSLO + ALWAS methods for decision support
Tool in **Streamlit / Python** that implements the multicriteria methodology **WENSLO** (weights) and **ALWAS** (integrated rankings), with support for validation by accumulation (Table V of the article) and generation of Excel reports and correlation heatmaps.

Based on the article:
> *A Novel WENSLO and ALWAS Multicriteria Methodology and Its Application to Green Growth Performance Evaluation* (DOI: 10.1109/TEM.2023.3321697)

---

## Repository content
- `app.py` — Streamlit application (web interface).
- `metodo.pdf` — Manual/method description (available in repositor or by download button in the interface).
- `requirements.txt` — Python dependencies.
- `README.md` — This file.

---

## Requirements
- Python 3.8+ (recommended 3.10 / 3.11)
- pip
- System with display (for matplotlib) – for cloud deployment (Streamlit Cloud).

Dependencies (also in the 'requirements.txt'):
- streamlit
- pandas
- numpy
- matplotlib
- openpyxl
- XlsxWriter

---

## Installation (local)

**1. Clone the repository:**
```bash
git clone https://github.com/andersonportella-collab/wenslo-alwas-tool.git
cd wenslo-alwas-tool


**2. Create and activate a virtual environment**

Windows (cmd)
python -m venv .venv
.venv\Scripts\activate

Windows (PowerShell)
python -m venv .venv
.venv\Scripts\Activate.ps1

macOS / Linux
python3 -m venv .venv
source .venv/bin/activate

**3. Installing dependencies**
pip install -r requirements.txt


### Section `Usage`
```markdown
## Uso

**Run the application locally**
```bash
streamlit run app.py

Open the browser in http://localhost:8501 (or the address that Streamlit indicates).

Excel file format (template)
The app offers a button to download the template_wenslo_alwas.xlsx. The format expected by the upload is:
•	Row 1: header with criteria names (the first cell A1 is empty).
•	Row 2: criterion type corresponding to each column from column B—MAX or MIN values only.
•	Row 3+: Alternatives — column A contains the name of the alternative (e.g., Canada), the following columns contain numeric values.

        |  C11  |  C12  |  C13
--------|-------|-------|------
        |  MAX  |  MIN  |  MAX
Canada  | 523.19|5727.67| 7.57
France  | 258.23|1237.01| 3.23
...

Once you have uploaded the file correctly, click Calculate to run the process

Features

Load decision matrix via Excel or use built-in G7 validation dataset.
Compute WENSLO: normalization, Δz, envelope E, q, weights.
Compute ALWAS: R1, R2, S and final ranking with sensitivity parameters.
Graduation/positioning.
Accumulation validation: MSE and correlation (real vs artificial accumulation).
Correlation heatmap between criteria (Pearson / Spearman) with PNG download.
Export all results to Excel.

Sensitivity parameters

Available in the sidebar:
ξ (xi) — integer (1–50)
φ (phi) — [0.0, 1.0]
θ (theta) — integer (1–50)

Changes require clicking Run to recompute results.
---
Deployment

Streamlit Cloud: simplest way to deploy. Link this GitHub repo to Streamlit Cloud and the app will run online.
Alternatives: Heroku, Railway, Docker.

---

Citation

If you use this tool in academic work, please cite the paper:
Method: A Novel WENSLO and ALWAS Multicriteria Methodology and Its Application to Green Growth Performance Evaluation (IEEE Transactions on Engineering Management, 2023).
Tool: XXXXX

---

Contact

Anderson Portella — https://www.linkedin.com/in/andersonportella/


## 3) Recommended auxiliars files
- `requirements.txt`.
- `setup.sh` (opcional) — Script for Linux/macOS Automate Venv creation and installation:
```bash
#!/usr/bin/env bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
echo "Prepared environment. Activate with: source .venv/bin/activate"

run.sh (opcional)

#!/usr/bin/env bash
source .venv/bin/activate
streamlit run app.py

.gitignore
.venv/
__pycache__/
*.pyc
*.pkl
*.xlsx
*.xls
.DS_Store


