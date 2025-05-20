# VIKOR Method Web App

This is a web-based tool for performing Multi-Criteria Decision-Making using the **VIKOR method**. It is built with Streamlit and supports Excel-based input and downloadable output reports.

## 🚀 Features

- Upload Excel files with alternatives and criteria
- Normalize based on benefit/cost types
- Apply weights to criteria
- Compute S, R, and Q scores
- Visualize Q values in a bar chart
- Export all steps to an Excel report

## 📁 Excel Input Format

### Sheet 1: `Data`

| Alternative | C1  | C2  | C3  |
|-------------|-----|-----|-----|
| A1          | 70  | 20  | 80  |
| A2          | 60  | 30  | 90  |
| A3          | 80  | 25  | 85  |

### Sheet 2: `CriteriaInfo`

| Criterion | Type    | Weight |
|-----------|---------|--------|
| C1        | benefit | 0.3    |
| C2        | cost    | 0.4    |
| C3        | benefit | 0.3    |

### Sheet 3: `TargetValues` (optional for compatibility)

| Criterion | Target |
|-----------|--------|
| C1        |        |
| C2        |        |
| C3        |        |

## 🛠 How to Run Locally

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Run the app:

```bash
streamlit run app_vikor.py
```

3. Open your browser to `http://localhost:8501/`

## 📦 Files Included

- `app_vikor.py` — Main Streamlit app
- `example_vikor_input.xlsx` — Sample Excel file
- `requirements.txt` — Python dependencies
- `README.md` — Documentation

## 📤 Output

The app allows you to download the results as an Excel report with the following sheets:
- 1_RawData
- 2_Normalized
- 3_Weighted
- 4_FinalScores

---

© 2025 Dr. Zahari Md Rodzi | UiTM Negeri Sembilan
