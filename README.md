# 📊 Excel Automator

A free, browser-based Excel automation toolkit built with Python + Streamlit.
No Excel license required. No data ever leaves your browser session.

## Tools included

| Tool | What it does |
|------|-------------|
| **Sheet / File Merger** | Combine multiple `.xlsx` files or sheets into one master file |
| **Auto-Formatter** | Highlight overdue dates 🔴, budget overruns 🟠, and negative values 🟡 |
| **Budget Variance Reporter** | Compare actuals vs. forecast, flag variances > threshold, generate a summary + bar chart |

---

## Run locally

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/excel-automator.git
cd excel-automator
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Launch the app
```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## Deploy for free on Streamlit Community Cloud

> Live in ~2 minutes. No credit card. No server to manage.

1. **Push this repo to GitHub** (must be public for the free tier)
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR_USERNAME/excel-automator.git
   git push -u origin main
   ```

2. **Go to [share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub.

3. Click **"New app"** → select your repo → set **Main file path** to `app.py` → click **Deploy**.

4. Your app gets a public URL like `https://your-username-excel-automator-app-xxxx.streamlit.app`

---

## File format guide

### Merger
- Upload any `.xlsx` files with consistent or inconsistent columns — a `_source_file` column is added automatically.

### Auto-Formatter
- Any `.xlsx` file works. Columns are detected by keyword:
  - **Date columns**: headers containing `date`, `due`, `deadline`, `expiry`
  - **Budget columns**: headers containing `budget`, `cost`, `spend`, `amount`

### Variance Reporter
Your file must have at minimum:

| Column type | Accepted names |
|-------------|---------------|
| Category | `category`, `department`, `item`, `account` |
| Budget | `budget`, `forecast`, `plan` |
| Actual | `actual`, `actuals`, `spend` |

---

## Project structure

```
excel-automator/
├── app.py                  # Streamlit UI
├── requirements.txt
├── tools/
│   ├── merger.py           # File/sheet merge logic
│   ├── formatter.py        # Highlighting & formatting
│   └── variance.py         # Variance report + chart
└── sample_data/
    ├── sample_merge_a.xlsx
    ├── sample_merge_b.xlsx
    ├── sample_format.xlsx
    └── sample_variance.xlsx
```

---

## Tech stack

- [Streamlit](https://streamlit.io) — UI framework
- [pandas](https://pandas.pydata.org) — data manipulation
- [openpyxl](https://openpyxl.readthedocs.io) — Excel read/write + styling
- [xlsxwriter](https://xlsxwriter.readthedocs.io) — Excel write with charts

---

## License

MIT — free to use, fork, and extend.
