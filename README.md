# AutoDoc — Engineering Report Automation Pipeline

> A Python-based automation pipeline that ingests structured Excel component specification data, auto-generates formatted PDF engineering reports, logs all output to a CSV data source, and visualizes metrics in a live Power BI dashboard.

---

## Features

- **Excel → PDF Pipeline** — Upload any `.xlsx` component spec sheet and get a formatted engineering report in seconds
- **Dynamic PDF Generation** — ReportLab-powered reports with branded headers, component tables, and summary stats
- **CSV Logging** — Every report run is logged with metadata (ID, project, component count, status, timestamp)
- **Power BI Integration** — The `report_log.csv` file serves as a live data source for a connected Power BI dashboard
- **Power Automate Trigger** — Webhook flow sends an email summary when a new report is generated
- **REST API** — Flask backend with clean endpoints for generation, download, and stats
- **Frontend Dashboard** — Vanilla JS + CSS dashboard UI for pipeline control and report analytics
- **CI/CD** — GitHub Actions workflow runs `flake8` lint and `pytest` unit tests on every push

---

## Tech Stack

| Layer | Technology |
|---|---|
| Core Engine | Python 3.11 |
| Excel Parsing | openpyxl |
| PDF Generation | ReportLab |
| Templating | Jinja2 |
| Web Backend | Flask + Flask-CORS |
| Frontend | HTML / CSS / JavaScript |
| Analytics | Power BI (CSV data source) |
| Notifications | Power Automate (HTTP trigger) |
| Version Control | Git + GitHub |
| CI/CD | GitHub Actions |

---

## Quick Start

```bash
# Clone the repo
git clone https://github.com/your-username/autodoc.git
cd autodoc

# Install dependencies
pip install -r requirements.txt

# Start the server
python app.py

# Open http://localhost:5000
```

---

## Usage

1. Open the web UI at `http://localhost:5000`
2. Click **↓ Sample Excel** to download a template
3. Fill in your component data and save
4. Enter your project name and engineer name
5. Upload the Excel file and click **⚡ Generate Report**
6. Download the PDF from the success message
7. View metrics on the **Dashboard** tab

---

## Command-Line Usage

```bash
python autodoc_engine.py components.xlsx "PDU-2024-Q1" "J. Smith"
```

---

## Power BI Setup

1. Open Power BI Desktop
2. **Get Data → Text/CSV** → select `report_log.csv`
3. Enable **Scheduled Refresh** to keep the dashboard live
4. Build visuals using fields: `report_id`, `project`, `component_count`, `status`, `generated_at`

---

## Running Tests

```bash
pytest tests/ -v
```

---

## Project Structure

```
autodoc/
├── autodoc_engine.py      # Core pipeline (parse → generate → log)
├── app.py                 # Flask REST API server
├── requirements.txt
├── README.md
├── report_log.csv         # Auto-generated; Power BI data source
├── frontend/
│   └── index.html         # Web UI (dashboard + generate)
├── generated_reports/     # Output PDFs (.gitignored)
├── uploads/               # Temp Excel uploads (.gitignored)
├── tests/
│   └── test_autodoc.py    # Unit tests (pytest)
└── .github/
    └── workflows/
        └── ci.yml         # GitHub Actions CI
```

---

## API Reference

| Method | Endpoint | Description |
|---|---|---|
| `POST` | `/api/generate` | Upload Excel, run pipeline, return report |
| `GET` | `/api/stats` | Aggregate metrics (feeds dashboard) |
| `GET` | `/api/logs` | Full report log as JSON |
| `GET` | `/api/download/:filename` | Download a generated PDF |
| `GET` | `/api/sample` | Download sample Excel template |

---

## License

MIT
