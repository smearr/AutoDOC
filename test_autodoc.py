"""
AutoDoc - Unit Tests
"""
import os
import csv
import pytest
import openpyxl
import tempfile
from autodoc_engine import parse_excel, generate_pdf_report, log_report, LOG_FILE


# ── Fixtures ──────────────────────────────────────────────────────────────────
@pytest.fixture
def sample_excel(tmp_path):
    """Creates a temporary Excel file with component data."""
    path = tmp_path / "test_components.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Component ID", "Name", "Type", "Voltage Rating (V)",
               "Current Rating (A)", "Material", "Status", "Engineer", "Notes"])
    ws.append(["C-001", "Main Breaker", "Circuit Breaker", 480, 100,
               "Steel", "Approved", "J. Smith", "UL Listed"])
    ws.append(["C-002", "Bus Bar L1", "Bus Bar", 480, 200,
               "Copper", "Under Review", "A. Patel", "Check torque"])
    ws.append(["C-003", "Control Relay", "Relay", 24, 5,
               "Plastic", "Approved", "J. Smith", "DIN rail mount"])
    wb.save(path)
    return str(path)


@pytest.fixture
def sample_components():
    return [
        {"Component ID": "C-001", "Name": "Main Breaker", "Type": "Circuit Breaker",
         "Voltage Rating (V)": 480, "Current Rating (A)": 100, "Material": "Steel",
         "Status": "Approved", "Engineer": "J. Smith", "Notes": "UL Listed"},
        {"Component ID": "C-002", "Name": "Bus Bar L1", "Type": "Bus Bar",
         "Voltage Rating (V)": 480, "Current Rating (A)": 200, "Material": "Copper",
         "Status": "Under Review", "Engineer": "A. Patel", "Notes": "Check torque"},
    ]


# ── Tests: Excel Parser ───────────────────────────────────────────────────────
class TestExcelParser:
    def test_parse_returns_list(self, sample_excel):
        result = parse_excel(sample_excel)
        assert isinstance(result, list)

    def test_parse_correct_row_count(self, sample_excel):
        result = parse_excel(sample_excel)
        assert len(result) == 3

    def test_parse_correct_keys(self, sample_excel):
        result = parse_excel(sample_excel)
        assert "Component ID" in result[0]
        assert "Name" in result[0]
        assert "Status" in result[0]

    def test_parse_correct_values(self, sample_excel):
        result = parse_excel(sample_excel)
        assert result[0]["Name"] == "Main Breaker"
        assert result[1]["Status"] == "Under Review"

    def test_parse_handles_none_cells(self, tmp_path):
        path = tmp_path / "sparse.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Component ID", "Name", "Notes"])
        ws.append(["C-001", "Relay", None])
        wb.save(path)
        result = parse_excel(str(path))
        assert result[0]["Notes"] == "N/A"


# ── Tests: PDF Generator ──────────────────────────────────────────────────────
class TestPDFGenerator:
    def test_generates_file(self, sample_components, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        os.makedirs("generated_reports", exist_ok=True)
        path = generate_pdf_report(sample_components, "Test Project", "J. Smith")
        assert os.path.isfile(path)

    def test_output_in_correct_directory(self, sample_components, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        os.makedirs("generated_reports", exist_ok=True)
        path = generate_pdf_report(sample_components, "PDU Project")
        assert "generated_reports" in path

    def test_pdf_not_empty(self, sample_components, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        os.makedirs("generated_reports", exist_ok=True)
        path = generate_pdf_report(sample_components, "PDU Project")
        assert os.path.getsize(path) > 1024  # at least 1KB

    def test_filename_contains_project(self, sample_components, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        os.makedirs("generated_reports", exist_ok=True)
        path = generate_pdf_report(sample_components, "My_Project")
        assert "My_Project" in path


# ── Tests: Logger ─────────────────────────────────────────────────────────────
class TestLogger:
    def test_creates_log_file(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        log_report("RPT-001", "Test", 5, "generated_reports/RPT-001.pdf")
        assert os.path.isfile(LOG_FILE)

    def test_log_contains_headers(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        log_report("RPT-001", "Test", 5, "generated_reports/RPT-001.pdf")
        with open(LOG_FILE, newline="") as f:
            reader = csv.DictReader(f)
            assert "report_id" in reader.fieldnames

    def test_log_row_values(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        log_report("RPT-XYZ", "MyProject", 12, "generated_reports/RPT-XYZ.pdf", "Success")
        with open(LOG_FILE, newline="") as f:
            rows = list(csv.DictReader(f))
        assert rows[0]["report_id"] == "RPT-XYZ"
        assert rows[0]["project"]   == "MyProject"
        assert rows[0]["component_count"] == "12"

    def test_log_appends_multiple_rows(self, tmp_path, monkeypatch):
        monkeypatch.chdir(tmp_path)
        log_report("RPT-001", "Proj A", 3, "path/a.pdf")
        log_report("RPT-002", "Proj B", 7, "path/b.pdf")
        with open(LOG_FILE, newline="") as f:
            rows = list(csv.DictReader(f))
        assert len(rows) == 2
