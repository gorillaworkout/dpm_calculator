"""Regression check for the downloadable blank KVB workbook template."""
import base64
import io
import os
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(tempfile.mkdtemp())
os.environ["DW_AUTH_DB"] = str(ROOT / "users.db")
os.environ["DW_JOB_DIR"] = str(ROOT / "jobs")
os.environ["DW_ADMIN_USER"] = "tester"
os.environ["DW_ADMIN_PASS"] = "tester-password-1"

from app import app

client = app.test_client()
client.environ_base["HTTP_AUTHORIZATION"] = "Basic " + base64.b64encode(
    b"tester:tester-password-1"
).decode()

page = client.get("/kvb/tool/dw")
assert page.status_code == 200
assert b'href="/kvb/template"' in page.data
assert b"Download Template KVB Plus.xlsx" in page.data

response = client.get("/kvb/template")
assert response.status_code == 200
assert "Template KVB Plus.xlsx" in response.headers["Content-Disposition"]

workbook = load_workbook(io.BytesIO(response.data), read_only=True, data_only=True)
assert workbook.sheetnames == ["Transfer", "D", "W", "Xero", "Handling Fee"]
assert workbook["Transfer"].max_row == 2
assert workbook["Transfer"]["A1"].value == "汇出"
assert workbook["Transfer"]["A2"].value == "汇款日期"
for sheet in ("D", "W", "Xero", "Handling Fee"):
    assert workbook[sheet].max_row == 1, f"{sheet} must contain headers only"
assert workbook["D"]["A1"].value == "Source.Name"
assert workbook["W"]["A1"].value == "Source.Name"
assert workbook["Xero"]["A1"].value == "Date"
assert workbook["Handling Fee"]["A1"].value == "Payment Channel"
workbook.close()

print("KVB template route, sidebar link, sheets, headers, and blank rows: OK")
