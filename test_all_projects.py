"""
Quick debug script — tests token auth and project ID fetching independently of Streamlit.
Run with: python test_all_projects.py
"""

import requests
import openpyxl
import io
from app import get_token, get_all_project_ids, get_endpoints, build_user_payload

# ── Fill these in ──────────────────────────────────────────────────────────────
REGION   = "North America"
EMAIL    = "guy.accettura@hammertechglobal.com"
PASSWORD = "G@etano1"
TENANT   = "usademo"
EXCEL_PATH = "UploadTemplateTest.xlsx"  # path to your test file
# ──────────────────────────────────────────────────────────────────────────────

auth_endpoint, api_base = get_endpoints(REGION)
projects_endpoint = f"{api_base}/projects"

print(f"Auth endpoint    : {auth_endpoint}")
print(f"Projects endpoint: {projects_endpoint}")
print()

# Step 1 — get token
print("Fetching token...")
try:
    token = get_token(auth_endpoint, EMAIL, PASSWORD, TENANT)
    print(f"  Token OK (first 20 chars): {token[:20]}...")
except Exception as e:
    print(f"  FAILED: {e}")
    raise SystemExit(1)

# Step 2 — fetch all project IDs
print("\nFetching ALL project IDs...")
try:
    all_ids = get_all_project_ids(token, projects_endpoint)
    print(f"  Retrieved {len(all_ids)} project IDs.")
except Exception as e:
    print(f"  FAILED: {e}")
    raise SystemExit(1)

# Step 3 — inspect the Excel cell values and built payload
print(f"\nReading Excel file: {EXCEL_PATH}")
try:
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb["Users"]
    rows = list(sheet.iter_rows(values_only=True))
    for i, row in enumerate(rows[1:], start=2):  # skip header
        if not any(row):
            continue
        raw_project_val = row[5] if len(row) > 5 else None
        print(f"\n  Row {i}:")
        print(f"    Col F raw value : {raw_project_val!r}  (type: {type(raw_project_val).__name__})")
        is_all = str(raw_project_val).strip().lower() == "all" if raw_project_val is not None else False
        print(f"    is_all          : {is_all}")
        _, payload = build_user_payload(row, all_project_ids=all_ids)
        print(f"    isAddToFutureProjects : {payload.get('isAddToFutureProjects')}")
        print(f"    userProjectIds count  : {len(payload.get('userProjectIds', []))}")
except FileNotFoundError:
    print(f"  File not found: {EXCEL_PATH} — update EXCEL_PATH in this script.")
except Exception as e:
    print(f"  FAILED: {e}")
