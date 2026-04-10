"""
Quick debug script — tests token auth and project ID fetching independently of Streamlit.
Run with: python test_all_projects.py
"""

import requests
from app import get_token, get_all_project_ids, get_endpoints

# ── Fill these in ──────────────────────────────────────────────────────────────
REGION   = "North America"   # or "North America" / "Europe/UK"
EMAIL    = "guy.accettura@hammertechglobal.com"
PASSWORD = "G@etano1"
TENANT   = "usademo"
# ──────────────────────────────────────────────────────────────────────────────

auth_endpoint, api_base = get_endpoints(REGION)
projects_endpoint = f"{api_base}/projects"

print(f"Auth endpoint : {auth_endpoint}")
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

# Step 2 — fetch first page raw so we can see the actual response shape
print("\nFetching first page of projects (raw)...")
headers = {"Authorization": "Bearer " + token, "accept": "application/json"}
resp = requests.get(projects_endpoint, headers=headers, params={"skip": 0, "take": 5}, timeout=60)
print(f"  HTTP {resp.status_code}")
print(f"  Response (truncated): {resp.text[:500]}")

# Step 3 — run the full paginated fetch
print("\nFetching ALL project IDs via get_all_project_ids()...")
try:
    ids = get_all_project_ids(token, projects_endpoint)
    print(f"  Retrieved {len(ids)} project IDs.")
    if ids:
        print(f"  First 5 IDs: {ids[:5]}")
    else:
        print("  WARNING: list is empty — check the raw response shape above.")
except Exception as e:
    print(f"  FAILED: {e}")
