import streamlit as st
import requests
import openpyxl
from datetime import datetime
import io

AUTH_ENDPOINT = "https://us-auth.hammertechonline.com/api/login/generatetoken"
WORKER_ENDPOINT = "https://us-api.hammertechonline.com/api/v1/users"


def get_token(email, password, tenant):
    headers = {"accept": "application/json", "Content-Type": "application/json"}
    body = {"email": email, "password": password, "tenant": tenant}
    r = requests.post(AUTH_ENDPOINT, headers=headers, json=body)
    r.raise_for_status()
    data = r.json()
    if "token" not in data:
        raise ValueError("No token found in response. Check credentials / tenant.")
    return data["token"]


def post_worker(token, payload):
    headers = {"Authorization": "Bearer " + token}
    r = requests.post(WORKER_ENDPOINT, headers=headers, json=payload)
    return r.status_code, r.text


st.set_page_config(page_title="HammerTech Standard User Uploader", layout="wide")
st.title("HammerTech Standard User Uploader")

st.markdown(
    """
Upload an Excel file, enter your HammerTech credentials and tenant, and this app
will create user profiles via the HammerTech API.
"""
)

# --- Credentials ---
st.header("API Credentials")
col1, col2 = st.columns(2)

with col1:
    email = st.text_input("Email")
    tenant = st.text_input("Tenant")

with col2:
    password = st.text_input("Password", type="password")

# --- File + sheet settings ---
st.header("Excel File")

uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx", "xlsm"])
sheet_name = st.text_input("Sheet name")
start_row = st.number_input(
    "Start row (1-based, usually 2 to skip header)",
    min_value=1,
    value=2
)

run_button = st.button("Run Upload")

if run_button:
    if not (email and password and tenant and uploaded_file and sheet_name):
        st.error("Please fill in all fields and upload a file.")
    else:
        # Authenticate
        try:
            with st.spinner("Authenticating with HammerTech..."):
                token = get_token(email, password, tenant)
            st.success("Authentication successful.")
        except Exception as e:
            st.error(f"Authentication failed: {e}")
            st.stop()

        # Load workbook
        try:
            file_bytes = uploaded_file.read()
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        except Exception as e:
            st.error(f"Failed to load workbook: {e}")
            st.stop()

        if sheet_name not in wb.sheetnames:
            st.error(f"Sheet '{sheet_name}' not found in workbook. Available sheets: {wb.sheetnames}")
            st.stop()

        sheet = wb[sheet_name]
        st.write(f"Using sheet: `{sheet_name}`")

        rows = list(sheet.iter_rows(values_only=True))
        total_rows = len(rows)

        row_count = 0
        success_count = 0
        fail_count = 0
        logs = []

        progress_bar = st.progress(0)
        log_area = st.empty()

        for i, row in enumerate(rows, start=1):
            if i < start_row:
                continue

            email = row[0]
            name = row[1]
            mobile = row[2]
            title = row[3]
            role_name = ["safetymanager"]
            internalIdentifier = row[4]
            user_project_ids = row[5]

            if email is None:
                # Skip empty row
                continue

            mobile_str = str(mobile) if mobile is not None else ""
            internalIdentifier_str = str(internalIdentifier) if internalIdentifier is not None else ""

            payload = {
                "name": name or "",
                "title": title or "",
                "mobile": mobile_str,
                "email": email or "",
                "internalIdentifier": internalIdentifier_str,
                "roleNames": role_name or "",
                "userProjectIds": [user_project_ids] or ""
            }

            row_count += 1
            logs.append(f"Row {i}: Sending worker {name}...")

            try:
                status_code, response_text = post_worker(token, payload)
                if 200 <= status_code < 300:
                    logs.append(f"Row {i}: ✅ Success (HTTP {status_code}).")
                    success_count += 1
                else:
                    logs.append(
                        f"Row {i}: ❌ Failed (HTTP {status_code}). Response: {response_text}"
                    )
                    fail_count += 1
            except Exception as e:
                logs.append(f"Row {i}: ❌ Error sending worker: {e}")
                fail_count += 1

            # Update UI
            progress_bar.progress(min(i / total_rows, 1.0))
            log_area.text("\n".join(logs[-20:]))

        st.success("Upload complete.")
        st.write(f"**Total rows processed:** {row_count}")
        st.write(f"✅ Successful: {success_count}")
        st.write(f"❌ Failed: {fail_count}")

        with st.expander("View full log"):
            st.text("\n".join(logs))
