import streamlit as st
import requests
import openpyxl
from datetime import datetime
import io

AUTH_ENDPOINT = "https://us-auth.hammertechonline.com/api/login/generatetoken"
API_BASE = "https://us-api.hammertechonline.com/api/v1"
USER_ENDPOINT = f"{API_BASE}/workerprofiles"
PROJECT_ENDPOINT = f"{API_BASE}/projects"

def get_token(email, password, tenant):
    headers = {"accept": "application/json", "Content-Type": "application/json"}
    body = {"email": email, "password": password, "tenant": tenant}
    r = requests.post(AUTH_ENDPOINT, headers=headers, json=body)
    r.raise_for_status()
    data = r.json()
    if "token" not in data:
        raise ValueError("No token found in response. Check credentials / tenant.")
    return data["token"]


def post_to_api(token, payload, upload_type: str):
    """Send either worker or project payload to the correct endpoint."""
    if upload_type == "Users":
        endpoint = USER_ENDPOINT
    else:
        endpoint = PROJECT_ENDPOINT

    headers = {"Authorization": "Bearer " + token}
    r = requests.post(endpoint, headers=headers, json=payload)
    return r.status_code, r.text


st.set_page_config(page_title="HammerTech Uploader", layout="wide")
st.title("HammerTech Uploader")

st.markdown(
    """
Upload an Excel file, enter your HammerTech credentials and tenant, and this app
will create **users** or **projects** via the HammerTech API.
"""
)

# --- Credentials ---
st.header("Upload Type")
upload_type = st.selectbox("What do you want to upload?", ["Users", "Projects"])

if upload_type == "Users":
    st.info(
        "You selected **Users**. Expected columns (in order): "
        "`Email`, `Full Name`, `Phone`, `Job Title`, `Internal Identifier`, `Demo Project ID`."
    )
else:
    st.info(
        "You selected **Projects**. Expected columns (in order): "
        "`ProjectName`, `siteAddress`, `timeZoneString`, `state`, `internalid`, `regionId`."
    )

# ----------------- CREDENTIALS -----------------
st.header("API Credentials")   
col1, col2 = st.columns(2)

with col1:
    email = st.text_input("Email")
    tenant = st.text_input("Tenant")

with col2:
    password = st.text_input("Password", type="password")

# --- File + sheet settings ---
st.header("Excel File")

with st.expander("User Upload Template"):
    st.write(
        "Download this Excel template, fill in the user details, "
        "and then upload it below."
    )
    try:
        with open("userUploadTemplate.xlsx", "rb") as f:
            st.download_button(
                label="Download user template",
                data=f,
                file_name="userUploadTemplate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except FileNotFoundError:
        st.error("Template file not found on the server. Please contact the administrator.")


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

        if upload_type = "Users"
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
        
        # ------------- PROJECTS UPLOAD FLOW -------------
        else:  # upload_type == "Projects"
            for i, row in enumerate(rows, start=1):
                if i < start_row:
                    continue

                ProjectName = row[0]
                siteAddress = row[1]
                timeZoneString = row[2]
                state = row[3]
                internalid = row[4]
                regionId = row[5]

                # skip completely empty rows
                if not any([ProjectName, siteAddress, timeZoneString, state, internalid, regionId]):
                    continue

                payload = {
                    "isArchived": False,
                    "name": ProjectName,
                    "siteAddress": siteAddress,
                    "regionId": regionId,
                    "state": state,
                    "timeZoneString": timeZoneString,
                    "country": "Canada",
                    "internalIdentifier": internalid,
                    "siteTiming": [
                        {
                            "dayOfWeek": "Wednesday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        },
                        {
                            "dayOfWeek": "Thursday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        },
                        {
                            "dayOfWeek": "Friday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        },
                        {
                            "dayOfWeek": "Saturday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        },
                        {
                            "dayOfWeek": "Sunday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        },
                        {
                            "dayOfWeek": "Monday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        },
                        {
                            "dayOfWeek": "Tuesday",
                            "startTime": "07:00:00",
                            "endTime": "17:00:00"
                        }
                    ]
                }

                row_count += 1
                logs.append(f"Row {i}: Sending project {ProjectName}...")

                try:
                    status_code, response_text = post_to_api(token, payload, upload_type)
                    if 200 <= status_code < 300:
                        logs.append(f"Row {i}: ✅ Success (HTTP {status_code}).")
                        success_count += 1
                    else:
                        logs.append(
                            f"Row {i}: ❌ Failed (HTTP {status_code}). Response: {response_text}"
                        )
                        fail_count += 1
                except Exception as e:
                    logs.append(f"Row {i}: ❌ Error sending project: {e}")
                    fail_count += 1

                progress_bar.progress(min(i / total_rows, 1.0))
                log_area.text("\n".join(logs[-20:]))

        # ------------- SUMMARY -------------
        st.success("Upload complete.")
        st.write(f"**Total rows processed:** {row_count}")
        st.write(f"✅ Successful: {success_count}")
        st.write(f"❌ Failed: {fail_count}")

        with st.expander("View full log"):
            st.text("\n".join(logs))
