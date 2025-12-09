import streamlit as st
import requests
import openpyxl
from datetime import datetime
import io

# --------- REGION-BASED ENDPOINTS ---------
def get_endpoints(region: str):
    """Return auth endpoint and API base (with /api/v1) for the selected region."""
    if region == "North America":
        auth = "https://us-auth.hammertechonline.com/api/login/generatetoken"
        api_base = "https://us-api.hammertechonline.com/api/v1"
    elif region == "Asia/Australia/NZ":
        auth = "https://au-auth.hammertechonline.com/api/login/generatetoken"
        api_base = "https://au-api.hammertechonline.com/api/v1"
    else:  # Europe/UK
        auth = "https://eu-auth.hammertechonline.com/api/login/generatetoken"
        api_base = "https://eu-api.hammertechonline.com/api/v1"

    return auth, api_base


def get_token(auth_endpoint, email, password, tenant):
    headers = {"accept": "application/json", "Content-Type": "application/json"}
    body = {"email": email, "password": password, "tenant": tenant}
    r = requests.post(auth_endpoint, headers=headers, json=body)
    r.raise_for_status()
    data = r.json()
    if "token" not in data:
        raise ValueError("No token found in response. Check credentials / tenant.")
    return data["token"]


def post_to_api(
    token,
    payload,
    upload_type: str,
    user_endpoint: str,
    project_endpoint: str,
    employer_endpoint: str,
):
    """
    Send payload to the correct endpoint based on upload_type:
    - Users            -> user_endpoint (Worker Profiles)
    - Projects         -> project_endpoint (Projects)
    - Employer Profiles -> employer_endpoint (EmployerProfiles)
    """
    if upload_type == "Users":
        endpoint = user_endpoint
    elif upload_type == "Projects":
        endpoint = project_endpoint
    else:  # Employer Profiles
        endpoint = employer_endpoint

    headers = {"Authorization": "Bearer " + token}
    r = requests.post(endpoint, headers=headers, json=payload)
    return r.status_code, r.text


st.set_page_config(page_title="HammerTech Uploader", layout="wide")
st.title("HammerTech Uploader")

st.markdown(
    """
Upload an Excel file, enter your HammerTech credentials and tenant, and this app
will create **users**, **projects**, or **employer profiles** via the HammerTech API.
"""
)

# ----------------- REGION SELECTION -----------------
st.header("Region")
region = st.selectbox(
    "Select your HammerTech region",
    ["North America", "Asia/Australia/NZ", "Europe/UK"],
)
st.caption(
    "This controls which auth and API endpoints are used (US, AU, or EU environments)."
)

auth_endpoint, api_base = get_endpoints(region)
USER_ENDPOINT = f"{api_base}/workerprofiles"
PROJECT_ENDPOINT = f"{api_base}/projects"
EMPLOYER_ENDPOINT = f"{api_base}/EmployerProfiles"

# ----------------- UPLOAD TYPE -----------------
st.header("Upload Type")
upload_type = st.selectbox(
    "What do you want to upload?",
    ["Users", "Projects", "Employer Profiles"],
)

if upload_type == "Users":
    st.info(
        "You selected **Users**. Expected columns (in order):\n\n"
        "1. Email\n"
        "2. Full Name\n"
        "3. Phone\n"
        "4. Job Title\n"
        "5. Internal Identifier\n"
        "6. Demo Project ID"
    )
elif upload_type == "Projects":
    st.info(
        "You selected **Projects**. Expected columns (in order):\n\n"
        "1. ProjectName\n"
        "2. Country\n"
        "3. siteAddress\n"
        "4. timeZoneString\n"
        "5. state\n"
        "6. internalid\n"
        "7. regionId"
    )
else:  # Employer Profiles
    st.info(
        "You selected **Employer Profiles**. Expected columns (in order):\n\n"
        "1. Business Name\n"
        "2. ABN\n"
        "3. Street Address\n"
        "4. City / Suburb\n"
        "5. State / Province\n"
        "6. Postal Code\n"
        "7. Country\n"
        "8. Internal Identifier"
    )

# ----------------- CREDENTIALS -----------------
st.header("API Credentials")
col1, col2 = st.columns(2)

with col1:
    email = st.text_input("Email")
    tenant = st.text_input("Tenant")

with col2:
    password = st.text_input("Password", type="password")

# ----------------- FILE + SHEET -----------------
st.header("Excel File")

with st.expander("Upload Templates"):
    st.write(
        "Download the needed template, fill it in, "
        "and then upload it below."
    )
    # User template
    try:
        with open("userUploadTemplate.xlsx", "rb") as f:
            st.download_button(
                label="Download User Upload Template",
                data=f,
                file_name="userUploadTemplate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except FileNotFoundError:
        st.error("User template file not found on the server. Please contact the administrator.")

    # Project template
    try:
        with open("projectUploadTemplate.xlsx", "rb") as f:
            st.download_button(
                label="Download Project Upload Template",
                data=f,
                file_name="projectUploadTemplate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except FileNotFoundError:
        st.error("Project template file not found on the server. Please contact the administrator.")

    # Employer Profile template (optional file if you create one)
    try:
        with open("employerProfileUploadTemplate.xlsx", "rb") as f:
            st.download_button(
                label="Download Employer Profile Upload Template",
                data=f,
                file_name="employerProfileUploadTemplate.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except FileNotFoundError:
        # Only show as error if they’ve selected Employer Profiles
        if upload_type == "Employer Profiles":
            st.error("Employer Profile template file not found on the server. Please contact the administrator.")

uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx", "xlsm"])
sheet_name = st.text_input("Sheet name")
start_row = st.number_input(
    "Start row (1-based, usually 2 to skip header)",
    min_value=1,
    value=2
)

run_button = st.button("Run Upload")

# ----------------- MAIN LOGIC -----------------
if run_button:
    if not (email and password and tenant and uploaded_file and sheet_name):
        st.error("Please fill in all fields and upload a file.")
    else:
        # Authenticate
        try:
            with st.spinner(f"Authenticating with HammerTech ({region})..."):
                token = get_token(auth_endpoint, email, password, tenant)
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

        # ------------- USERS UPLOAD FLOW -------------
        if upload_type == "Users":
            for i, row in enumerate(rows, start=1):
                if i < start_row:
                    continue

                email_cell = row[0]
                name = row[1]
                mobile = row[2]
                title = row[3]
                internalIdentifier = row[4]
                user_project_ids = row[5]

                if email_cell is None:
                    # Skip empty row
                    continue

                mobile_str = str(mobile) if mobile is not None else ""
                internalIdentifier_str = (
                    str(internalIdentifier) if internalIdentifier is not None else ""
                )

                # roleNames as a list, project IDs as a list (if present)
                role_names = ["safetymanager"]
                user_project_ids_list = (
                    [user_project_ids] if user_project_ids is not None else []
                )

                payload = {
                    "name": name or "",
                    "title": title or "",
                    "mobile": mobile_str,
                    "email": email_cell or "",
                    "internalIdentifier": internalIdentifier_str,
                    "roleNames": role_names,
                    "userProjectIds": user_project_ids_list,
                }

                row_count += 1
                logs.append(f"Row {i}: Sending user {name}...")

                try:
                    status_code, response_text = post_to_api(
                        token,
                        payload,
                        upload_type,
                        USER_ENDPOINT,
                        PROJECT_ENDPOINT,
                        EMPLOYER_ENDPOINT,
                    )
                    if 200 <= status_code < 300:
                        logs.append(f"Row {i}: ✅ Success (HTTP {status_code}).")
                        success_count += 1
                    else:
                        logs.append(
                            f"Row {i}: ❌ Failed (HTTP {status_code}). Response: {response_text}"
                        )
                        fail_count += 1
                except Exception as e:
                    logs.append(f"Row {i}: ❌ Error sending user: {e}")
                    fail_count += 1

                # Update UI
                progress_bar.progress(min(i / total_rows, 1.0))
                log_area.text("\n".join(logs[-20:]))

        # ------------- PROJECTS UPLOAD FLOW -------------
        elif upload_type == "Projects":
            for i, row in enumerate(rows, start=1):
                if i < start_row:
                    continue

                ProjectName = row[0]
                country = row[1]
                siteAddress = row[2]
                timeZoneString = row[3]
                state = row[4]
                internalid = row[5]
                regionId = row[6]

                # skip completely empty rows
                if not any([ProjectName, country, siteAddress, timeZoneString, state, internalid, regionId]):
                    continue

                payload = {
                    "isArchived": False,
                    "name": ProjectName,
                    "siteAddress": siteAddress,
                    "regionId": regionId,
                    "state": state,
                    "timeZoneString": timeZoneString,
                    "country": country,
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
                    status_code, response_text = post_to_api(
                        token,
                        payload,
                        upload_type,
                        USER_ENDPOINT,
                        PROJECT_ENDPOINT,
                        EMPLOYER_ENDPOINT,
                    )
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

        # ------------- EMPLOYER PROFILES UPLOAD FLOW -------------
        else:  # upload_type == "Employer Profiles"
            for i, row in enumerate(rows, start=1):
                if i < start_row:
                    continue

                business_name = row[0]
                abn = row[1]
                streetAddress = row[2]
                city = row[3]
                state = row[4]
                postal_code = row[5]
                country = row[6]
                internalIdentifier = row[7] if len(row) > 7 else None

                # skip completely empty rows
                if not any(
                    [
                        business_name,
                        abn,
                        streetAddress,
                        city,
                        state,
                        postal_code,
                        country,
                        internalIdentifier,
                    ]
                ):
                    continue

                internalIdentifier_str = (
                    str(internalIdentifier) if internalIdentifier is not None else ""
                )
                abn_str = str(abn) if abn is not None else ""
                post_code_str = str(postal_code) if postal_code is not None else ""

                payload = {
                    "businessName": business_name or "",
                    "abn": abn_str,
                    "addresses": [
                        {
                        "addressType": "Physical",
                        "streetAddress": streetAddress or "",
                        "suburb": city or "",
                        "state": state or "",
                        "postCode": post_code_str or "",
                        "country": country or "",
                        }
                    ],
                    "internalIdentifier": internalIdentifier_str,
                }

                row_count += 1
                logs.append(f"Row {i}: Sending employer profile {business_name}...")

                try:
                    status_code, response_text = post_to_api(
                        token,
                        payload,
                        upload_type,
                        USER_ENDPOINT,
                        PROJECT_ENDPOINT,
                        EMPLOYER_ENDPOINT,
                    )
                    if 200 <= status_code < 300:
                        logs.append(f"Row {i}: ✅ Success (HTTP {status_code}).")
                        success_count += 1
                    else:
                        logs.append(
                            f"Row {i}: ❌ Failed (HTTP {status_code}). Response: {response_text}"
                        )
                        fail_count += 1
                except Exception as e:
                    logs.append(f"Row {i}: ❌ Error sending employer profile: {e}")
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
