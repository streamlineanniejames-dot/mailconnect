import streamlit as st
import pandas as pd
import base64
import time
import re
import json
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="📧 Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge System")

# ========================================
# Google OAuth Configuration
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
]

def gmail_authenticate():
    """Authenticate user with Google OAuth"""
    if "credentials" not in st.session_state:
        st.session_state.credentials = None

    if st.session_state.credentials:
        creds = Credentials.from_authorized_user_info(st.session_state.credentials, SCOPES)
        if creds and creds.valid:
            return creds
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            return creds

    # OAuth Flow
    flow = Flow.from_client_secrets_file(
        "credentials.json",
        scopes=SCOPES,
        redirect_uri="http://localhost:8501"
    )
    auth_url, _ = flow.authorization_url(prompt="consent")

    st.markdown(f"[Authorize Gmail Access]({auth_url})")

    auth_code = st.text_input("Paste the authorization code here:")
    if auth_code:
        flow.fetch_token(code=auth_code)
        creds = flow.credentials
        st.session_state.credentials = json.loads(creds.to_json())
        st.success("✅ Authentication successful!")
        return creds
    return None

# ========================================
# Label Handling — Stable Version
# ========================================
def get_or_create_label(service, label_name="Mail Merge Sent"):
    """Get or create a Gmail label reliably, even if called multiple times."""
    try:
        response = service.users().labels().list(userId="me").execute()
        labels = response.get("labels", [])

        # Try exact match (case-insensitive)
        existing_label = next((l for l in labels if l["name"].lower() == label_name.lower()), None)
        if existing_label:
            return existing_label["id"]

        # Create new label if not found
        new_label = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        created = service.users().labels().create(userId="me", body=new_label).execute()

        # Wait to ensure Gmail syncs new label
        time.sleep(2)

        # Confirm label creation
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        confirmed = next((l for l in labels if l["name"].lower() == label_name.lower()), None)

        return confirmed["id"] if confirmed else created["id"]

    except Exception as e:
        st.error(f"⚠️ Label error: {e}")
        return None

# ========================================
# Gmail Send Function
# ========================================
def create_message(sender, to, subject, message_text):
    msg = MIMEText(message_text, "html")
    msg["to"] = to
    msg["from"] = sender
    msg["subject"] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}

def send_email(service, sender, to, subject, body, label_id):
    msg_body = create_message(sender, to, subject, body)
    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
    msg_id = sent_msg["id"]

    # Apply label after send
    if label_id:
        try:
            modified = service.users().messages().modify(
                userId="me",
                id=msg_id,
                body={"addLabelIds": [label_id]},
            ).execute()
            st.write(f"✅ Label applied: {modified.get('labelIds', [])}")
        except Exception as e:
            st.error(f"❌ Label apply failed: {e}")

# ========================================
# UI Section
# ========================================
st.header("📤 Upload Recipient List")
uploaded_file = st.file_uploader("Upload CSV file with columns: Email, Name", type=["csv"])
subject = st.text_input("Email Subject:")
body = st.text_area("Email Body (you can use {Name} placeholders):")
label_name = st.text_input("Label Name:", value="Mail Merge Sent")

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df.head())

creds = gmail_authenticate()

# ========================================
# Sending Logic
# ========================================
if creds and uploaded_file and st.button("🚀 Send Emails"):
    try:
        service = build("gmail", "v1", credentials=creds)  # fresh client
        sender_info = service.users().getProfile(userId="me").execute()
        sender_email = sender_info["emailAddress"]
        label_id = get_or_create_label(service, label_name)

        for index, row in df.iterrows():
            to = row["Email"]
            msg_body = body.format(**row)
            send_email(service, sender_email, to, subject, msg_body, label_id)
            time.sleep(1)

        st.success("✅ All emails sent successfully with label applied!")

    except Exception as e:
        st.error(f"❌ Error sending emails: {e}")
