import streamlit as st
import pandas as pd
import base64
import time
import json
import random
import os
import re
from datetime import datetime, timedelta
import pytz
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

st.set_page_config(page_title="Send Emails", layout="wide")
st.title("🚀 Step 2: Send / Draft Emails")

# ===============================
# Load saved data from session
# ===============================
if "recipients_df" not in st.session_state:
    st.warning("⚠️ Please go to the 'Compose & Upload' page first to prepare your email and data.")
    st.stop()

df = st.session_state["recipients_df"]
subject_template = st.session_state["subject_template"]
body_template = st.session_state["body_template"]

# ===============================
# Gmail OAuth Setup
# ===============================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.compose",
]

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
else:
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code[0])
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize this app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

service = build("gmail", "v1", credentials=creds)

# ===============================
# Helper functions
# ===============================
def convert_bold(text):
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"\[(.*?)\]\((https?://[^\s)]+)\)", r'<a href="\2" style="color:#1a73e8;">\1</a>', text)
    return f"<html><body style='font-family: Verdana; font-size:14px;'>{text.replace(chr(10),'<br>')}</body></html>"

def get_or_create_label(service, label_name="Mail Merge Sent"):
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for l in labels:
        if l["name"].lower() == label_name.lower():
            return l["id"]
    new_label = service.users().labels().create(userId="me", body={"name": label_name}).execute()
    return new_label["id"]

def send_backup_csv(service, csv_path):
    try:
        user_email = service.users().getProfile(userId="me").execute().get("emailAddress")
        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        msg.attach(MIMEText("Attached is the backup CSV from your recent mail merge.", "plain"))

        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        st.info(f"📩 Backup CSV sent to {user_email}.")
    except Exception as e:
        st.warning(f"Backup email failed: {e}")

# ===============================
# Settings and Sending
# ===============================
label_name = st.text_input("🏷️ Gmail label name", value="Mail Merge Sent")
delay = st.slider("Delay between emails (seconds)", 20, 75, 25)
send_mode = st.radio("Choose Mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

if st.button("🚀 Start Sending / Save Drafts"):
    label_id = get_or_create_label(service, label_name)
    sent_count, skipped = 0, []

    with st.spinner("Sending emails..."):
        for idx, row in df.iterrows():
            to_addr = row.get("Email")
            if not to_addr:
                skipped.append(row)
                continue
            subject = subject_template.format(**row)
            html = convert_bold(body_template.format(**row))
            message = MIMEText(html, "html")
            message["To"] = to_addr
            message["Subject"] = subject
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
            msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                if label_id:
                    service.users().messages().modify(userId="me", id=sent_msg["id"], body={"addLabelIds": [label_id]}).execute()

            time.sleep(random.uniform(delay * 0.9, delay * 1.1))
            sent_count += 1

    # Save CSV Backup
    csv_path = f"/tmp/mail_merge_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    df.to_csv(csv_path, index=False)
    with open(csv_path, "rb") as f:
        st.download_button("⬇️ Download Updated CSV", f, file_name=os.path.basename(csv_path))

    send_backup_csv(service, csv_path)
    st.success(f"✅ Completed sending {sent_count} emails. Skipped {len(skipped)} rows.")
