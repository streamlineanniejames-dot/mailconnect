# ========================================
# Gmail Mail Merge Tool - Modern UI Edition (Resume Fix + Preview Restored)
# ========================================
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")

with st.sidebar:
    st.image("logo.png", width=180)
    st.markdown("---")
    st.markdown("### 📧 Gmail Mail Merge Tool")
    st.markdown("A powerful Gmail-based mail merge app with batch send, resume, and follow-up support.")
    st.markdown("---")
    st.markdown("**Quick Links:**")
    st.markdown("- 🏠 Home")
    st.markdown("- 🔁 New Run / Reset")
    st.markdown("- 🗂️ Merge History")
    st.markdown("---")
    st.caption("Developed by Ranjith")

st.markdown("<h1 style='text-align:center;'>📧 Gmail Mail Merge Tool</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:gray;'>with Follow-up Replies, Draft Save & Resume Support</p>", unsafe_allow_html=True)
st.markdown("---")

# ========================================
# Gmail API Setup
# ========================================
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

# ========================================
# Constants
# ========================================
DONE_FILE = "/tmp/mailmerge_done.json"
STATE_FILE = "/tmp/mailmerge_state.json"
BATCH_SIZE_DEFAULT = 50
DRAFT_BATCH_SIZE_DEFAULT = 110

# ========================================
# Helper Functions
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8;text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"""
    <html><body style="font-family: Verdana, Arial, sans-serif; font-size: 14px; line-height: 1.6;">
        {text}
    </body></html>
    """

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        new_label = service.users().labels().create(
            userId="me",
            body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
        ).execute()
        return new_label["id"]
    except Exception:
        return None

def fetch_message_id_header(service, message_id):
    for _ in range(6):
        try:
            msg = service.users().messages().get(
                userId="me", id=message_id, format="metadata", metadataHeaders=["Message-ID"]
            ).execute()
            headers = msg.get("payload", {}).get("headers", [])
            for h in headers:
                if h.get("name", "").lower() == "message-id":
                    return h.get("value")
        except Exception:
            pass
        time.sleep(random.uniform(1, 2))
    return ""

def send_email_backup(service, csv_path):
    try:
        user_email = service.users().getProfile(userId="me").execute()["emailAddress"]
        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        msg.attach(MIMEText("Attached is the backup CSV for your mail merge run.", "plain"))
        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        st.info(f"📧 Backup CSV emailed to {user_email}")
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

# ========================================
# OAuth Flow
# ========================================
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
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Main UI
# ========================================
if not st.session_state.get("sending", False):
    st.subheader("📤 Step 1: Upload Recipient List")
    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df = pd.read_csv(uploaded, encoding="utf-8")
            else:
                df = pd.read_excel(uploaded)
        except Exception:
            uploaded.seek(0)
            df = pd.read_csv(uploaded, encoding="latin1")

        for c in ["ThreadId", "RfcMessageId", "Status"]:
            if c not in df.columns:
                df[c] = ""

        st.markdown("### ✏️ Edit Your List")
        df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        subject_template = st.text_input("✉️ Subject", "Hello {Name}")
        body_template = st.text_area("📝 Body (Markdown + Variables like {Name})",
                                     "Dear {Name},\n\nWelcome to **Mail Merge App** demo.\n\nThanks,\n**Your Company**")

        label_name = st.text_input("🏷️ Gmail Label", "Mail Merge Sent")
        delay = st.slider("⏱️ Delay (seconds)", 20, 75, 20)
        mode = st.radio("📬 Send Mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        # Preview restored
        if not df.empty:
            r = df.iloc[0]
            try:
                ps = subject_template.format(**r)
                pb = convert_bold(body_template.format(**r))
            except Exception as e:
                ps = subject_template
                pb = body_template
                st.warning(f"⚠️ Could not render preview: {e}")
            st.markdown("---")
            st.markdown("### 👀 Email Preview (First Row)")
            st.markdown(f"**Subject:** {ps}")
            st.markdown(pb, unsafe_allow_html=True)

        if st.button("🚀 Start Mail Merge"):
            df = df.fillna("")
            pending = df.index[~df["Status"].isin(["Sent", "Draft"])].tolist()
            st.session_state.update({
                "sending": True,
                "df": df,
                "pending": pending,
                "subject": subject_template,
                "body": body_template,
                "label": label_name,
                "delay": delay,
                "mode": mode,
            })
            st.rerun()

# ========================================
# Sending Section (with Resume)
# ========================================
if st.session_state.get("sending"):
    df = st.session_state["df"]
    pending = st.session_state["pending"]
    subject_template = st.session_state["subject"]
    body_template = st.session_state["body"]
    label_name = st.session_state["label"]
    delay = st.session_state["delay"]
    mode = st.session_state["mode"]

    st.subheader("📨 Sending Emails... (Resumable)")
    progress = st.progress(0)
    status_box = st.empty()

    if os.path.exists(STATE_FILE):
        try:
            state = json.load(open(STATE_FILE))
            last_index = state.get("last_index", -1)
            pending = [i for i in pending if i > last_index]
            df = pd.read_csv(state.get("csv_path", ""), encoding="utf-8")
        except Exception:
            pass

    total = len(pending)
    sent_count, errors = 0, []
    label_id = None
    if mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    for i, idx in enumerate(pending):
        pct = int(((i + 1) / total) * 100)
        progress.progress(min(max(pct, 0), 100))
        status_box.info(f"📩 Processing {i + 1}/{total}")

        row = df.loc[idx]
        to_addr = extract_email(str(row.get("Email", "")).strip())
        if not to_addr:
            df.loc[idx, "Status"] = "Skipped"
            continue

        try:
            subject = subject_template.format(**row)
            body_html = convert_bold(body_template.format(**row))
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject

            raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
            msg_body = {"raw": raw}

            if mode == "💾 Save as Draft":
                service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                df.loc[idx, "Status"] = "Draft"
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                df.loc[idx, "RfcMessageId"] = fetch_message_id_header(service, sent_msg.get("id", ""))
                df.loc[idx, "Status"] = "Sent"

            sent_count += 1
            json.dump({"last_index": idx, "csv_path": "/tmp/mailmerge_temp.csv"}, open(STATE_FILE, "w"))
            df.to_csv("/tmp/mailmerge_temp.csv", index=False)
            time.sleep(random.uniform(delay * 0.9, delay * 1.1))
        except Exception as e:
            df.loc[idx, "Status"] = "Error"
            errors.append((to_addr, str(e)))

    # Save & backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = f"/tmp/Updated_{re.sub(r'[^A-Za-z0-9_-]', '_', label_name)}_{timestamp}.csv"
    df.to_csv(file_path, index=False)
    send_email_backup(service, file_path)
    json.dump({"done_time": str(datetime.now()), "file": file_path}, open(DONE_FILE, "w"))
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)

    st.session_state["sending"] = False
    st.session_state["done"] = True
    st.session_state["summary"] = {"sent": sent_count, "errors": errors}
    st.rerun()

# ========================================
# Completion Summary
# ========================================
if st.session_state.get("done"):
    s = st.session_state["summary"]
    st.subheader("✅ Mail Merge Completed")
    st.success(f"Sent: {s.get('sent', 0)}")
    if s.get("errors"):
        st.warning(f"⚠️ Errors: {len(s['errors'])}")
    st.download_button("⬇️ Download Final CSV", open(json.load(open(DONE_FILE))["file"], "rb"), file_name="Updated_Merge.csv")
    if st.button("🔁 New Run / Reset"):
        for f in [DONE_FILE, STATE_FILE]:
            if os.path.exists(f): os.remove(f)
        st.session_state.clear()
        st.experimental_rerun()
