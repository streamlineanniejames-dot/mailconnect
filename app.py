# ======================================== 
# Gmail Mail Merge Tool - Modern UI Edition (Duplicate Send Fix + Email Preview)
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
    st.markdown("### 📧 Gmail Mail Merge Tool")
    st.caption("Batch Gmail sender with draft, follow-up & resume.")
    st.markdown("---")
    st.caption("Developed by Ranjith")

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

DONE_FILE = "/tmp/mailmerge_done.json"
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
        r'<a href="\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"<html><body style='font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;'>{text}</body></html>"

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        created_label = service.users().labels().create(
            userId="me",
            body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
        ).execute()
        return created_label["id"]
    except Exception:
        return None

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
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

def fetch_message_id_header(service, message_id):
    for _ in range(6):
        try:
            msg_detail = service.users().messages().get(
                userId="me", id=message_id, format="metadata", metadataHeaders=["Message-ID"]
            ).execute()
            headers = msg_detail.get("payload", {}).get("headers", [])
            for h in headers:
                if h.get("name", "").lower() == "message-id":
                    return h.get("value")
        except Exception:
            pass
        time.sleep(random.uniform(1, 2))
    return ""

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
# App State Machine
# ========================================
if "step" not in st.session_state:
    st.session_state["step"] = "upload"

# ---------- STEP 1: Upload & Template ----------
if st.session_state["step"] == "upload":
    st.header("📤 Step 1: Upload Recipient List & Create Template")
    uploaded_file = st.file_uploader("Upload your CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
        except Exception:
            df = pd.read_excel(uploaded_file)

        for col in ["ThreadId", "RfcMessageId", "Status"]:
            if col not in df.columns:
                df[col] = ""

        st.dataframe(df.head())

        subject_template = st.text_input("✉️ Subject", "Hello {Name}")
        body_template = st.text_area(
            "📝 Body (Markdown supported)",
            """Dear {Name},

Welcome to **Mail Merge App** demo.

Thanks,  
**Your Company**""",
            height=250,
        )

        # 🔍 Email preview (first row)
        if not df.empty:
            preview_row = df.iloc[0]
            try:
                preview_subject = subject_template.format(**preview_row)
                preview_body = convert_bold(body_template.format(**preview_row))
            except Exception as e:
                preview_subject = subject_template
                preview_body = body_template
                st.warning(f"⚠️ Could not render preview: {e}")

            st.markdown("---")
            st.subheader("👀 Email Preview (First Row)")
            st.markdown(f"**Subject:** {preview_subject}")
            st.markdown(preview_body, unsafe_allow_html=True)

        st.markdown("---")
        label_name = st.text_input("🏷️ Gmail Label", "Mail Merge Sent")
        delay = st.slider("⏱️ Delay between emails (seconds)", 10, 60, 20)
        send_mode = st.radio("📬 Send Mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        if st.button("🚀 Start Mail Merge"):
            df = df.fillna("")
            pending_indices = df.index[~df["Status"].isin(["Sent", "Draft", "Error"])].tolist()
            st.session_state.update({
                "step": "sending",
                "df": df,
                "pending_indices": pending_indices,
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode,
            })
            st.rerun()

# ---------- STEP 2: Sending ----------
elif st.session_state["step"] == "sending":
    df = st.session_state["df"]
    pending_indices = st.session_state["pending_indices"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]

    st.header("📨 Sending Emails...")
    progress = st.progress(0)
    status = st.empty()

    total = len(pending_indices)
    sent_count = 0
    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    sent_message_ids = []
    batch_limit = DRAFT_BATCH_SIZE_DEFAULT if send_mode == "💾 Save as Draft" else BATCH_SIZE_DEFAULT

    for i, idx in enumerate(pending_indices):
        if i >= batch_limit:
            break

        row = df.loc[idx]
        to_addr = extract_email(row.get("Email", ""))
        if not to_addr:
            df.loc[idx, "Status"] = "Skipped"
            continue

        try:
            subject = subject_template.format(**row)
            body_html = convert_bold(body_template.format(**row))
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject

            msg_body = {"raw": base64.urlsafe_b64encode(message.as_bytes()).decode()}

            if send_mode == "💾 Save as Draft":
                service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                df.loc[idx, "Status"] = "Draft"
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                msg_id = sent_msg.get("id", "")
                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                df.loc[idx, "RfcMessageId"] = fetch_message_id_header(service, msg_id) or msg_id
                df.loc[idx, "Status"] = "Sent"
                if label_id:
                    sent_message_ids.append(msg_id)

            sent_count += 1
            progress.progress(int((i + 1) / total * 100))
            status.info(f"📩 Sent {sent_count}/{total} - {to_addr}")
            time.sleep(random.uniform(delay * 0.9, delay * 1.1))
        except Exception as e:
            df.loc[idx, "Status"] = "Error"
            st.error(f"Error sending to {to_addr}: {e}")

    # Labeling + Backup
    if sent_message_ids and label_id:
        try:
            service.users().messages().batchModify(
                userId="me", body={"ids": sent_message_ids, "addLabelIds": [label_id]}
            ).execute()
        except Exception as e:
            st.warning(f"⚠️ Label apply failed: {e}")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"Updated_{label_name}_{timestamp}.csv"
    file_path = os.path.join("/tmp", file_name)
    df.to_csv(file_path, index=False)

    try:
        send_email_backup(service, file_path)
    except Exception:
        pass

    st.session_state["summary"] = {"sent": sent_count}
    st.session_state["step"] = "done"
    st.session_state["df"] = df
    st.session_state["csv_path"] = file_path
    st.rerun()

# ---------- STEP 3: Done ----------
elif st.session_state["step"] == "done":
    st.header("✅ Mail Merge Completed")
    summary = st.session_state["summary"]
    st.success(f"Sent: {summary['sent']}")
    with open(st.session_state["csv_path"], "rb") as f:
        st.download_button("⬇️ Download Updated CSV", data=f, file_name=os.path.basename(st.session_state["csv_path"]))
    if st.button("🔁 New Run / Reset"):
        st.session_state.clear()
        st.experimental_rerun()
