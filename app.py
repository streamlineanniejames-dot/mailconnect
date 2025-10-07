import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import os
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="📧 Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge - Streamlit Cloud Ready")

# ========================================
# Gmail API Scopes
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
]

# ========================================
# Write credentials.json from st.secrets
# ========================================
if not os.path.exists("credentials.json"):
    with open("credentials.json", "w") as f:
        f.write(st.secrets["gmail"]["credentials_json"])

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
# Email Regex
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")
def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

# ========================================
# Gmail Label Helper (Stable)
# ========================================
def get_or_create_label(service, label_name="Mail Merge Sent"):
    """Get or create Gmail label reliably."""
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        existing = next((l for l in labels if l["name"].lower() == label_name.lower()), None)
        if existing:
            return existing["id"]

        new_label = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        created = service.users().labels().create(userId="me", body=new_label).execute()
        time.sleep(2)  # wait for Gmail to sync
        # confirm label creation
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        confirmed = next((l for l in labels if l["name"].lower() == label_name.lower()), None)
        return confirmed["id"] if confirmed else created["id"]
    except Exception as e:
        st.error(f"⚠️ Label error: {e}")
        return None

# ========================================
# Convert **bold** and [link](url) to
