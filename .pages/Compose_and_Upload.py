import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Compose & Upload", layout="wide")

st.title("📤 Step 1: Upload Recipients & Compose Email")

EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

uploaded_file = st.file_uploader("📁 Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())

    df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="recipient_editor_inline")

    st.subheader("✍️ Compose Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body (supports **bold**, [links](https://example.com), and placeholders like {Name})",
        """Dear {Name},

Welcome to our **Mail Merge App** demo.

Thanks,  
**Your Company**""",
        height=250,
    )

    if st.button("💾 Save & Proceed to Send Page"):
        st.session_state["recipients_df"] = df
        st.session_state["subject_template"] = subject_template
        st.session_state["body_template"] = body_template
        st.success("✅ Data and template saved! Go to the **Send Emails** page to continue.")
