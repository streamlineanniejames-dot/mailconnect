# ========================================
# Send Emails — Inbox-Optimized Version
# ========================================
if st.button("🚀 Send Emails"):
    label_id = get_or_create_label(service, label_name)
    sent_count = 0
    skipped = []
    errors = []

    with st.spinner("📨 Sending emails... please wait."):
        for idx, row in df.iterrows():
            to_addr_raw = str(row.get("Email", "")).strip()
            to_addr = extract_email(to_addr_raw)
            if not to_addr:
                skipped.append(to_addr_raw)
                continue

            try:
                # Prepare subject and body
                subject = subject_template.format(**row)
                body_text = body_template.format(**row)

                # Add invisible preheader for inbox credibility
                preheader = f"Hello {row.get('Name', '')}, this is a personalized message from Your Company."
                html_body = f"""
                <html>
                  <body style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
                    <span style="display:none; color:transparent; max-height:0; max-width:0; opacity:0; overflow:hidden;">
                      {preheader}
                    </span>
                    {convert_bold(body_text)}
                    <!-- ID:{idx} -->
                  </body>
                </html>
                """

                # Build email
                message = MIMEText(html_body, "html")
                message["to"] = to_addr
                message["subject"] = subject
                message["from"] = "Your Name <youremail@gmail.com>"
                message["reply-to"] = "youremail@gmail.com"

                # Encode and send
                raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                msg_body = {"raw": raw}
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                # Apply label after sending
                if label_id:
                    service.users().messages().modify(
                        userId="me",
                        id=sent_msg["id"],
                        body={"addLabelIds": [label_id]},
                    ).execute()

                sent_count += 1
                time.sleep(30)  # 30-second gap between emails

            except Exception as e:
                errors.append((to_addr, str(e)))

    # Summary
    st.success(f"✅ Successfully sent {sent_count} emails.")
    if skipped:
        st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
    if errors:
        st.error(f"❌ Failed to send {len(errors)} emails: {errors}")
