import streamlit as st
from libratom.lib.pff import PffArchive
from bs4 import BeautifulSoup
import re
import os

# ======== CONFIG ========
PST_PATH = r"c:\Users\irt\Downloads\dhruv-outlook-official.pst"
# =========================

# Helpers
def extract_header_field(headers, field_name):
    headers = headers or ""
    pattern = rf"^{field_name}:\s*(.*)$"
    match = re.search(pattern, headers, re.MULTILINE | re.IGNORECASE)
    return match.group(1).strip() if match else f"(No {field_name})"

def clean_html(html):
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text()

def get_message_body(message):
    try:
        body = getattr(message, "plain_text_body", None)
        if body and body.strip():
            return body.strip()

        html_body = getattr(message, "html_body", None)
        if html_body and html_body.strip():
            return clean_html(html_body.strip())

        rtf_body = getattr(message, "rtf_body", None)
        if rtf_body and rtf_body.strip():
            return f"[RTF Body]\n{rtf_body.strip()}"

        return "(No Body Content Found)"
    except Exception as e:
        return f"(Error reading body: {e})"

def safe_getattr(obj, attr, default="(Unavailable)"):
    try:
        return getattr(obj, attr, default)
    except Exception as e:
        return f"{default} - Error: {e}"

# Streamlit UI
st.set_page_config(page_title="📬 PST Email Viewer", layout="wide")
st.title("📬 PST Email Viewer")

if not os.path.exists(PST_PATH):
    st.error(f"PST file not found at:\n`{PST_PATH}`")
else:
    if st.button("📨 View Emails"):
        with st.spinner("Reading emails..."):
            with PffArchive(PST_PATH) as archive:
                count = 0
                for message in archive.messages():
                    count += 1
                    subject = safe_getattr(message, "subject", "(No Subject)")
                    sender = safe_getattr(message, "sender_name", "(Unknown Sender)")
                    sender_email = safe_getattr(message, "sender_email_address", "(No Email)")
                    headers = safe_getattr(message, "transport_headers", "")
                    to = extract_header_field(headers, "To")
                    cc = extract_header_field(headers, "Cc")
                    sent_time = safe_getattr(message, "client_submit_time", "(No Date)")
                    body = get_message_body(message)

                    with st.expander(f"📧 {subject}"):
                        st.markdown(f"**From:** {sender} `<{sender_email}>`")
                        st.markdown(f"**To:** {to}")
                        st.markdown(f"**CC:** {cc}")
                        st.markdown(f"**Date:** {sent_time}")
                        st.markdown("---")
                        st.markdown(f"**Body:**\n\n```\n{body}\n```")

                if count == 0:
                    st.warning("No messages found in the PST file.")
