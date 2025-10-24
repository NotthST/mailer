
import os
from datetime import datetime, date, time
import pandas as pd
import streamlit as st
import configparser

from sharepoint_utils import load_distribution_df
from subject_templates import CET_TZ, COMM_TYPES, REQUIRED_FIELDS, build_subject, AVAIL_SENTENCE, normalize_dist_columns
from mailer import create_outlook_draft

st.set_page_config(page_title="Release Communication – Draft Generator", layout="wide")

with st.spinner("Loading distribution list…"):
    dist_df = load_distribution_df()
dist_df = normalize_dist_columns(dist_df)
apps = sorted(dist_df.get("Applications", pd.Series(dtype=str)).dropna().unique().tolist())

st.title("Release Communication – Draft Generator (CET)")
app_choice = st.selectbox("Application", options=apps, index=0 if apps else None)

comm_label = st.radio("Communication type", options=list(COMM_TYPES.keys()))
comm_key = COMM_TYPES[comm_label]

full_ref = st.text_input("Release reference (full, e.g. CAGIPCHG0294114)")

req = REQUIRED_FIELDS.get(comm_key, {"date":False,"start":False,"end":False})
col1, col2, col3 = st.columns(3)
with col1:
    the_date = st.date_input("Date (CET)", disabled=not req["date"])
with col2:
    start_time = st.time_input("Start time (CET)", value=time(10,45), disabled=not req["start"])
with col3:
    end_time = st.time_input("End time (CET)", value=time(11,30), disabled=not req["end"])

avail = st.selectbox("Availability during delivery", ["Available", "Not available"]) == "Available"
rel_type = st.selectbox("Type", ["Functional", "Technical"])

release_title = st.text_input("Release title (e.g., Release note (MTM-Client))")
context_text = st.text_area("1) Context – text (manual)", height=120)
context_bullets = st.text_area("1) Context – bullets (one per line)", height=120)
content_bullets = st.text_area("2) Release Content – bullets (one per line)", height=140)

# recipients
rec_to, rec_cc, rec_bcc = [], [], []
support_mailbox = ""
if app_choice and not dist_df.empty:
    row = dist_df.loc[dist_df["Applications"] == app_choice].head(1)
    if not row.empty:
        def split_emails(val):
            if pd.isna(val): return []
            return [e.strip() for e in str(val).replace(",", ";").split(";") if e.strip()]
        rec_to  = split_emails(row.iloc[0].get("support mail", ""))
        rec_cc  = split_emails(row.iloc[0].get("cc", ""))
        rec_bcc = split_emails(row.iloc[0].get("cci", ""))
        support_mailbox = rec_to[0] if rec_to else ""

# dates
date_dt  = datetime.combine(the_date, time(0,0)) if isinstance(the_date, date) else None
start_dt = datetime.combine(the_date, start_time) if req["start"] else None
end_dt   = datetime.combine(the_date, end_time) if req["end"] else None

subject = build_subject(app_choice or "Application", comm_key, full_ref or "", date_dt, start_dt, end_dt, context_text or "Purpose of change")

st.subheader("Preview")
st.text_input("Subject", value=subject, disabled=True)

# branding
CFG = configparser.ConfigParser(); CFG.read("config.ini")
header_img = CFG.get("branding", "HEADER_IMG_URL", fallback="")
border_col = CFG.get("branding", "BORDER_COLOR", fallback="#9AC28F")

def li_html(lines):
    items = [f"<li>{l.strip()}</li>" for l in (lines.splitlines() if lines else []) if l.strip()]
    return "\n".join(items)

header_img_html = f"<img src='{header_img}' alt='Company' style='height:40px;'>" if header_img else ""

availability_phrase = "the application will be available" if avail else "the application will not be available"

body_html = f"""
<html>
<body style='margin:0;padding:0;background:#ffffff;'>
  <table width='100%' cellpadding='0' cellspacing='0' style='font-family:Segoe UI,Arial;color:#1c2331;'>
    <tr><td style='padding:16px 18px;'>
      <table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #d9dfe6;'>
        <tr>
          <td style='padding:12px 16px;border-bottom:3px solid {border_col};'>
            {header_img_html}
          </td>
        </tr>
        <tr>
          <td style='padding:16px;'>
            <table width='100%' cellpadding='0' cellspacing='0' style='border:1px solid #cfd6de;'>
              <tr>
                <td style='padding:10px 14px;background:#f6f8fa;border-bottom:1px solid #cfd6de;'>
                  <div style='font-weight:600;font-size:18px;text-transform:capitalize;'>{app_choice or ''}</div>
                </td>
              </tr>
              <tr>
                <td style='padding:16px;'>
                  <p style='margin:0 0 8px 0;'>Dear {app_choice or ''} Users,</p>
                  <p style='margin:0 0 8px 0;'>This mail is to bring to your kind notice that a release will start in <b>{app_choice or ''}</b>.</p>
                  <p style='margin:0 0 8px 0;'><b>Reference:</b> {full_ref or ''}</p>
                  <p style='margin:0 0 8px 0;'><b>Start:</b> {(start_dt.strftime('%d/%m/%Y %H:%M CET') if start_dt else (date_dt.strftime('%d/%m/%Y') if date_dt else ''))}</p>
                  <p style='margin:0 0 8px 0;'><b>End:</b> {(end_dt.strftime('%d/%m/%Y %H:%M CET') if end_dt else '')}</p>
                  <p style='margin:0 0 14px 0;color:#18794e;'><b>Application availability during the release:</b> {availability_phrase} during the delivery</p>
                  <p style='margin:10px 0 6px 0;font-weight:600;'>1. Context</p>
                  <p style='margin:0 0 6px 0;'>{context_text or ''}</p>
                  <ul style='margin:0 0 10px 20px;'>
                    {li_html(context_bullets)}
                  </ul>
                  <p style='margin:14px 0 6px 0;font-weight:600;'>2. Release Content</p>
                  <ul style='margin:0 0 6px 20px;'>
                    {li_html(content_bullets)}
                  </ul>
                  <p style='margin:0 0 6px 0;'><b>Type:</b> {rel_type}</p>
                  <p style='margin:16px 0 0 0;font-size:12px;color:#6a7385;'>NB: All users are BCC intentional in this email.</p>
                  <p style='margin:8px 0 0 0;'>Please reach out to us in case you need any more information.</p>
                  <p style='margin:16px 0 0 0;'>Best Regards,</p>
                  <table cellpadding='0' cellspacing='0' style='margin-top:8px;border:1px solid #cfd6de;'>
                    <tr>
                      <td style='padding:10px 12px;background:#f6f8fa;border-right:1px solid #cfd6de;'>
                        <div style='font-weight:700;color:#1d7de3;'>GIT</div>
                      </td>
                      <td style='padding:10px 12px;'>
                        Company<br>
                        Global IT – CMIT Support<br>
                        For a more efficient CFI service level, write in English to  <a href='mailto:{support_mailbox}'>{support_mailbox}</a>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body>
</html>
"""
)

st.markdown("**Email body preview (white style)**")
st.components.v1.html(body_html, height=650, scrolling=True)

if st.button("Generate communication draft"):
    if not app_choice:
        st.error("Please select an application."); st.stop()
    if not full_ref.strip():
        st.error("Please enter the full release reference."); st.stop()
    if req["date"] and not isinstance(the_date, date):
        st.error("Please set the Date (CET)."); st.stop()
    if req["start"] and start_time is None:
        st.error("Please set the Start time."); st.stop()
    if req["end"] and end_time is None:
        st.error("Please set the End time."); st.stop()

    create_outlook_draft(to_list=rec_to, cc_list=rec_cc, bcc_list=rec_bcc, subject=subject, html_body=body_html)
    st.success("Draft generated in Outlook. You can add extra info and send.")
