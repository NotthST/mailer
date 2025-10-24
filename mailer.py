
import os
try:
    import win32com.client
except Exception:
    win32com = None

def create_outlook_draft(to_list, cc_list, bcc_list, subject, html_body):
    if win32com is None:
        raise RuntimeError("Outlook COM backend not available. Run on Windows with Outlook installed.")
    outlook = win32com.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To  = ";".join(to_list or [])
    mail.CC  = ";".join(cc_list or [])
    mail.BCC = ";".join(bcc_list or [])
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Save()
    mail.Display()
