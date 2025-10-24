
import pytz, pandas as pd
from datetime import datetime

CET_TZ = pytz.timezone("CET")

COMM_TYPES = {
    "Day before the release": "day_before",
    "30 minutes before the release": "t_minus_30",
    "At the beginning of the release": "start",
    "At the end of the release": "end",
}

REQUIRED_FIELDS = {
    "day_before": {"date": True,  "start": False, "end": False},
    "t_minus_30": {"date": True,  "start": False, "end": False},
    "start":      {"date": True,  "start": True,  "end": True},
    "end":        {"date": True,  "start": False, "end": False},
}

def build_subject(app_name: str, comm_key: str, full_ref: str, date_val: datetime|None, start_dt: datetime|None, end_dt: datetime|None, purpose_text: str) -> str:
    app = app_name or "Application"
    purpose_short = (purpose_text or "Purpose of change").split("\n")[0][:80]
    date_str = date_val.strftime("%d/%m/%Y") if date_val else ""
    if comm_key == "day_before":
        return f"{app} [{full_ref}] [Planned on {date_str}] - Change Release - {purpose_short}"
    elif comm_key == "t_minus_30":
        return f"{app} [{full_ref}] [Starting in 30 minutes on {date_str}] - Change Release - {purpose_short}"
    elif comm_key == "start":
        if start_dt and end_dt:
            ts = f"{start_dt.strftime('%H:%M')}–{end_dt.strftime('%H:%M')} CET on {start_dt.strftime('%d/%m/%Y')}"
        elif start_dt:
            ts = f"{start_dt.strftime('%H:%M')} CET on {start_dt.strftime('%d/%m/%Y')}"
        else:
            ts = ""
        return f"{app} [{full_ref}] [Starting now {ts}] - Change Release - {purpose_short}".strip()
    elif comm_key == "end":
        return f"{app} [{full_ref}] [Release on {date_str} has ended] - Change Release - {purpose_short}"
    else:
        return f"{app} [{full_ref}] - Change Release - {purpose_short}"

AVAIL_SENTENCE = {
    True:  "The application will be available during the delivery.",
    False: "The application will not be available during the delivery.",
}

def normalize_dist_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return df
    cols_map = {}
    for c in df.columns:
        cl = c.lower()
        if cl.startswith("application"):
            cols_map[c] = "Applications"
        elif cl == "cc":
            cols_map[c] = "cc"
        elif cl in ("cci", "bcc"):
            cols_map[c] = "cci"
        elif cl.startswith("support"):
            cols_map[c] = "support mail"
    return df.rename(columns=cols_map)
