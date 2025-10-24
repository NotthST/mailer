
# Release Communication – V3 (bb Style Email)

- Subject starts with application name and follows your comm-type rules.
- Body matches bb look (white), with banner placeholder, Context + Release Content sections, and CMIT footer using the app mailbox (from CSV 'support mail').
- No upload. All text areas are manual.

## Run
```powershell
py -3.12 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Configure
- Edit `config.ini` → set `HEADER_IMG_URL` to your bb banner URL.
- Set SharePoint site + CSV path, or use a local CSV via `[local]`.
