# Fleet Speeding Alert Summary

A simple app: upload your weekly **Vehicle Exception Report**, download the **Driver Summary**.

## One-time setup
1. Install Python 3 (python.org) — tick "Add Python to PATH" during install.
2. Put `speeding_app.py`, `requirements.txt` (and `run_windows.bat`) in one folder.

## Run it each week
- **Windows:** double-click **`run_windows.bat`**.
- **Or any OS**, in a terminal in that folder:
  ```
  pip install -r requirements.txt
  streamlit run speeding_app.py
  ```
A browser tab opens. **Drag your raw CSV in**, then click **Download Driver Summary (Excel)**.

## What it does
- Flags a speeding event only when the speed reaches:
  20→35, 30→44, 40→56, 50→66, 60→76, 70→91 mph.
- Groups flagged events within **15 minutes** (per driver, across the week) into **one real alert**.
- Output sheet **Driver Summary**: Driver, Vehicle Reg, Speeding events, Real Alerts, and Real Alerts per day.
- Also includes **Flagged detail** and **Rules** sheets.

The 15-minute window can be changed in the sidebar.
