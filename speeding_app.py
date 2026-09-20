"""
Fleet Speeding Alert Summary — weekly tool
Upload the raw Vehicle Exception Report; download the Driver Summary.

Run it:
    pip install streamlit pandas openpyxl
    streamlit run speeding_app.py
Then use the browser page that opens (drag your weekly CSV in).
"""
import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---- speeding rules: flag when Speed >= threshold for that limit ----
THRESHOLDS = {20: 35, 30: 44, 40: 56, 50: 66, 60: 76, 70: 91}
#   20 -> 35+   30 -> 44+ (exceeds 43)   40 -> 56+ (exceeds 55)
#   50 -> 66+   60 -> 76+                70 -> 91+ (exceeds 90)

TEAL = "0E6E63"; DTEAL = "0A4A43"; WHITE = "FFFFFF"; BORD = "D5DEDC"; ROW2 = "F5F9F8"
thin = Side(style="thin", color=BORD); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def load_raw(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded, dtype=str)
    else:
        df = pd.read_csv(uploaded, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def process(df: pd.DataFrame, window_min: int = 15):
    need = ["Driver", "Vehicle", "Date", "Time", "Speed", "Speed Limit"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected column(s): {', '.join(missing)}")

    d = df.copy()
    for c in ["Driver", "Vehicle"]:
        d[c] = d[c].astype(str).str.strip()
    d["Speed"] = pd.to_numeric(d["Speed"], errors="coerce")
    d["Speed Limit"] = pd.to_numeric(d["Speed Limit"], errors="coerce")
    d["dt"] = pd.to_datetime(d["Date"].astype(str).str.strip() + " " + d["Time"].astype(str).str.strip(),
                             dayfirst=True, errors="coerce")
    d = d[d["dt"].notna() & d["Speed"].notna() & d["Speed Limit"].notna() & (d["Driver"] != "")]

    # flag
    d["_thr"] = d["Speed Limit"].map(THRESHOLDS)
    d["Flagged"] = d["_thr"].notna() & (d["Speed"] >= d["_thr"])
    flagged = d[d["Flagged"]].sort_values(["Driver", "dt"]).copy()

    # continuous 15-min rule per driver -> a new alert when gap > window
    gap = flagged.groupby("Driver")["dt"].diff().dt.total_seconds()
    flagged["New alert"] = ((gap.isna()) | (gap > window_min * 60)).astype(int)
    flagged["day"] = flagged["dt"].dt.date

    days = sorted(flagged["day"].unique())
    day_cols = {dd: dd.strftime("%a %d/%m") for dd in days}

    rows = []
    for drv, g in flagged.groupby("Driver"):
        veh = g["Vehicle"].mode()
        veh = veh.iloc[0] if len(veh) else ""
        row = {"Driver": drv, "Vehicle Reg": veh,
               "Speeding events": int(len(g)), "Real Alerts": int(g["New alert"].sum()), "": ""}
        per_day = g.groupby("day")["New alert"].sum()
        for dd in days:
            row[day_cols[dd]] = int(per_day.get(dd, 0))
        rows.append(row)

    cols = ["Driver", "Vehicle Reg", "Speeding events", "Real Alerts", ""] + [day_cols[dd] for dd in days]
    summary = pd.DataFrame(rows, columns=cols).sort_values("Speeding events", ascending=False).reset_index(drop=True)

    detail = flagged[["Vehicle", "Driver", "Date", "Time", "Speed", "Speed Limit", "New alert"]].copy()
    meta = {"events": int(len(d)), "flagged": int(len(flagged)),
            "alerts": int(flagged["New alert"].sum()), "drivers": summary.shape[0],
            "date_from": min(days).strftime("%d/%m/%Y") if days else "-",
            "date_to": max(days).strftime("%d/%m/%Y") if days else "-"}
    return summary, detail, meta


def to_excel(summary: pd.DataFrame, detail: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        summary.to_excel(xl, sheet_name="Driver Summary", index=False)
        detail.to_excel(xl, sheet_name="Flagged detail", index=False)
        rules = pd.DataFrame({"Speed limit": list(THRESHOLDS), "Flag when speed is at least": list(THRESHOLDS.values())})
        rules.to_excel(xl, sheet_name="Rules", index=False)
        for sn in xl.book.sheetnames:
            ws = xl.book[sn]
            for j, cell in enumerate(ws[1], 1):
                cell.font = Font("Calibri", bold=True, color=WHITE, size=10)
                cell.fill = PatternFill("solid", fgColor=TEAL)
                cell.alignment = Alignment("center", "center", wrap_text=True); cell.border = BORDER
            ws.row_dimensions[1].height = 26
            for col in ws.columns:
                L = col[0].column_letter
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[L].width = min(max(width + 2, 10), 30)
            ws.freeze_panes = "A2"
    return buf.getvalue()


# ------------------------------- UI -------------------------------
st.set_page_config(page_title="Fleet Speeding Alerts", page_icon="🚗", layout="wide")
st.title("🚗 Fleet speeding alert summary")
st.caption("Upload this week's Vehicle Exception Report and download the driver alert summary.")

with st.sidebar:
    st.subheader("Rules")
    st.markdown(
        "**Flag a speeding event when speed reaches:**\n"
        "- 20 → 35 mph\n- 30 → 44 mph\n- 40 → 56 mph\n- 50 → 66 mph\n- 60 → 76 mph\n- 70 → 91 mph"
    )
    window_min = st.number_input("Group events within this many minutes as one alert",
                                 min_value=1, max_value=60, value=15, step=1)
    st.caption("Consecutive flagged events inside this window count as a single real alert.")

up = st.file_uploader("Drop your raw report here (CSV or Excel)", type=["csv", "xlsx", "xls"])

if up is not None:
    try:
        raw = load_raw(up)
        summary, detail, meta = process(raw, window_min)
    except Exception as e:
        st.error(f"Could not process the file: {e}")
        st.stop()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Drivers", meta["drivers"])
    c2.metric("Flagged events", f"{meta['flagged']:,}")
    c3.metric("Real alerts", f"{meta['alerts']:,}")
    c4.metric("Week", f"{meta['date_from']} – {meta['date_to']}")

    st.subheader("Driver Summary")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Download Driver Summary (Excel)",
        data=to_excel(summary, detail),
        file_name=f"Driver_Speeding_Summary_{meta['date_from'].replace('/','-')}_to_{meta['date_to'].replace('/','-')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Waiting for a file… drag your weekly CSV above.")
