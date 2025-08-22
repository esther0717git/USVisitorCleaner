import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter


# ───── 1) Info Banner ──────────────────────────────────────────────────────────
st.info(
    """
    **Data Integrity Is Our Foundation**  
    At every step—from file upload to final report—we enforce strict validation to guarantee your visitor data is accurate, complete, and compliant.  
    Maintaining integrity not only expedites gate clearance, it protects our facilities and ensures we meet all regulatory requirements.
    """
)

# ───── 2) Why Data Integrity? ───────────────────────────────────────────────────
with st.expander("Why is Data Integrity Important?"):
    st.write(
        """
        - **Accuracy**: Correct visitor details reduce clearance delays.  
        - **Security**: Reliable ID checks prevent unauthorized access.  
        - **Compliance**: Audit-ready records ensure regulatory adherence.  
        - **Efficiency**: Trustworthy data powers faster reporting and analytics.
        """
    )

# ───── Download Sample Template ────────────────────────────────────────────────
# This reads the Excel you committed as sample_template.xlsx in your repo root
#with open("sample_template.xlsx", "rb") as f:
#    sample_bytes = f.read()
#st.download_button(
#    label="🌟 Download Template",
#    data=sample_bytes,
#    file_name="sample_template.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#)

# ───── 3) Uploader & Warning ───────────────────────────────────────────────────

st.markdown(
    """<div style='font-size:14px; font-weight:bold; color:#38761d;'>
    Please ensure your spreadsheet has no missing or malformed fields.<br>
    Columns E and F are not required to be filled in.
    </div>""",
    unsafe_allow_html=True
)

uploaded = st.file_uploader("📁 Upload file", type=["xlsx"])


# ───── Streamlit setup ────────────────────────────────────────────────────────
st.set_page_config(page_title="Visitor List Cleaner (US)", layout="wide")
st.title("🇺🇸 CLARITY GATE - US VISITOR DATA CLEANING & VALIDATION 🫧")

# ───── Download Sample Template ───────────────────────────────────────────────
with open("us_template.xlsx", "rb") as f:
    st.download_button(
        label="📎 Download US Sample Template",
        data=f,
        file_name="us_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ───── Helper functions ────────────────────────────────────────────────────────

def split_name(full_name):
    s = str(full_name).strip()
    if " " in s:
        i = s.find(" ")
        return pd.Series([s[:i], s[i+1:]])
    return pd.Series([s, ""])

def clean_gender(g):
    v = str(g).strip().upper()
    if v == "M":
        return "Male"
    if v == "F":
        return "Female"
    if v in ("MALE","FEMALE"):
        return v.title()
    return v.title()

def fix_mobile(x):
    d = re.sub(r"\D", "", str(x))
    # if too long...
    if len(d) > 10:
        extra = len(d) - 10
        if d.endswith("0" * extra):
            d = d[:-extra]
        else:
            d = d[-10:]
    if len(d) < 10:
        d = d.zfill(10)
    return d

# ───── Core Cleaning Logic ────────────────────────────────────────────────────

def clean_data_us(df: pd.DataFrame) -> pd.DataFrame:
    # 1) Trim to exactly 10 cols then rename
    df = df.iloc[:, :10]
    df.columns = [
        "S/N",
        "Vehicle Plate Number",
        "Company Full Name",
        "Full Name",
        "First Name",
        "Middle and Last Name",
        "Driver License Number",
        "Nationality (Country Name)",
        "Gender",
        "Mobile Number",
    ]

    # 2) Drop rows where all of Full Name → Mobile are blank
    df = df.dropna(subset=df.columns[3:10], how="all")

    # 3) Normalize nationality (incl. Indian → India, etc.)
    nat_map = {
        "chinese":     "China",
        "singaporean": "Singapore",
        "malaysian":   "Malaysia",
        "indian":      "India",
        "usa":         "United States",
        "us":          "United States",
    }
    df["Nationality (Country Name)"] = (
        df["Nationality (Country Name)"]
          .astype(str)
          .str.strip()
          .str.lower()
          .replace(nat_map, regex=False)
          .str.title()
    )

    # 4) Sort by Company → Country → Full Name
    df = df.sort_values(
        ["Company Full Name", "Nationality (Country Name)", "Full Name"],
        ignore_index=True,
    )

    # 5) Reset S/N
    df["S/N"] = range(1, len(df) + 1)

    # 6) Standardize Vehicle Plate Number
    df["Vehicle Plate Number"] = (
        df["Vehicle Plate Number"]
          .astype(str)
          .str.replace(r"[\/,]", ";", regex=True)
          .str.replace(r"\s*;\s*", ";", regex=True)
          .str.strip()
          .replace("nan","", regex=False)
    )

    # 7) Proper-case & split names (in case the template didn't pre-split)
    df["Full Name"] = df["Full Name"].astype(str).str.title()
    df[["First Name","Middle and Last Name"]] = (
        df["Full Name"].apply(split_name)
    )

    # 8) Fix Mobile Number → 10 digits
    df["Mobile Number"] = df["Mobile Number"].apply(fix_mobile)

    # 9) Normalize gender
    df["Gender"] = df["Gender"].apply(clean_gender)

    # 10) Truncate Driver License Number to last 4 characters
    df["Driver License Number"] = df["Driver License Number"].astype(str).str.strip().str[-4:]

    return df

# ───── Build & style the single-sheet Excel ──────────────────────────────────

def generate_visitor_only_us(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Visitor List")
        ws = writer.sheets["Visitor List"]

        # styling objects
        header_fill  = PatternFill("solid", fgColor="94B455")
        border       = Border(Side("thin"),Side("thin"),Side("thin"),Side("thin"))
        center       = Alignment("center","center")
        normal_font  = Font(name="Calibri", size=9)
        bold_font    = Font(name="Calibri", size=9, bold=True)

        # 1) Apply borders, alignment, font
        for row in ws.iter_rows():
            for cell in row:
                cell.border    = border
                cell.alignment = center
                cell.font      = normal_font

        # 2) Style header row
        for col in range(1, ws.max_column + 1):
            h = ws[f"{get_column_letter(col)}1"]
            h.fill = header_fill
            h.font = bold_font

        # 3) Freeze top row
        ws.freeze_panes = ws["A2"]

        # 4) Auto-fit columns & set row height
        for col in ws.columns:
            w = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[get_column_letter(col[0].column)].width = w + 2
        for row in ws.iter_rows():
            ws.row_dimensions[row[0].row].height = 20

        # 5) Vehicles summary
        plates = []
        for v in df["Vehicle Plate Number"].dropna():
            plates += [x.strip() for x in str(v).split(";") if x.strip()]
        ins = ws.max_row + 2
        if plates:
            ws[f"B{ins}"].value     = "Vehicles"
            ws[f"B{ins}"].border    = border
            ws[f"B{ins}"].alignment = center
            ws[f"B{ins+1}"].value   = ";".join(sorted(set(plates)))
            ws[f"B{ins+1}"].border  = border
            ws[f"B{ins+1}"].alignment = center
            ins += 2

        # 6) Total Visitors
        ws[f"B{ins}"].value     = "Total Visitors"
        ws[f"B{ins}"].border    = border
        ws[f"B{ins}"].alignment = center
        ws[f"B{ins+1}"].value   = df["Company Full Name"].notna().sum()
        ws[f"B{ins+1}"].border  = border
        ws[f"B{ins+1}"].alignment = center

    buf.seek(0)
    return buf

# ───── Streamlit UI: Upload & Download ────────────────────────────────────────
uploaded = st.file_uploader("📁 Upload your US-template Excel", type=["xlsx"])
if uploaded:
    raw_df = pd.read_excel(uploaded, sheet_name="Visitor List")
    cleaned = clean_data_us(raw_df)
    out_buf = generate_visitor_only_us(cleaned)

    # Build filename: CompanyName_YYYYMMDD.xlsx in US/Eastern time
    today = datetime.now(ZoneInfo("America/New_York")).strftime("%Y%m%d")
    company_cell = raw_df.iloc[0, 2]
    company = (
        str(company_cell).strip()
        if pd.notna(company_cell) and str(company_cell).strip()
        else "VisitorList"
    )
    fname = f"{company}_{today}.xlsx"

    st.download_button(
        label="📥 Download Cleaned Visitor List (US)",
        data=out_buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
