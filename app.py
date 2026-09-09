import io
import re
import pandas as pd
import streamlit as st
from docx import Document
from pdf2image import convert_from_bytes
import pytesseract


def extract_text_with_ocr(pdf_bytes: bytes) -> list[str]:
    """Converts PDF pages into images and runs OCR to extract text from scanned forms."""
    images = convert_from_bytes(pdf_bytes)
    page_texts = []
    for img in images:
        text = pytesseract.image_to_string(img)
        page_texts.append(text)
    return page_texts


def process_toast_csv(csv_files) -> pd.DataFrame:
    """Combines and cleans Toast POS time entry CSV files."""
    dfs = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)
    df_all.columns = df_all.columns.str.strip()
    return df_all


def get_date_variants(date_str) -> list[str]:
    """Generates short date variants (e.g., '8/7', '08/07', '8/7/26') for OCR matching."""
    if pd.isna(date_str):
        return []
    
    try:
        dt = pd.to_datetime(date_str)
        m, d, y = dt.month, dt.day, str(dt.year)[-2:]
        return [
            f"{m}/{d}",
            f"{m:02d}/{d:02d}",
            f"{m}/{d}/{y}",
            f"{m:02d}/{d:02d}/{y}",
            f"{m}/{d}/{dt.year}"
        ]
    except Exception:
        return [str(date_str).strip()]


def clean_name_tokens(name_str: str) -> list[str]:
    """Splits Toast 'Last, First' names into individual searchable name tokens."""
    if pd.isna(name_str):
        return []
    # Clean common suffixes or punctuation
    parts = re.split(r"[\s,]+", str(name_str))
    # Filter out short tokens or common middle initial noises
    return [p.strip() for ppart in parts for p in [ppart] if len(p) > 2]


def analyze_edits(df: pd.DataFrame, ocr_pages: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Analyzes manager edits and checks against OCR PDF pages."""
    # Column mapping for Toast POS export
    mgr_col = next((c for c in df.columns if "Manager" in c or "Edited" in c), "Manager")
    emp_col = next((c for c in df.columns if "Employee" in c), "Employee")
    change_col = next((c for c in df.columns if "Change" in c), None)
    in_date_col = next((c for c in df.columns if "In Date" in c or "Date" in c), None)

    # Filter out 'CREATE' actions (keep edits/modifications)
    filtered = df.copy()
    if change_col and change_col in filtered.columns:
        filtered = filtered[filtered[change_col] != "CREATE"]

    # Exclude system and ghost entries
    system_terms = ["system", "ghost", "house", "auto", "toast"]
    pattern = "|".join(system_terms)
    
    filtered = filtered[
        ~filtered[emp_col].astype(str).str.lower().str.contains(pattern, na=False) &
        ~filtered[mgr_col].astype(str).str.lower().str.contains(pattern, na=False)
    ].dropna(subset=[mgr_col])

    details = []

    for _, row in filtered.iterrows():
        manager = str(row[mgr_col]).strip()
        employee = str(row[emp_col]).strip()
        raw_date = row[in_date_col] if in_date_col else None

        name_tokens = clean_name_tokens(employee)
        date_vars = get_date_variants(raw_date)

        matched = False
        for page_text in ocr_pages:
            # Check if at least 2 primary name tokens (e.g. First & Last) appear on the page
            name_hits = sum(1 for token in name_tokens if re.search(r"\b" + re.escape(token) + r"\b", page_text, re.IGNORECASE))
            has_name = name_hits >= min(2, len(name_tokens)) if name_tokens else False

            # Check date match on page
            has_date = any(re.search(re.escape(d), page_text) for d in date_vars) if date_vars else True

            if has_name and has_date:
                matched = True
                break

        details.append({
            "Manager": manager,
            "Employee": employee,
            "Edit_Date": str(raw_date) if raw_date else "N/A",
            "Has_Signed_Form": matched
        })

    details_df = pd.DataFrame(details)

    if details_df.empty:
        summary_df = pd.DataFrame(columns=["Manager", "Total_Edits", "Forms_Present", "Forms_Missing", "Compliance_Pct"])
        return summary_df, details_df

    # Aggregate by Manager
    summary_df = details_df.groupby("Manager").agg(
        Total_Edits=("Has_Signed_Form", "count"),
        Forms_Present=("Has_Signed_Form", "sum")
    ).reset_index()

    summary_df["Forms_Missing"] = summary_df["Total_Edits"] - summary_df["Forms_Present"]
    summary_df["Compliance_Pct"] = ((summary_df["Forms_Present"] / summary_df["Total_Edits"]) * 100).round(1)

    return summary_df, details_df


def create_word_docx(summary_df: pd.DataFrame, details_df: pd.DataFrame) -> io.BytesIO:
    """Generates audit summary report in Word document format."""
    doc = Document()
    doc.add_heading("Toast POS Time Edit Audit Report", level=1)

    doc.add_heading("1. Manager Summary", level=2)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text = (
        "Manager", "Total Edits", "Forms Present", "Forms Missing", "Compliance %"
    )

    for _, row in summary_df.iterrows():
        r = table.add_row().cells
        r[0].text, r[1].text, r[2].text, r[3].text, r[4].text = (
            str(row["Manager"]), str(row["Total_Edits"]), str(row["Forms_Present"]),
            str(row["Forms_Missing"]), f"{row['Compliance_Pct']}%"
        )

    doc.add_heading("2. Missing Forms Detail", level=2)
    missing = details_df[~details_df["Has_Signed_Form"]]
    if missing.empty:
        doc.add_paragraph("All edits have corresponding signed forms present.")
    else:
        m_table = doc.add_table(rows=1, cols=3)
        m_table.style = "Table Grid"
        m_hdr = m_table.rows[0].cells
        m_hdr[0].text, m_hdr[1].text, m_hdr[2].text = "Manager", "Employee", "Edit Date"
        for _, row in missing.iterrows():
            r = m_table.add_row().cells
            r[0].text, r[1].text, r[2].text = str(row["Manager"]), str(row["Employee"]), str(row["Edit_Date"])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- STREAMLIT UI ---

st.set_page_config(page_title="Toast Audit Helper", page_icon="📋", layout="wide")

st.title("📋 Toast POS Time Edit Audit Tool")
st.write("Upload Toast POS Time Audit CSVs and Scanned PDF Acknowledgment Forms.")

col1, col2 = st.columns(2)
with col1:
    csv_files = st.file_uploader("Upload Toast Audit CSV(s)", type=["csv"], accept_multiple_files=True)
with col2:
    pdf_files = st.file_uploader("Upload Signed PDF Forms", type=["pdf"], accept_multiple_files=True)

if st.button("Process & Generate Audit", type="primary"):
    if not csv_files or not pdf_files:
        st.error("Please upload both CSV and PDF files.")
    else:
        with st.spinner("Running OCR on scanned PDFs and matching edits..."):
            all_ocr_pages = []
            for pdf_file in pdf_files:
                pages = extract_text_with_ocr(pdf_file.read())
                all_ocr_pages.extend(pages)

            csv_df = process_toast_csv(csv_files)
            summary_df, details_df = analyze_edits(csv_df, all_ocr_pages)

            st.success("Audit complete!")

            total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
            total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
            overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manager Edits", total_edits)
            m2.metric("Signed Forms Found", total_forms)
            m3.metric("Overall Compliance", f"{overall_pct}%")

            st.subheader("Manager Summary")
            st.dataframe(summary_df, use_container_width=True)

            st.subheader("Missing Forms Detail")
            missing_df = details_df[~details_df["Has_Signed_Form"]]
            st.dataframe(missing_df, use_container_width=True)

            docx_buf = create_word_docx(summary_df, details_df)
            st.download_button(
                label="📥 Download Audit Report (.docx)",
                data=docx_buf,
                file_name="Time_Edit_Audit_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
