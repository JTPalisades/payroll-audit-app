import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document


def extract_pdf_pages(pdf_files) -> list[str]:
    """Extracts text page-by-page from uploaded PDF files."""
    pages_text = []
    for pdf_file in pdf_files:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages_text.append(text)
    return pages_text


def process_toast_csv(csv_files) -> pd.DataFrame:
    """Combines and cleans Toast POS time entry CSV files."""
    dfs = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)
    df_all.columns = df_all.columns.str.strip().str.lower()
    return df_all


def generate_date_variants(date_val) -> list[str]:
    """Generates common string variations of a date to match against PDF text."""
    if pd.isna(date_val):
        return []

    try:
        parsed_date = pd.to_datetime(date_val)
    except Exception:
        return [str(date_val).strip()]

    # Formats: 09/15/2026, 9/15/2026, 09/15/26, Sep 15, 2026, 2026-09-15
    variants = {
        parsed_date.strftime("%m/%d/%Y"),
        f"{parsed_date.month}/{parsed_date.day}/{parsed_date.year}",
        parsed_date.strftime("%m/%d/%y"),
        parsed_date.strftime("%b %d, %Y"),
        parsed_date.strftime("%B %d, %Y"),
        parsed_date.strftime("%Y-%m-%d"),
    }
    return list(variants)


def analyze_edits_strict(df: pd.DataFrame, pdf_pages: list[str]) -> pd.DataFrame:
    """Analyzes manager edits and checks for employee name AND edit date on the same PDF page."""
    # Column auto-detection based on lowercased names
    mgr_col = next((c for c in df.columns if ("edit" in c and "by" in c) or "manager" in c), df.columns[0])
    emp_col = next((c for c in df.columns if "employee" in c or "name" in c), df.columns[1])
    date_col = next((c for c in df.columns if "date" in c or "time" in c or "shift" in c), None)

    # Filter out system entries
    system_identifiers = ["system", "auto", "ghost", "house", "toast"]
    filtered_df = df[
        ~df[emp_col].astype(str).str.lower().isin(system_identifiers)
        & ~df[mgr_col].astype(str).str.lower().isin(system_identifiers)
    ].dropna(subset=[mgr_col]).copy()

    results = []

    for idx, row in filtered_df.iterrows():
        manager = str(row[mgr_col]).strip()
        employee = str(row[emp_col]).strip()
        date_raw = row[date_col] if date_col else None

        # Extract name tokens
        emp_tokens = [part for part in re.split(r"\s+|,", employee) if len(part) > 1]
        date_variants = generate_date_variants(date_raw) if date_raw is not None else []

        has_valid_form = False

        # Scan page by page
        for page_text in pdf_pages:
            # Check 1: Does the page contain the employee's name?
            name_match = all(
                re.search(r"\b" + re.escape(token) + r"\b", page_text, re.IGNORECASE)
                for token in emp_tokens
            ) if emp_tokens else False

            # Check 2: Does the page contain any variant of the edit date?
            date_match = any(
                re.search(r"\b" + re.escape(d_var) + r"\b", page_text, re.IGNORECASE)
                for d_var in date_variants
            ) if date_variants else True

            if name_match and date_match:
                has_valid_form = True
                break

        results.append({
            "Manager": manager,
            "Employee": employee,
            "Edit_Date": str(date_raw) if date_raw is not None else "N/A",
            "Has_Signed_Form": has_valid_form
        })

    res_df = pd.DataFrame(results)

    if res_df.empty:
        return pd.DataFrame(columns=["Manager", "Total_Edits", "Forms_Present", "Forms_Missing", "Compliance_Pct"])

    # Summarize by Manager
    summary = (
        res_df.groupby("Manager")
        .agg(
            Total_Edits=("Has_Signed_Form", "count"),
            Forms_Present=("Has_Signed_Form", "sum")
        )
        .reset_index()
    )

    summary["Forms_Missing"] = summary["Total_Edits"] - summary["Forms_Present"]
    summary["Compliance_Pct"] = (
        (summary["Forms_Present"] / summary["Total_Edits"]) * 100
    ).round(1)

    return summary


def create_word_docx(summary_df: pd.DataFrame) -> io.BytesIO:
    """Generates a formatted Word Document summarizing the audit."""
    doc = Document()
    doc.add_heading("Time Entry Edit Audit Summary", level=1)

    doc.add_paragraph(
        "This report outlines manager time edits and verifies corresponding signed acknowledgment forms."
    )

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Manager"
    hdr_cells[1].text = "Total Edits"
    hdr_cells[2].text = "Forms Present"
    hdr_cells[3].text = "Forms Missing"
    hdr_cells[4].text = "Compliance %"

    for _, row in summary_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row["Manager"])
        row_cells[1].text = str(row["Total_Edits"])
        row_cells[2].text = str(row["Forms_Present"])
        row_cells[3].text = str(row["Forms_Missing"])
        row_cells[4].text = f"{row['Compliance_Pct']}%"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- STREAMLIT USER INTERFACE ---

st.set_page_config(page_title="Toast Audit Helper", page_icon="📋", layout="wide")

st.title("📋 Toast POS Time Edit Audit Tool")
st.write(
    "Upload your Toast CSV audit files and signed PDF acknowledgment sheets to compare manager edits against signed forms."
)

col1, col2 = st.columns(2)

with col1:
    csv_files = st.file_uploader(
        "Upload Toast Time Entry Audit CSV(s)",
        type=["csv"],
        accept_multiple_files=True,
    )

with col2:
    pdf_files = st.file_uploader(
        "Upload Signed PDF Acknowledgment Form(s)",
        type=["pdf"],
        accept_multiple_files=True,
    )

if st.button("Process & Generate Audit", type="primary"):
    if not csv_files or not pdf_files:
        st.error("Please upload at least one CSV file and one PDF file.")
    else:
        with st.spinner("Analyzing files for name and date matches..."):
            # 1. Parse PDF pages
            pdf_pages = extract_pdf_pages(pdf_files)

            # 2. Process CSV
            csv_df = process_toast_csv(csv_files)

            # 3. Analyze Edits using strict date + name logic
            summary_df = analyze_edits_strict(csv_df, pdf_pages)

            st.success("Audit complete!")

            # 4. Display Metrics & Summary Table
            st.subheader("Audit Overview")

            total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
            total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
            overall_pct = (
                round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0
            )

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manager Edits", total_edits)
            m2.metric("Signed Forms Found", total_forms)
            m3.metric("Overall Compliance Rate", f"{overall_pct}%")

            st.dataframe(summary_df, use_container_width=True)

            # 5. Generate Word Document Download Button
            if not summary_df.empty:
                docx_buffer = create_word_docx(summary_df)

                st.download_button(
                    label="📥 Download Audit Report (.docx)",
                    data=docx_buffer,
                    file_name="Time_Edit_Audit_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
