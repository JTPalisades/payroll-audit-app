import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document


def extract_pdf_text(pdf_files) -> str:
    """Extracts all raw text from uploaded PDF files."""
    combined_text = ""
    for pdf_file in pdf_files:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    combined_text += text + "\n"
    return combined_text


def process_toast_csv(csv_files) -> pd.DataFrame:
    """Combines and cleans Toast POS time entry CSV files."""
    dfs = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)

    # Standardize column headers (lowercased/stripped)
    df_all.columns = df_all.columns.str.strip().str.lower()

    return df_all


def analyze_edits(df: pd.DataFrame, pdf_text: str) -> pd.DataFrame:
    """Analyzes manager edits and checks for corresponding employee signatures in PDF text."""
    # Common Toast column names for edits/managers/employees
    # Adjust these string keys if your Toast export uses different column headers
    mgr_col = next((c for c in df.columns if "edit" in c and "by" in c or "manager" in c), None)
    emp_col = next((c for c in df.columns if "employee" in c or "name" in c), None)

    if not mgr_col or not emp_col:
        # Fallback to direct selection if columns aren't auto-detected
        mgr_col = df.columns[0]
        emp_col = df.columns[1]

    # Filter out system/ghost employees or non-edited rows
    filtered_df = df.copy()

    # Filter out common system entries
    system_identifiers = ["system", "auto", "ghost", "house", "toast"]
    filtered_df = filtered_df[
        ~filtered_df[emp_col].astype(str).str.lower().isin(system_identifiers)
    ]
    filtered_df = filtered_df[
        ~filtered_df[mgr_col].astype(str).str.lower().isin(system_identifiers)
    ]

    # Drop rows where no manager edit took place (if applicable)
    filtered_df = filtered_df.dropna(subset=[mgr_col])

    results = []

    for idx, row in filtered_df.iterrows():
        manager = str(row[mgr_col]).strip()
        employee = str(row[emp_col]).strip()

        # Check if employee's name appears in the PDF acknowledgment text
        # Clean employee name for robust string matching
        emp_first_last = [
            part for part in re.split(r"\s+|,", employee) if len(part) > 1
        ]
        has_form = False

        if emp_first_last:
            # Matches if both first and last name appear in proximity or general document
            matches = [
                re.search(r"\b" + re.escape(part) + r"\b", pdf_text, re.IGNORECASE)
                for part in emp_first_last
            ]
            has_form = all(m is not None for m in matches)

        results.append(
            {
                "Manager": manager,
                "Employee": employee,
                "Has_Signed_Form": has_form,
            }
        )

    res_df = pd.DataFrame(results)

    # Group by Manager
    summary = (
        res_df.groupby("Manager")
        .agg(
            Total_Edits=("Has_Signed_Form", "count"),
            Forms_Present=("Has_Signed_Form", "sum"),
        )
        .reset_index()
    )

    summary["Forms_Missing"] = summary["Total_Edits"] - summary["Forms_Present"]
    summary["Compliance_Pct"] = (
        summary["Forms_Present"] / summary["Total_Edits"] * 100
    ).round(1)

    return summary


def create_word_docx(summary_df: pd.DataFrame) -> io.BytesIO:
    """Generates a formatted Word Document summarizing the audit."""
    doc = Document()
    doc.add_heading("Time Entry Edit Audit Summary", level=1)

    doc.add_paragraph(
        "This report outlines manager time edits and verifies corresponding signed acknowledgment forms."
    )

    # Add Table
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

    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- STREAMLIT UI ---

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
        with st.spinner("Analyzing files..."):
            # 1. Parse PDF
            pdf_text = extract_pdf_text(pdf_files)

            # 2. Process CSV
            csv_df = process_toast_csv(csv_files)

            # 3. Analyze Edits
            summary_df = analyze_edits(csv_df, pdf_text)

            st.success("Audit complete!")

            # 4. Display Metrics & Summary Table
            st.subheader("Audit Overview")

            total_edits = summary_df["Total_Edits"].sum()
            total_forms = summary_df["Forms_Present"].sum()
            overall_pct = (
                round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0
            )

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manager Edits", total_edits)
            m2.metric("Signed Forms Found", total_forms)
            m3.metric("Overall Compliance Rate", f"{overall_pct}%")

            st.dataframe(summary_df, use_container_width=True)

            # 5. Generate Word Document Download Button
            docx_buffer = create_word_docx(summary_df)

            st.download_button(
                label="📥 Download Audit Report (.docx)",
                data=docx_buffer,
                file_name="Time_Edit_Audit_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
