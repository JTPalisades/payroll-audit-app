import io
import re
import difflib
import pandas as pd
import streamlit as st
from docx import Document
from pdf2image import convert_from_bytes
import pytesseract


def extract_text_with_ocr(pdf_bytes: bytes) -> list[str]:
    """Converts scanned PDF pages into images and runs OCR to extract text page-by-page."""
    if not pdf_bytes:
        return []
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


def get_date_variants(date_val) -> list[str]:
    """Generates short date variants (e.g., '8/7', '08/07', '8/7/26') for matching."""
    if pd.isna(date_val):
        return []
    
    try:
        dt = pd.to_datetime(date_val)
        m, d, y = dt.month, dt.day, str(dt.year)[-2:]
        return [
            f"{m}/{d}",
            f"{m:02d}/{d:02d}",
            f"{m}/{d}/{y}",
            f"{m:02d}/{d:02d}/{y}"
        ]
    except Exception:
        return [str(date_val).strip()]


def get_name_tokens(name_str: str) -> list[str]:
    """Extracts major name parts, ignoring common filler terms."""
    if pd.isna(name_str):
        return []
    parts = re.split(r"[\s,]+", str(name_str).strip())
    ignore = {"pm", "bar", "boh", "foh", "take", "out", "lunch", "ghost", "drawer", "id"}
    return [p for p in parts if len(p) >= 3 and p.lower() not in ignore and not p.isdigit()]


def fuzzy_match_tokens(csv_tokens: list[str], ocr_text: str, threshold: float = 0.65) -> bool:
    """Handles handwritten OCR typos (e.g. matching 'Silva' with 'Siles')."""
    words = [w.strip(".,;:()") for w in re.split(r"\s+", ocr_text) if len(w) >= 3]
    for ct in csv_tokens:
        for ow in words:
            if difflib.SequenceMatcher(None, ct.lower(), ow.lower()).ratio() >= threshold:
                return True
    return False


def analyze_edits(df: pd.DataFrame, ocr_pages: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Page-first audit engine ensuring each PDF sheet matches its corresponding CSV entry."""
    
    emp_col = "Employee" if "Employee" in df.columns else next((c for c in df.columns if "employee" in c.lower() and "id" not in c.lower()), df.columns[0])
    mgr_col = "Manager" if "Manager" in df.columns else next((c for c in df.columns if "manager" in c.lower() or "edited" in c.lower()), df.columns[1])
    change_col = "Change" if "Change" in df.columns else next((c for c in df.columns if "change" in c.lower()), None)
    in_date_col = "In Date" if "In Date" in df.columns else next((c for c in df.columns if "date" in c.lower()), None)
    time_edit_col = "Time" if "Time" in df.columns else None

    filtered = df.copy()
    
    # Filter out CREATE entries (keep manager modifications/deletions)
    if change_col and change_col in filtered.columns:
        filtered = filtered[filtered[change_col] != "CREATE"]

    # Exclude system / ghost entries
    system_terms = ["system", "ghost", "house", "auto", "toast", "drawer"]
    pattern = "|".join(system_terms)
    
    filtered = filtered[
        ~filtered[emp_col].astype(str).str.lower().str.contains(pattern, na=False) &
        ~filtered[mgr_col].astype(str).str.lower().str.contains(pattern, na=False)
    ].dropna(subset=[mgr_col]).copy()

    # Flag column for form match
    filtered["Has_Signed_Form"] = False
    used_csv_indices = set()

    # PROCESS PDF PAGES FIRST
    for page_idx, page_text in enumerate(ocr_pages):
        # Ignore pages that are paycheck rosters or cover sheets lacking request text
        if "EDITED PUNCH REQUEST" not in page_text.upper() and "SOLICITUD" not in page_text.upper():
            continue

        best_match_idx = None

        # Pass 1: Strict Match (Name AND Date match)
        for idx, row in filtered.iterrows():
            if idx in used_csv_indices:
                continue

            employee = str(row[emp_col]).strip()
            tokens = get_name_tokens(employee)
            date_vars = get_date_variants(row[in_date_col]) + (get_date_variants(row[time_edit_col]) if time_edit_col else [])

            has_name = any(re.search(r"\b" + re.escape(t[:3]) + r"[a-z]*\b", page_text, re.IGNORECASE) for t in tokens) or fuzzy_match_tokens(tokens, page_text)
            has_date = any(d in page_text for d in date_vars) if date_vars else True

            if has_name and has_date:
                best_match_idx = idx
                break

        # Pass 2: Name-Only Fallback Match for this page
        if best_match_idx is None:
            for idx, row in filtered.iterrows():
                if idx in used_csv_indices:
                    continue

                employee = str(row[emp_col]).strip()
                tokens = get_name_tokens(employee)
                has_name = any(re.search(r"\b" + re.escape(t[:3]) + r"[a-z]*\b", page_text, re.IGNORECASE) for t in tokens) or fuzzy_match_tokens(tokens, page_text)

                if has_name:
                    best_match_idx = idx
                    break

        if best_match_idx is not None:
            filtered.loc[best_match_idx, "Has_Signed_Form"] = True
            used_csv_indices.add(best_match_idx)

    # Build summary and details DataFrames
    details_list = []
    for _, row in filtered.iterrows():
        details_list.append({
            "Manager": str(row[mgr_col]).strip(),
            "Employee": str(row[emp_col]).strip(),
            "Edit_Date": str(row[in_date_col]) if in_date_col else "N/A",
            "Has_Signed_Form": row["Has_Signed_Form"]
        })

    details_df = pd.DataFrame(details_list)

    if details_df.empty:
        summary_df = pd.DataFrame(columns=["Manager", "Total_Edits", "Forms_Present", "Forms_Missing", "Compliance_Pct"])
        return summary_df, details_df

    summary_df = details_df.groupby("Manager").agg(
        Total_Edits=("Has_Signed_Form", "count"),
        Forms_Present=("Has_Signed_Form", "sum")
    ).reset_index()

    summary_df["Forms_Missing"] = summary_df["Total_Edits"] - summary_df["Forms_Present"]
    summary_df["Compliance_Pct"] = ((summary_df["Forms_Present"] / summary_df["Total_Edits"]) * 100).round(1)

    return summary_df, details_df


def create_word_docx(summary_df: pd.DataFrame, details_df: pd.DataFrame) -> io.BytesIO:
    """Generates downloadable Word audit summary report."""
    doc = Document()
    doc.add_heading("Toast POS Time Edit Audit Report", level=1)

    doc.add_heading("Manager Audit Summary", level=2)
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

    doc.add_heading("Missing Forms Detail", level=2)
    missing = details_df[~details_df["Has_Signed_Form"]]
    if missing.empty:
        doc.add_paragraph("All manager edits have corresponding signed forms present.")
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
        with st.spinner("Running page-by-page OCR and matching edits..."):
            all_ocr_pages = []
            for pdf_file in pdf_files:
                pdf_bytes = pdf_file.getvalue()
                pages = extract_text_with_ocr(pdf_bytes)
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
