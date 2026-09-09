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


def process_toast_csv(csv_files) -> tuple[pd.DataFrame, str]:
    """Combines Toast POS CSV files and extracts the store location name."""
    dfs = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)
    df_all.columns = df_all.columns.str.strip()

    # Extract location name from CSV
    loc_col = next((c for c in df_all.columns if "location" in c.lower() and "code" not in c.lower()), None)
    location_name = "Location"
    
    if loc_col and not df_all[loc_col].dropna().empty:
        raw_loc = str(df_all[loc_col].dropna().iloc[0]).strip()
        # Clean location name for safe filename usage
        location_name = re.sub(r'[^\w\s-]', '', raw_loc).strip().replace(" ", "_")

    return df_all, location_name


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


def extract_time_from_datetime(datetime_val) -> str:
    """Extracts just the time string (e.g. '10:00 AM') from a full datetime string."""
    if pd.isna(datetime_val):
        return "N/A"
    
    val_str = str(datetime_val).strip()
    try:
        dt = pd.to_datetime(val_str)
        return dt.strftime("%I:%M %p").lstrip("0")
    except Exception:
        match = re.search(r"\d{1,2}:\d{2}(?:\s*[AP]M)?", val_str, re.IGNORECASE)
        return match.group(0) if match else val_str


def analyze_edits(df: pd.DataFrame, ocr_pages: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Page-first audit engine with refined employee, break, and timestamp filtering."""
    
    emp_col = "Employee" if "Employee" in df.columns else next((c for c in df.columns if "employee" in c.lower() and "id" not in c.lower()), df.columns[0])
    mgr_col = "Manager" if "Manager" in df.columns else next((c for c in df.columns if "manager" in c.lower() or "edited" in c.lower()), df.columns[1])
    change_col = "Change" if "Change" in df.columns else next((c for c in df.columns if "change" in c.lower()), None)
    in_date_col = "In Date" if "In Date" in df.columns else next((c for c in df.columns if "date" in c.lower()), None)
    out_date_col = "Out Date" if "Out Date" in df.columns else None
    job_title_col = "Job Title" if "Job Title" in df.columns else None
    time_edit_col = "Time" if "Time" in df.columns else None

    # 1. Ignore rows where Manager is blank/NaN
    filtered = df.dropna(subset=[mgr_col]).copy()

    # 2. Filter out CREATE entries
    if change_col and change_col in filtered.columns:
        filtered = filtered[filtered[change_col] != "CREATE"]

    # 3. Ignore ghost employees (names containing digits or standard system words)
    has_digit_pattern = r"\d"
    system_terms = ["system", "ghost", "house", "auto", "toast", "drawer"]
    system_pattern = "|".join(system_terms)
    
    filtered = filtered[
        ~filtered[emp_col].astype(str).str.contains(has_digit_pattern, regex=True, na=False) &
        ~filtered[emp_col].astype(str).str.lower().str.contains(system_pattern, na=False) &
        ~filtered[mgr_col].astype(str).str.lower().str.contains(system_pattern, na=False)
    ].copy()

    # Flag column for form match
    filtered["Has_Signed_Form"] = False
    used_csv_indices = set()

    # PROCESS PDF PAGES FIRST
    for page_idx, page_text in enumerate(ocr_pages):
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

    # Build detailed records
    details_list = []
    for _, row in filtered.iterrows():
        job_title = str(row[job_title_col]) if job_title_col and not pd.isna(row[job_title_col]) else ""
        is_break = "Yes" if "break" in job_title.lower() else "No"

        in_time = extract_time_from_datetime(row[in_date_col]) if in_date_col else "N/A"
        out_time = extract_time_from_datetime(row[out_date_col]) if out_date_col else "N/A"

        shift_date = "N/A"
        if in_date_col and not pd.isna(row[in_date_col]):
            try:
                shift_date = pd.to_datetime(row[in_date_col]).strftime("%m/%d/%Y")
            except Exception:
                shift_date = str(row[in_date_col]).split()[0]

        details_list.append({
            "Manager": str(row[mgr_col]).strip(),
            "Employee": str(row[emp_col]).strip(),
            "Shift_Date": shift_date,
            "In_Time": in_time,
            "Out_Time": out_time,
            "Is_Break_Edit": is_break,
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


def create_word_docx(summary_df: pd.DataFrame, details_df: pd.DataFrame, location_name: str) -> io.BytesIO:
    """Generates downloadable Word audit summary report with location name in title."""
    doc = Document()
    doc.add_heading(f"Toast POS Time Edit Audit Report - {location_name.replace('_', ' ')}", level=1)

    total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
    total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
    overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0.0

    # Executive Overview
    doc.add_heading("Audit Executive Summary", level=2)
    p = doc.add_paragraph()
    p.add_run("Location: ").bold = True
    p.add_run(f"{location_name.replace('_', ' ')}\n")
    p.add_run("Total Manager Edits: ").bold = True
    p.add_run(f"{total_edits}\n")
    p.add_run("Signed Forms Found: ").bold = True
    p.add_run(f"{total_forms}\n")
    p.add_run("Overall Compliance Rate: ").bold = True
    p.add_run(f"{overall_pct}%")

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

    doc.add_heading("Time Edit Audit Details", level=2)
    d_table = doc.add_table(rows=1, cols=7)
    d_table.style = "Table Grid"
    d_hdr = d_table.rows[0].cells
    d_hdr[0].text, d_hdr[1].text, d_hdr[2].text, d_hdr[3].text, d_hdr[4].text, d_hdr[5].text, d_hdr[6].text = (
        "Manager", "Employee", "Shift Date", "In Time", "Out Time", "Break Edit?", "Form Signed?"
    )

    for _, row in details_df.iterrows():
        r = d_table.add_row().cells
        r[0].text = str(row["Manager"])
        r[1].text = str(row["Employee"])
        r[2].text = str(row["Shift_Date"])
        r[3].text = str(row["In_Time"])
        r[4].text = str(row["Out_Time"])
        r[5].text = str(row["Is_Break_Edit"])
        r[6].text = "Yes" if row["Has_Signed_Form"] else "No"

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

            csv_df, location_name = process_toast_csv(csv_files)
            summary_df, details_df = analyze_edits(csv_df, all_ocr_pages)

            st.success("Audit complete!")

            total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
            total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
            overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0

            # Metrics
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manager Edits", total_edits)
            m2.metric("Signed Forms Found", total_forms)
            m3.metric("Overall Compliance", f"{overall_pct}%")

            st.subheader(f"Manager Summary ({location_name.replace('_', ' ')})")
            st.dataframe(summary_df, use_container_width=True)

            st.subheader("Audit Detail Report")
            st.dataframe(details_df, use_container_width=True)

            docx_buf = create_word_docx(summary_df, details_df, location_name)
            
            # Dynamic Filename using extracted Location Name
            report_filename = f"{location_name}_Time_Edit_Audit_Report.docx"

            st.download_button(
                label=f"📥 Download {location_name} Audit Report (.docx)",
                data=docx_buf,
                file_name=report_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
