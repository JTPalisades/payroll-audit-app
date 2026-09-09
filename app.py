import io
import re
import difflib
import pandas as pd
import streamlit as st
from docx import Document
from pdf2image import convert_from_bytes
import pytesseract


def extract_text_with_ocr_split(pdf_bytes: bytes) -> list[tuple[int, str]]:
    """
    Converts PDF pages to images and splits each page into TOP and BOTTOM halves.
    Returns tuples of (page_index, section_text) to bind sections to physical pages.
    """
    if not pdf_bytes:
        return []
    
    images = convert_from_bytes(pdf_bytes)
    page_sections = []

    for page_idx, img in enumerate(images):
        width, height = img.size
        
        # Crop Top and Bottom halves of page
        top_half = img.crop((0, 0, width, int(height * 0.52)))
        bottom_half = img.crop((0, int(height * 0.48), width, height))

        txt_top = pytesseract.image_to_string(top_half)
        txt_bottom = pytesseract.image_to_string(bottom_half)

        if txt_top and len(txt_top.strip()) > 20:
            page_sections.append((page_idx, txt_top))
        if txt_bottom and len(txt_bottom.strip()) > 20:
            page_sections.append((page_idx, txt_bottom))

    return page_sections


def process_toast_csv(csv_files) -> tuple[pd.DataFrame, str, str, str]:
    """Combines Toast POS CSV files, extracts store location name, and determines audit date range."""
    dfs = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)
    df_all.columns = df_all.columns.str.strip()

    # Extract location name
    loc_col = next((c for c in df_all.columns if "location" in c.lower() and "code" not in c.lower()), None)
    location_name = "Location"
    
    if loc_col and not df_all[loc_col].dropna().empty:
        raw_loc = str(df_all[loc_col].dropna().iloc[0]).strip()
        location_name = re.sub(r'[^\w\s-]', '', raw_loc).strip().replace(" ", "_")

    # Extract date range
    in_date_col = "In Date" if "In Date" in df_all.columns else next((c for c in df_all.columns if "date" in c.lower()), None)
    
    date_filename_str = "Audit_Report"
    date_display_str = "N/A"

    if in_date_col and not df_all[in_date_col].dropna().empty:
        dates = pd.to_datetime(df_all[in_date_col].dropna(), format="mixed", errors="coerce").dropna()
        if not dates.empty:
            min_dt, max_dt = dates.min(), dates.max()
            date_filename_str = f"{min_dt.strftime('%m.%d.%y')}_to_{max_dt.strftime('%m.%d.%y')}"
            date_display_str = f"{min_dt.strftime('%m/%d/%Y')} - {max_dt.strftime('%m/%d/%Y')}"

    return df_all, location_name, date_filename_str, date_display_str


def get_date_variants(date_val) -> list[str]:
    """Generates date variants supporting slashes, hyphens, and dot/period separators."""
    if pd.isna(date_val):
        return []
    
    try:
        dt = pd.to_datetime(date_val)
        m, d, y = dt.month, dt.day, str(dt.year)[-2:]
        return [
            f"{m}/{d}", f"{m:02d}/{d:02d}", f"{m}/{d}/{y}", f"{m:02d}/{d:02d}/{y}",
            f"{m}.{d}", f"{m:02d}.{d:02d}", f"{m}.{d}.{y}", f"{m:02d}.{d:02d}.{y}", f"{m}.{d}.{dt.year}",
            f"{m}-{d}", f"{m:02d}-{d:02d}", f"{m}-{d}-{y}", f"{m:02d}-{d:02d}-{y}"
        ]
    except Exception:
        s = str(date_val).strip()
        return [s, s.replace("/", "."), s.replace("/", "-")]


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
    """Extracts time string (e.g. '10:00 AM') from datetime string."""
    if pd.isna(datetime_val):
        return "N/A"
    
    val_str = str(datetime_val).strip()
    try:
        dt = pd.to_datetime(val_str)
        return dt.strftime("%I:%M %p").lstrip("0")
    except Exception:
        match = re.search(r"\d{1,2}:\d{2}(?:\s*[AP]M)?", val_str, re.IGNORECASE)
        return match.group(0) if match else val_str


def analyze_edits(df: pd.DataFrame, ocr_sections: list[tuple[int, str]]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Audit engine matching OCR sections against CSV edit entries with physical page deduplication."""
    
    emp_col = "Employee" if "Employee" in df.columns else next((c for c in df.columns if "employee" in c.lower() and "id" not in c.lower()), df.columns[0])
    mgr_col = "Manager" if "Manager" in df.columns else next((c for c in df.columns if "manager" in c.lower() or "edited" in c.lower() or "by" in c.lower()), df.columns[1])
    change_col = "Change" if "Change" in df.columns else next((c for c in df.columns if "change" in c.lower() or "action" in c.lower()), None)
    in_date_col = "In Date" if "In Date" in df.columns else next((c for c in df.columns if "date" in c.lower() or "time" in c.lower()), None)
    out_date_col = "Out Date" if "Out Date" in df.columns else None
    job_title_col = "Job Title" if "Job Title" in df.columns else next((c for c in df.columns if "job" in c.lower()), None)
    time_edit_col = "Time" if "Time" in df.columns else None

    # 1. Filter out missing managers
    filtered = df.dropna(subset=[mgr_col]).copy()

    # 2. Filter out CREATE actions
    if change_col and change_col in filtered.columns:
        filtered = filtered[filtered[change_col] != "CREATE"]

    # 3. Filter ghost employees
    has_digit_pattern = r"\d"
    system_terms = ["system", "ghost", "house", "auto", "toast", "drawer"]
    system_pattern = "|".join(system_terms)
    
    filtered = filtered[
        ~filtered[emp_col].astype(str).str.contains(has_digit_pattern, regex=True, na=False) &
        ~filtered[emp_col].astype(str).str.lower().str.contains(system_pattern, na=False) &
        ~filtered[mgr_col].astype(str).str.lower().str.contains(system_pattern, na=False)
    ].copy()

    # Deduplicate multiple edit rows for the same employee shift date
    filtered["Shift_Key"] = filtered[emp_col].astype(str) + "_" + pd.to_datetime(filtered[in_date_col], format="mixed", errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    
    matched_shift_keys = set()
    used_page_indices = set()

    # SECTION-FIRST MATCHING WITH PHYSICAL PAGE DEDUPLICATION
    for page_idx, sec_text in ocr_sections:
        if page_idx in used_page_indices:
            continue

        best_match_key = None

        # Pass 1: Strict Name AND Date Match
        for _, row in filtered.iterrows():
            shift_key = row["Shift_Key"]
            if shift_key in matched_shift_keys:
                continue

            employee = str(row[emp_col]).strip()
            tokens = get_name_tokens(employee)
            date_vars = get_date_variants(row[in_date_col]) + (get_date_variants(row[time_edit_col]) if time_edit_col else [])

            has_name = any(re.search(r"\b" + re.escape(t[:3]) + r"[a-z]*\b", sec_text, re.IGNORECASE) for t in tokens) or fuzzy_match_tokens(tokens, sec_text)
            has_date = any(d in sec_text for d in date_vars) if date_vars else True

            if has_name and has_date:
                best_match_key = shift_key
                break

        # Pass 2: Name-Only Fallback
        if best_match_key is None:
            for _, row in filtered.iterrows():
                shift_key = row["Shift_Key"]
                if shift_key in matched_shift_keys:
                    continue

                employee = str(row[emp_col]).strip()
                tokens = get_name_tokens(employee)
                has_name = any(re.search(r"\b" + re.escape(t[:3]) + r"[a-z]*\b", sec_text, re.IGNORECASE) for t in tokens) or fuzzy_match_tokens(tokens, sec_text)

                if has_name:
                    best_match_key = shift_key
                    break

        if best_match_key is not None:
            matched_shift_keys.add(best_match_key)
            used_page_indices.add(page_idx)

    # Apply match flags back to CSV rows
    filtered["Has_Signed_Form"] = filtered["Shift_Key"].isin(matched_shift_keys)

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

    # Summarize unique shift edits per manager
    summary_df = details_df.drop_duplicates(subset=["Employee", "Shift_Date", "Manager"]).groupby("Manager").agg(
        Total_Edits=("Has_Signed_Form", "count"),
        Forms_Present=("Has_Signed_Form", "sum")
    ).reset_index()

    summary_df["Forms_Missing"] = summary_df["Total_Edits"] - summary_df["Forms_Present"]
    summary_df["Compliance_Pct"] = ((summary_df["Forms_Present"] / summary_df["Total_Edits"]) * 100).round(1)

    return summary_df, details_df


def create_word_docx(summary_df: pd.DataFrame, details_df: pd.DataFrame, location_name: str, date_display_str: str) -> io.BytesIO:
    """Generates downloadable Word audit summary report with location and date range in headers."""
    doc = Document()
    clean_loc = location_name.replace("_", " ")
    
    doc.add_heading(f"Toast POS Time Edit Audit Report", level=1)
    doc.add_heading(f"{clean_loc} ({date_display_str})", level=2)

    total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
    total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
    overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0.0

    # Executive Overview
    doc.add_heading("Audit Executive Summary", level=3)
    p = doc.add_paragraph()
    p.add_run("Location: ").bold = True
    p.add_run(f"{clean_loc}\n")
    p.add_run("Date Range: ").bold = True
    p.add_run(f"{date_display_str}\n")
    p.add_run("Total Manager Shift Edits: ").bold = True
    p.add_run(f"{total_edits}\n")
    p.add_run("Signed Forms Found: ").bold = True
    p.add_run(f"{total_forms}\n")
    p.add_run("Overall Compliance Rate: ").bold = True
    p.add_run(f"{overall_pct}%")

    doc.add_heading("Manager Audit Summary", level=3)
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

    doc.add_heading("Time Edit Audit Details", level=3)
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
            all_ocr_sections = []
            for pdf_file in pdf_files:
                pdf_bytes = pdf_file.getvalue()
                sections = extract_text_with_ocr_split(pdf_bytes)
                all_ocr_sections.extend(sections)

            csv_df, location_name, date_filename_str, date_display_str = process_toast_csv(csv_files)
            summary_df, details_df = analyze_edits(csv_df, all_ocr_sections)

            st.success("Audit complete!")

            total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
            total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
            overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0

            # Display metrics in UI
            st.markdown(f"### Audit for **{location_name.replace('_', ' ')}** ({date_display_str})")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Manager Shift Edits", total_edits)
            m2.metric("Signed Forms Found", total_forms)
            m3.metric("Overall Compliance", f"{overall_pct}%")

            st.subheader("Manager Summary")
            st.dataframe(summary_df, use_container_width=True)

            st.subheader("Audit Detail Report")
            st.dataframe(details_df, use_container_width=True)

            docx_buf = create_word_docx(summary_df, details_df, location_name, date_display_str)
            report_filename = f"{location_name}_Audit_{date_filename_str}.docx"

            st.download_button(
                label=f"📥 Download {location_name} Audit Report (.docx)",
                data=docx_buf,
                file_name=report_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
