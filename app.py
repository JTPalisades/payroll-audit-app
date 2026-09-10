import io
import re
import difflib
import pandas as pd
import streamlit as st
from docx import Document
from pdf2image import convert_from_bytes
import pytesseract


def is_not_cover_or_report(ocr_text: str) -> bool:
    """
    Location-agnostic filter: Rejects paycheck rosters, cover sheets, 
    and embedded Toast POS audit log tables while accepting all edit form pages.
    """
    txt_upper = ocr_text.upper()
    
    # 1. Reject Payroll / Paycheck Sign-off Rosters
    roster_terms = [
        "PAYCHECK SIGNATURES", "PAYROLL SIGNATURES", "FOH PAYCHECK", "BOH PAYCHECK"
    ]
    if any(term in txt_upper for term in roster_terms):
        return False
        
    # 2. Reject Cover / Policy Confirmation Sheets
    cover_terms = [
        "CONFIRMATION AND ACKNOWLEDGEMENT", "PAY PERIOD ENDING:", "APPROPRIATE PAYMENT FOR HOURS WORKED"
    ]
    if any(term in txt_upper for term in cover_terms):
        return False
        
    # 3. Reject Printed Toast POS Audit Log Tables
    if ("TIMEIN" in txt_upper and "TIMEOUT" in txt_upper and "EMPLOYEE" in txt_upper) or ("ACTIONTYPE" in txt_upper and "FIELD" in txt_upper):
        return False
        
    return True


def extract_text_with_ocr_smart(pdf_bytes: bytes) -> tuple[list[tuple[int, str, str]], int]:
    """
    Extracts text sections from PDF pages.
    For dual-form sheets (separated by 'SOLICITUD DE HORAS EDITADAS'), splits into Top/Bottom.
    Otherwise, extracts full intact page text.
    """
    if not pdf_bytes:
        return [], 0
    
    try:
        images = convert_from_bytes(pdf_bytes)
    except Exception as e:
        st.error(f"PDF Image Rendering Error: {e}. Verify poppler-utils is installed in packages.txt.")
        return [], 0

    page_sections = []
    total_pages = len(images)

    for page_idx, img in enumerate(images):
        txt_full = pytesseract.image_to_string(img)
        if txt_full and len(txt_full.strip()) > 30 and is_not_cover_or_report(txt_full):
            if "SOLICITUD DE HORAS EDITADAS" in txt_full.upper():
                # Dual-form page split
                width, height = img.size
                top_half = img.crop((0, 0, width, int(height * 0.55)))
                bottom_half = img.crop((0, int(height * 0.45), width, height))

                txt_top = pytesseract.image_to_string(top_half)
                txt_bottom = pytesseract.image_to_string(bottom_half)

                if txt_top and len(txt_top.strip()) > 20:
                    page_sections.append((page_idx, "top", txt_top))
                if txt_bottom and len(txt_bottom.strip()) > 20:
                    page_sections.append((page_idx, "bottom", txt_bottom))
            else:
                # Single-form page
                page_sections.append((page_idx, "full", txt_full))

    return page_sections, total_pages


def process_toast_csv(csv_files) -> tuple[pd.DataFrame, str, str, str]:
    """Location-agnostic CSV cleaning and column auto-detection."""
    dfs = []
    for csv_file in csv_files:
        try:
            df = pd.read_csv(csv_file)
            dfs.append(df)
        except Exception as e:
            st.warning(f"Could not parse CSV file {csv_file.name}: {e}")

    if not dfs:
        return pd.DataFrame(), "Unknown_Location", "Audit_Report", "N/A"

    df_all = pd.concat(dfs, ignore_index=True)
    df_all.columns = df_all.columns.str.strip()

    # Dynamic Location Extraction
    loc_col = next((c for c in df_all.columns if "location" in c.lower() and "code" not in c.lower() and "id" not in c.lower()), None)
    location_name = "Location"
    
    if loc_col and not df_all[loc_col].dropna().empty:
        raw_loc = str(df_all[loc_col].dropna().iloc[0]).strip()
        location_name = re.sub(r'[^\w\s-]', '', raw_loc).strip().replace(" ", "_")
    elif csv_files:
        clean_fname = re.sub(r'[^\w\s-]', '', csv_files[0].name.split('.')[0]).strip().replace(" ", "_")
        location_name = clean_fname if clean_fname else "Location"

    # Date Range Extraction
    in_date_col = next((c for c in df_all.columns if "in date" in c.lower() or "clock in" in c.lower() or "shiftdate" in c.lower() or "date" in c.lower()), None)
    
    date_filename_str = "Audit_Report"
    date_display_str = "N/A"

    if in_date_col and not df_all[in_date_col].dropna().empty:
        dates = pd.to_datetime(df_all[in_date_col].dropna(), format="mixed", errors="coerce").dropna()
        if not dates.empty:
            min_dt, max_dt = dates.min(), dates.max()
            date_filename_str = f"{min_dt.strftime('%m.%d.%y')}_to_{max_dt.strftime('%m.%d.%y')}"
            date_display_str = f"{min_dt.strftime('%m/%d/%Y')} - {max_dt.strftime('%m/%d/%Y')}"

    return df_all, location_name, date_filename_str, date_display_str


def get_date_variants_universal(date_val) -> list[str]:
    """Generates slash, dot, hyphen, and compact digit date variants (e.g. 73026, 073026, 7.30.2026)."""
    if pd.isna(date_val):
        return []
    
    variants = set()
    try:
        dt = pd.to_datetime(date_val)
        m, d, y = dt.month, dt.day, str(dt.year)[-2:]
        full_y = dt.year
        
        # Standard Slashes
        variants.add(f"{m}/{d}")
        variants.add(f"{m:02d}/{d:02d}")
        variants.add(f"{m}/{d}/{y}")
        variants.add(f"{m:02d}/{d:02d}/{y}")
        
        # Dots/Periods
        variants.add(f"{m}.{d}")
        variants.add(f"{m:02d}.{d:02d}")
        variants.add(f"{m}.{d}.{y}")
        variants.add(f"{m:02d}.{d:02d}.{y}")
        variants.add(f"{m}.{d}.{full_y}")
        
        # Hyphens
        variants.add(f"{m}-{d}")
        variants.add(f"{m:02d}-{d:02d}")
        variants.add(f"{m}-{d}-{y}")
        variants.add(f"{m:02d}-{d:02d}-{y}")
        
        # Compact digits (73026, 073026, 72626)
        variants.add(f"{m}{d:02d}{y}")
        variants.add(f"{m:02d}{d:02d}{y}")
        variants.add(f"{m}{d}{y}")
        
    except Exception:
        s = str(date_val).strip()
        variants.add(s)
        variants.add(s.replace("/", "."))
        variants.add(s.replace("/", "-"))
        variants.add(s.replace("/", ""))

    return list(variants)


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


def analyze_edits(df: pd.DataFrame, ocr_sections: list[tuple[int, str, str]]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Universal audit engine matching PDF form sections against CSV shift edits."""
    if df.empty:
        summary_df = pd.DataFrame(columns=["Manager", "Total_Edits", "Forms_Present", "Forms_Missing", "Compliance_Pct"])
        return summary_df, pd.DataFrame()

    # Dynamic Column Auto-Detection
    emp_col = next((c for c in df.columns if "employee" in c.lower() and "id" not in c.lower() and "guid" not in c.lower() and "external" not in c.lower()), None)
    if not emp_col:
        emp_col = next((c for c in df.columns if "employee" in c.lower()), df.columns[0])

    mgr_col = next((c for c in df.columns if "manager" in c.lower() or "edited by" in c.lower() or "user" in c.lower()), None)
    if not mgr_col:
        mgr_col = df.columns[1]

    change_col = next((c for c in df.columns if "change" in c.lower() or "action" in c.lower() or "type" in c.lower()), None)
    in_date_col = next((c for c in df.columns if "in date" in c.lower() or "clock in" in c.lower() or "shiftdate" in c.lower() or "date" in c.lower()), df.columns[0])
    out_date_col = next((c for c in df.columns if "out date" in c.lower() or "clock out" in c.lower() or "timeout" in c.lower()), None)
    job_title_col = next((c for c in df.columns if "job" in c.lower() and "id" not in c.lower() and "guid" not in c.lower() and "code" not in c.lower()), None)
    time_edit_col = next((c for c in df.columns if "time" in c.lower() and "total" not in c.lower() and "break" not in c.lower() and "in" not in c.lower() and "out" not in c.lower()), None)

    # 1. Filter out missing managers
    filtered = df.dropna(subset=[mgr_col]).copy()

    # 2. Filter out CREATE actions
    if change_col and change_col in filtered.columns:
        filtered = filtered[filtered[change_col].astype(str).str.upper() != "CREATE"]

    # 3. Filter ghost employees
    has_digit_pattern = r"\d"
    system_terms = ["system", "ghost", "house", "auto", "toast", "drawer"]
    system_pattern = "|".join(system_terms)
    
    filtered = filtered[
        ~filtered[emp_col].astype(str).str.contains(has_digit_pattern, regex=True, na=False) &
        ~filtered[emp_col].astype(str).str.lower().str.contains(system_pattern, na=False) &
        ~filtered[mgr_col].astype(str).str.lower().str.contains(system_pattern, na=False)
    ].copy()

    filtered["Shift_Date_Clean"] = pd.to_datetime(filtered[in_date_col], format="mixed", errors="coerce").dt.strftime("%m/%d/%Y").fillna("")

    # Deduplicate CSV to Unique Manager Shift Edits
    shift_master = filtered.drop_duplicates(subset=[emp_col, "Shift_Date_Clean", mgr_col]).copy().reset_index()
    shift_master["Has_Signed_Form"] = False

    used_shift_indices = set()
    used_sections = set()

    # MATCH EACH PDF FORM SECTION TO A UNIQUE SHIFT EDIT
    for page_idx, sec_type, sec_text in ocr_sections:
        sec_key = (page_idx, sec_type)
        if sec_key in used_sections:
            continue

        best_match_idx = None

        # Pass 1: Strict Employee Name AND Shift Date Match
        for idx, row in shift_master.iterrows():
            if idx in used_shift_indices:
                continue

            employee = str(row[emp_col]).strip()
            tokens = get_name_tokens(employee)
            date_vars = get_date_variants_universal(row[in_date_col]) + (get_date_variants_universal(row[time_edit_col]) if time_edit_col and time_edit_col in row else [])

            has_emp = any(re.search(r"\b" + re.escape(t[:3]) + r"[a-z]*\b", sec_text, re.IGNORECASE) for t in tokens) or fuzzy_match_tokens(tokens, sec_text)
            has_date = any(d in sec_text for d in date_vars) if date_vars else True

            if has_emp and has_date:
                best_match_idx = idx
                break

        # Pass 2: Fallback Employee Name Match
        if best_match_idx is None:
            for idx, row in shift_master.iterrows():
                if idx in used_shift_indices:
                    continue

                employee = str(row[emp_col]).strip()
                tokens = get_name_tokens(employee)
                has_emp = any(re.search(r"\b" + re.escape(t[:3]) + r"[a-z]*\b", sec_text, re.IGNORECASE) for t in tokens) or fuzzy_match_tokens(tokens, sec_text)

                if has_emp:
                    best_match_idx = idx
                    break

        if best_match_idx is not None:
            shift_master.loc[best_match_idx, "Has_Signed_Form"] = True
            used_shift_indices.add(best_match_idx)
            used_sections.add(sec_key)

    # Map Has_Signed_Form back to all filtered rows
    signed_shifts_set = set(
        shift_master[shift_master["Has_Signed_Form"]][
            [emp_col, "Shift_Date_Clean", mgr_col]
        ].itertuples(index=False, name=None)
    )

    filtered["Has_Signed_Form"] = filtered[
        [emp_col, "Shift_Date_Clean", mgr_col]
    ].apply(lambda r: tuple(r) in signed_shifts_set, axis=1)

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

    # Manager Summary based on Unique Shift Edits
    unique_shifts_df = details_df.drop_duplicates(subset=["Employee", "Shift_Date", "Manager"])

    summary_df = unique_shifts_df.groupby("Manager").agg(
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
            try:
                # 1. OCR Processing
                all_ocr_sections = []
                total_pdf_pages = 0
                for pdf_file in pdf_files:
                    pdf_bytes = pdf_file.getvalue()
                    sections, page_count = extract_text_with_ocr_smart(pdf_bytes)
                    all_ocr_sections.extend(sections)
                    total_pdf_pages += page_count

                # 2. CSV Processing & Location Auto-Detection
                csv_df, location_name, date_filename_str, date_display_str = process_toast_csv(csv_files)

                # 3. Perform Matching & Audit Analysis
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

                # Diagnostic Status Expander
                with st.expander("🔍 Interactive Processing Diagnostics", expanded=False):
                    st.write(f"**Location Extracted:** {location_name.replace('_', ' ')}")
                    st.write(f"**Date Range Extracted:** {date_display_str}")
                    st.write(f"**PDF Pages Uploaded:** {total_pdf_pages}")
                    st.write(f"**Valid Form Sections Recognized:** {len(all_ocr_sections)}")
                    st.write(f"**Total CSV Records Loaded:** {len(csv_df)}")

                if total_forms == 0 and total_edits > 0:
                    st.warning("⚠️ 0 signed forms were matched. Please verify that the PDF contains valid form sheets and that scans are oriented right-side up.")

                st.subheader("Manager Summary")
                st.dataframe(summary_df, use_container_width=True)

                st.subheader("Audit Detail Report")
                st.dataframe(details_df, use_container_width=True)

                if not summary_df.empty:
                    docx_buf = create_word_docx(summary_df, details_df, location_name, date_display_str)
                    report_filename = f"{location_name}_Audit_{date_filename_str}.docx"

                    st.download_button(
                        label=f"📥 Download {location_name.replace('_', ' ')} Audit Report (.docx)",
                        data=docx_buf,
                        file_name=report_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as main_err:
                st.error(f"An unexpected error occurred during processing: {main_err}")
                st.info("Please verify that your uploaded files match the standard Toast POS export format.")
