import io
import re
import difflib
import pandas as pd
import streamlit as st
from docx import Document
from pdf2image import convert_from_bytes
import pytesseract


def extract_text_with_ocr(pdf_bytes: bytes) -> list[str]:
    """Converts scanned PDF pages into images and runs OCR page-by-page."""
    if not pdf_bytes:
        return []
    images = convert_from_bytes(pdf_bytes)
    page_texts = []
    for img in images:
        text = pytesseract.image_to_string(img)
        page_texts.append(text)
    return page_texts


def process_toast_csv(csv_files) -> tuple[pd.DataFrame, str, str]:
    """Combines Toast CSVs, extracts location name, and assigns pay periods."""
    dfs = []
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        dfs.append(df)

    df_all = pd.concat(dfs, ignore_index=True)
    df_all.columns = df_all.columns.str.strip()

    # Extract Location Name
    loc_col = "Location" if "Location" in df_all.columns else next((c for c in df_all.columns if "location" in c.lower()), None)
    location_name = "Toast_Audit"
    if loc_col and not df_all[loc_col].dropna().empty:
        raw_loc = str(df_all[loc_col].dropna().iloc[0]).strip()
        location_name = re.sub(r"[^\w\s-]", "", raw_loc).replace(" ", "_")

    # Extract Pay Periods based on 'In Date'
    in_date_col = "In Date" if "In Date" in df_all.columns else next((c for c in df_all.columns if "date" in c.lower()), None)
    
    if in_date_col:
        df_all["In_DT"] = pd.to_datetime(df_all[in_date_col], errors="coerce")
        min_date_str = df_all["In_DT"].min().strftime("%m.%d") if not df_all["In_DT"].dropna().empty else "Start"
        max_date_str = df_all["In_DT"].max().strftime("%m.%d") if not df_all["In_DT"].dropna().empty else "End"
        overall_range_str = f"{min_date_str}_to_{max_date_str}"

        # Assign 14-day Pay Period labels
        min_dt = df_all["In_DT"].min()
        if pd.notna(min_dt):
            df_all["Pay_Period_Days"] = (df_all["In_DT"] - min_dt).dt.days // 14
            
            def make_period_label(group):
                p_min = group["In_DT"].min().strftime("%m/%d/%Y")
                p_max = group["In_DT"].max().strftime("%m/%d/%Y")
                return f"Pay Period ({p_min} - {p_max})"
            
            period_map = df_all.groupby("Pay_Period_Days").apply(make_period_label).to_dict()
            df_all["Pay_Period"] = df_all["Pay_Period_Days"].map(period_map)
        else:
            df_all["Pay_Period"] = "All Edits"
    else:
        df_all["Pay_Period"] = "All Edits"
        overall_range_str = "Report"

    return df_all, location_name, overall_range_str


def get_date_variants(date_val) -> list[str]:
    """Generates short date variants for matching."""
    if pd.isna(date_val):
        return []
    try:
        dt = pd.to_datetime(date_val)
        m, d, y = dt.month, dt.day, str(dt.year)[-2:]
        return [f"{m}/{d}", f"{m:02d}/{d:02d}", f"{m}/{d}/{y}", f"{m:02d}/{d:02d}/{y}"]
    except Exception:
        return [str(date_val).strip()]


def get_name_tokens(name_str: str) -> list[str]:
    """Extracts searchable name parts."""
    if pd.isna(name_str):
        return []
    parts = re.split(r"[\s,]+", str(name_str).strip())
    ignore = {"pm", "bar", "boh", "foh", "take", "out", "lunch", "ghost", "drawer", "id"}
    return [p for p in parts if len(p) >= 3 and p.lower() not in ignore and not p.isdigit()]


def fuzzy_match_tokens(csv_tokens: list[str], ocr_text: str, threshold: float = 0.65) -> bool:
    """Handles handwritten OCR typos."""
    words = [w.strip(".,;:()") for w in re.split(r"\s+", ocr_text) if len(w) >= 3]
    for ct in csv_tokens:
        for ow in words:
            if difflib.SequenceMatcher(None, ct.lower(), ow.lower()).ratio() >= threshold:
                return True
    return False


def extract_time_from_datetime(datetime_val) -> str:
    """Extracts formatted time string."""
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
    """Audit engine matching edits per pay period."""
    emp_col = "Employee" if "Employee" in df.columns else next((c for c in df.columns if "employee" in c.lower() and "id" not in c.lower()), df.columns[0])
    mgr_col = "Manager" if "Manager" in df.columns else next((c for c in df.columns if "manager" in c.lower() or "edited" in c.lower()), df.columns[1])
    change_col = "Change" if "Change" in df.columns else next((c for c in df.columns if "change" in c.lower()), None)
    in_date_col = "In Date" if "In Date" in df.columns else next((c for c in df.columns if "date" in c.lower()), None)
    out_date_col = "Out Date" if "Out Date" in df.columns else None
    job_title_col = "Job Title" if "Job Title" in df.columns else None
    time_edit_col = "Time" if "Time" in df.columns else None

    # Filter out blank managers and CREATE entries
    filtered = df.dropna(subset=[mgr_col]).copy()
    if change_col and change_col in filtered.columns:
        filtered = filtered[filtered[change_col] != "CREATE"]

    # Exclude system/ghost employees
    has_digit_pattern = r"\d"
    system_terms = ["system", "ghost", "house", "auto", "toast", "drawer"]
    system_pattern = "|".join(system_terms)
    
    filtered = filtered[
        ~filtered[emp_col].astype(str).str.contains(has_digit_pattern, regex=True, na=False) &
        ~filtered[emp_col].astype(str).str.lower().str.contains(system_pattern, na=False) &
        ~filtered[mgr_col].astype(str).str.lower().str.contains(system_pattern, na=False)
    ].copy()

    filtered["Has_Signed_Form"] = False
    used_csv_indices = set()

    # Process PDF pages first
    for page_idx, page_text in enumerate(ocr_pages):
        if "EDITED PUNCH REQUEST" not in page_text.upper() and "SOLICITUD" not in page_text.upper():
            continue

        best_match_idx = None
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
            "Pay_Period": row.get("Pay_Period", "All Edits"),
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
        summary_df = pd.DataFrame(columns=["Pay_Period", "Manager", "Total_Edits", "Forms_Present", "Forms_Missing", "Compliance_Pct"])
        return summary_df, details_df

    summary_df = details_df.groupby(["Pay_Period", "Manager"]).agg(
        Total_Edits=("Has_Signed_Form", "count"),
        Forms_Present=("Has_Signed_Form", "sum")
    ).reset_index()

    summary_df["Forms_Missing"] = summary_df["Total_Edits"] - summary_df["Forms_Present"]
    summary_df["Compliance_Pct"] = ((summary_df["Forms_Present"] / summary_df["Total_Edits"]) * 100).round(1)

    return summary_df, details_df


def create_word_docx(summary_df: pd.DataFrame, details_df: pd.DataFrame, location_name: str) -> io.BytesIO:
    """Generates Word document broken down by Pay Period."""
    doc = Document()
    doc.add_heading(f"{location_name.replace('_', ' ')} - Time Edit Audit Report", level=1)

    total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
    total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
    overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0.0

    doc.add_heading("Overall Executive Summary", level=2)
    p = doc.add_paragraph()
    p.add_run("Location: ").bold = True
    p.add_run(f"{location_name.replace('_', ' ')}\n")
    p.add_run("Total Manager Edits: ").bold = True
    p.add_run(f"{total_edits}\n")
    p.add_run("Signed Forms Found: ").bold = True
    p.add_run(f"{total_forms}\n")
    p.add_run("Overall Compliance Rate: ").bold = True
    p.add_run(f"{overall_pct}%")

    # Group by Pay Period
    pay_periods = details_df["Pay_Period"].unique() if not details_df.empty else []

    for period in pay_periods:
        doc.add_heading(f"📅 {period}", level=2)

        p_summary = summary_df[summary_df["Pay_Period"] == period]
        doc.add_heading("Manager Summary", level=3)
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text, hdr[4].text = (
            "Manager", "Total Edits", "Forms Present", "Forms Missing", "Compliance %"
        )

        for _, row in p_summary.iterrows():
            r = table.add_row().cells
            r[0].text, r[1].text, r[2].text, r[3].text, r[4].text = (
                str(row["Manager"]), str(row["Total_Edits"]), str(row["Forms_Present"]),
                str(row["Forms_Missing"]), f"{row['Compliance_Pct']}%"
            )

        p_details = details_df[details_df["Pay_Period"] == period]
        doc.add_heading("Audit Details", level=3)
        d_table = doc.add_table(rows=1, cols=7)
        d_table.style = "Table Grid"
        d_hdr = d_table.rows[0].cells
        d_hdr[0].text, d_hdr[1].text, d_hdr[2].text, d_hdr[3].text, d_hdr[4].text, d_hdr[5].text, d_hdr[6].text = (
            "Manager", "Employee", "Shift Date", "In Time", "Out Time", "Break Edit?", "Form Signed?"
        )

        for _, row in p_details.iterrows():
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
        with st.spinner("Processing OCR and categorizing by pay period..."):
            all_ocr_pages = []
            for pdf_file in pdf_files:
                pdf_bytes = pdf_file.getvalue()
                pages = extract_text_with_ocr(pdf_bytes)
                all_ocr_pages.extend(pages)

            csv_df, location_name, overall_range_str = process_toast_csv(csv_files)
            summary_df, details_df = analyze_edits(csv_df, all_ocr_pages)

            st.success("Audit complete!")

            total_edits = summary_df["Total_Edits"].sum() if not summary_df.empty else 0
            total_forms = summary_df["Forms_Present"].sum() if not summary_df.empty else 0
            overall_pct = round((total_forms / total_edits) * 100, 1) if total_edits > 0 else 0

            # Display metrics
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Location", location_name.replace("_", " "))
            m2.metric("Total Manager Edits", total_edits)
            m3.metric("Signed Forms Found", total_forms)
            m4.metric("Overall Compliance", f"{overall_pct}%")

            # Display Tabbed Pay Period Results
            pay_periods = details_df["Pay_Period"].unique() if not details_df.empty else ["Summary"]
            tabs = st.tabs(pay_periods)

            for tab, period in zip(tabs, pay_periods):
                with tab:
                    p_summary = summary_df[summary_df["Pay_Period"] == period].drop(columns=["Pay_Period"])
                    p_details = details_df[details_df["Pay_Period"] == period].drop(columns=["Pay_Period"])

                    st.subheader(f"Manager Summary ({period})")
                    st.dataframe(p_summary, use_container_width=True)

                    st.subheader(f"Audit Detail Report ({period})")
                    st.dataframe(p_details, use_container_width=True)

            file_name_str = f"{location_name}_Time_Edit_Audit_Report_{overall_range_str}.docx"
            docx_buf = create_word_docx(summary_df, details_df, location_name)

            st.download_button(
                label=f"📥 Download Audit Report ({file_name_str})",
                data=docx_buf,
                file_name=file_name_str,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
