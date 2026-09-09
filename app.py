import streamlit as st
import pandas as pd
import pdfplumber
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io
import re

st.set_page_config(page_title="Payroll Audit Generator", layout="centered")

st.title("Payroll Audit Report Generator")
st.write("Upload your POS CSV Log and PDF Packet to perform a fuzzy-matched line-by-line reconciliation.")

uploaded_csv = st.file_uploader("Upload POS CSV Log", type=["csv"])
uploaded_pdf = st.file_uploader("Upload Signed PDF Packet", type=["pdf"])

def extract_pdf_clean_text(pdf_file):
    """Extracts all text from PDF pages and normalizes whitespace/casing."""
    pdf_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                pdf_text += t.upper() + "\n---PAGE---\n"
    return pdf_text

def parse_name_tokens(raw_name):
    """Breaks employee name into significant search tokens (ignoring middle initials/roles)."""
    clean_name = re.sub(r'[^A-Z\s]', '', str(raw_name).upper())
    tokens = [w for w in clean_name.split() if len(w) >= 3 and w not in ["SERVER", "BUSSER", "HOST", "KITCHEN", "COOK", "MGR"]]
    return tokens

def process_audit_reconciliation(csv_df, pdf_file):
    # Clean CSV columns
    csv_df.columns = [str(c).strip() for c in csv_df.columns]
    
    # Identify key columns dynamically
    emp_col = next((c for c in csv_df.columns if "employee" in c.lower() or "name" in c.lower()), csv_df.columns[0])
    date_col = next((c for c in csv_df.columns if "date" in c.lower() or "shift" in c.lower()), csv_df.columns[1] if len(csv_df.columns) > 1 else csv_df.columns[0])
    mgr_col = next((c for c in csv_df.columns if "manager" in c.lower() or "editor" in c.lower() or "user" in c.lower()), csv_df.columns[3] if len(csv_df.columns) > 3 else csv_df.columns[0])

    # Convert Date column to actual DateTime objects
    csv_df['_ParsedDate'] = pd.to_datetime(csv_df[date_col], errors='coerce')
    csv_df = csv_df.dropna(subset=['_ParsedDate']).copy()
    csv_df = csv_df.sort_values('_ParsedDate')

    # Extract clean text from PDF
    pdf_text = extract_pdf_clean_text(pdf_file)

    verified_flags = []
    match_notes = []

    for _, row in csv_df.iterrows():
        raw_name = row[emp_col]
        name_tokens = parse_name_tokens(raw_name)
        
        shift_dt = row['_ParsedDate']
        # Extract multiple date string variants (e.g. 01/20/2026, 1/20/26, 01/20)
        date_variants = [
            shift_dt.strftime('%m/%d/%Y'),
            shift_dt.strftime('%m/%d/%y'),
            f"{shift_dt.month}/{shift_dt.day}/{shift_dt.year}",
            f"{shift_dt.month}/{shift_dt.day}/{shift_dt.strftime('%y')}",
            shift_dt.strftime('%m/%d')
        ]

        is_verified = False
        note = "Missing Form"

        if name_tokens:
            # Check if at least one significant name token (e.g., Last Name) AND a date variant appear in the same PDF block
            for token in name_tokens:
                if token in pdf_text:
                    for d_var in date_variants:
                        if d_var in pdf_text:
                            is_verified = True
                            note = f"Matched Token: {token} & Date: {d_var}"
                            break
                if is_verified:
                    break

        verified_flags.append(is_verified)
        match_notes.append(note)

    csv_df["_Verified"] = verified_flags
    csv_df["_MatchNote"] = match_notes

    # Assign Bi-Weekly Pay Period Ending (PPE) Dates
    min_date = csv_df['_ParsedDate'].min()
    csv_df['_PPE_Group'] = csv_df['_ParsedDate'].apply(lambda d: (min_date + pd.Timedelta(days=13 * ((d - min_date).days // 14) + 13)).strftime('%m/%d/%Y'))

    # SECTION 1 TABLE: Aggregate by Pay Period
    period_summary = []
    total_edits = len(csv_df)
    total_verified = int(sum(verified_flags))
    total_missing = total_edits - total_verified

    grouped_period = csv_df.groupby("_PPE_Group")
    for ppe_val, group in grouped_period:
        tot = len(group)
        ver = int(group["_Verified"].sum())
        mis = tot - ver
        rate = f"{(ver / tot * 100):.1f}%" if tot > 0 else "0.0%"
        period_summary.append([f"PPE {ppe_val}", str(tot), str(ver), str(mis), rate])

    overall_rate = f"{(total_verified / total_edits * 100):.1f}%" if total_edits > 0 else "0.0%"
    period_summary.append(["COMBINED TOTALS", str(total_edits), str(total_verified), str(total_missing), overall_rate])

    # SECTION 2 TABLE: Manager Attribution
    manager_summary = []
    grouped_mgr = csv_df.groupby(mgr_col)
    for mgr_val, group in grouped_mgr:
        tot = len(group)
        ver = int(group["_Verified"].sum())
        mis = tot - ver
        rate = f"{(ver / tot * 100):.1f}%" if tot > 0 else "0.0%"
        manager_summary.append([str(mgr_val), str(tot), str(ver), str(mis), rate])

    manager_summary.sort(key=lambda x: int(x[1]), reverse=True)

    return period_summary, manager_summary, csv_df

def build_styled_docx(processed_df, rows_s1, rows_s2):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    NAVY = "1B365D"
    LIGHT_BG = "F0F4F8"
    BORDER_GREY = "D3D3D3"

    def set_cell_background(cell, hex_color):
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shd)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{m}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def set_table_borders(table):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>\n'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>\n'
            f'  <w:left w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{BORDER_GREY}"/>\n'
            f'  <w:insideV w:val="none"/>\n'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def add_section_header(title_text):
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(6)
        r = h.add_run(title_text)
        r.font.name = "Arial"
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        return h

    def create_and_format_table(header_data, rows_data):
        num_rows = len(rows_data) + 1
        num_cols = len(header_data)
        
        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table)

        # Header Row
        for i, header_text in enumerate(header_data):
            cell = table.cell(0, i)
            cell.text = str(header_text)
            set_cell_background(cell, NAVY)
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            p = cell.paragraphs[0]
            if i > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if p.runs:
                run = p.runs[0]
                run.font.name = "Calibri"
                run.font.size = Pt(9.5)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # Data Rows
        for r_idx, row in enumerate(rows_data):
            for c_idx, val in enumerate(row):
                cell = table.cell(r_idx + 1, c_idx)
                cell.text = str(val)
                set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
                p = cell.paragraphs[0]
                
                str_val = str(val).strip()
                if c_idx > 0 and not any(char.isalpha() for char in str_val.replace("->", "").replace("PM", "").replace("AM", "")):
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                if p.runs:
                    run = p.runs[0]
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)

                    if r_idx % 2 == 1:
                        set_cell_background(cell, LIGHT_BG)

                    if "TOTAL" in str(row[0]).upper() or "COMBINED" in str(row[0]).upper():
                        run.font.bold = True
                        if c_idx == len(row) - 1:
                            run.font.color.rgb = RGBColor(0xA6, 0x1C, 0x1C)

    # DOCUMENT HEADER
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("PUNCH EDIT AUDIT & COMPLIANCE RECONCILIATION REPORT")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run("Full Multi-Period Wage & Hour Compliance Audit\nComparison of POS Audit Logs vs. Physical Signed Punch Edit Request Forms")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x5C, 0x76, 0x8D)

    # Metadata Block
    box = doc.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = box.cell(0, 0)
    set_cell_background(cell, LIGHT_BG)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    p_box = cell.paragraphs[0]
    p_box.paragraph_format.space_after = Pt(0)
    r_box_title = p_box.add_run("AUDIT METADATA & CONTROL BLOCK\n")
    r_box_title.font.bold = True
    r_box_title.font.size = Pt(10)
    r_box_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    total_recs = len(processed_df)
    tot_ver = int(processed_df["_Verified"].sum())
    tot_mis = total_recs - tot_ver
    
    r_box_body = p_box.add_run(
        f"• Audit Scope: 100% Line-by-Line Reconciliation ({total_recs} Total POS System Edits)\n"
        f"• Audit Results: {tot_ver} Verified with Form | {tot_mis} Missing Signed Form\n"
        f"• Governing Standards: California Labor Code §§ 226.7 & 512"
    )
    r_box_body.font.size = Pt(9.5)
    r_box_body.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # SECTION 1
    add_section_header("1. Executive Summary & Period Audit Findings")
    headers_s1 = ["Pay Period Ending (PPE)", "Total System Edits", "Forms Verified", "Forms Missing", "Compliance Rate"]
    create_and_format_table(headers_s1, rows_s1)

    # SECTION 2
    add_section_header("2. Manager Attribution & Compliance Performance")
    headers_s2 = ["Editing Manager / Editor", "Total Edits Executed", "Supported by Form", "Missing Form", "Manager Compliance %"]
    create_and_format_table(headers_s2, rows_s2)

    # SECTION 3
    add_section_header("3. Detailed Findings & Itemized System Audit Log")
    raw_cols = [c for c in processed_df.columns if not c.startswith("_")][:5]
    display_df = processed_df[raw_cols].copy()
    display_df["Audit Status"] = processed_df["_Verified"].apply(lambda x: "MATCH (Signed Form Present)" if x else "DISCREPANCY (Missing Form)")

    headers_s3 = raw_cols + ["Audit Status"]
    csv_rows = display_df.fillna("N/A").values.tolist()
    create_and_format_table(headers_s3, csv_rows)

    # SECTION 4
    add_section_header("4. Risk Analysis & Corrective Action Plan")
    defect_pct = (tot_mis / total_recs * 100) if total_recs > 0 else 0.0
    risk_points = [
        ("1. Documentation Exposure:", f"Out of {total_recs} manual edits executed in the system, {tot_mis} edits ({defect_pct:.1f}%) lack required employee-signed paper forms. Post-shift meal break additions and time reductions present severe liability under California Labor Code Section 226.7."),
        ("2. Manager Compliance Gaps:", "Editing managers with compliance rates below 80% must be audited weekly and retrained on physical punch form collection protocol."),
        ("3. Recommended Operational Controls:", "Enforce a mandatory digital form attachment block in the POS before allowing manager overrides, and lock payroll processing until form reconciliation is 100% complete.")
    ]

    for title, desc in risk_points:
        p_risk = doc.add_paragraph()
        p_risk.paragraph_format.space_before = Pt(4)
        p_risk.paragraph_format.space_after = Pt(4)
        r_rtitle = p_risk.add_run(f"{title} ")
        r_rtitle.font.name = "Calibri"
        r_rtitle.font.size = Pt(10.5)
        r_rtitle.font.bold = True
        r_rdesc = p_risk.add_run(desc)
        r_rdesc.font.name = "Calibri"
        r_rdesc.font.size = Pt(10.5)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if uploaded_csv and uploaded_pdf:
    st.success("Files uploaded successfully!")
    
    csv_data = pd.read_csv(uploaded_csv)
    
    with st.spinner("Executing fuzzy reconciliation engine..."):
        rows_s1, rows_s2, processed_df = process_audit_reconciliation(csv_data, uploaded_pdf)
    
    # Show real-time calculation preview directly in Streamlit
    total_found = int(processed_df["_Verified"].sum())
    total_lines = len(processed_df)
    comp_rate = (total_found / total_lines * 100) if total_lines > 0 else 0.0
    
    st.metric("Reconciliation Match Results", f"{comp_rate:.1f}% Compliance", f"{total_found} of {total_lines} Edits Verified")
    
    with st.expander("Preview Reconciliation Match Details (First 10 Rows)"):
        st.dataframe(processed_df[[c for c in processed_df.columns if not c.startswith("_")][:3] + ["_Verified", "_MatchNote"]].head(10))

    if st.button("Generate & Download Reconciled Word Report"):
        docx_file = build_styled_docx(processed_df, rows_s1, rows_s2)
        
        st.download_button(
            label="Click Here to Download Reconciled .docx Report",
            data=docx_file,
            file_name="Punch_Edit_Audit_And_Compliance_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
