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
st.write("Upload your POS CSV Log and PDF Packet to perform a pay-period aggregated reconciliation and generate your Word audit report.")

uploaded_csv = st.file_uploader("Upload POS CSV Log", type=["csv"])
uploaded_pdf = st.file_uploader("Upload Signed PDF Packet", type=["pdf"])

def extract_pdf_records(pdf_file):
    """
    Extracts text from PDF and builds a structured lookup map of 
    (Normalized Employee Name, Normalized Date) present on physical signed forms.
    """
    pdf_records = set()
    pdf_full_text = ""
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            pdf_full_text += text + "\n"
            
            lines = text.split('\n')
            for line in lines:
                clean_line = line.upper().strip()
                dates = re.findall(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', clean_line)
                
                if dates:
                    for d in dates:
                        # Extract valid employee names (words >= 3 letters)
                        words = re.findall(r'[A-Z]{3,}', clean_line)
                        for word in words:
                            pdf_records.add((word, d))
                            
    return pdf_records, pdf_full_text

def process_audit_reconciliation(csv_df, pdf_file):
    """
    Performs line-by-line matching and aggregates data by Pay Period Ending (PPE) dates.
    """
    # Clean CSV column names
    csv_df.columns = [str(c).strip() for c in csv_df.columns]
    
    # Identify key columns dynamically
    emp_col = next((c for c in csv_df.columns if "employee" in c.lower() or "name" in c.lower()), csv_df.columns[0])
    date_col = next((c for c in csv_df.columns if "date" in c.lower() or "shift" in c.lower()), csv_df.columns[1] if len(csv_df.columns) > 1 else csv_df.columns[0])
    mgr_col = next((c for c in csv_df.columns if "manager" in c.lower() or "editor" in c.lower() or "user" in c.lower()), csv_df.columns[3] if len(csv_df.columns) > 3 else csv_df.columns[0])

    # Convert Date column to actual DateTime objects to group by 14-day Pay Periods
    csv_df['_ParsedDate'] = pd.to_datetime(csv_df[date_col], errors='coerce')
    
    # Drop invalid date rows if any
    csv_df = csv_df.dropna(subset=['_ParsedDate']).copy()
    csv_df = csv_df.sort_values('_ParsedDate')

    # Extract text & records from PDF
    pdf_records, pdf_full_text = extract_pdf_records(pdf_file)
    pdf_text_upper = pdf_full_text.upper()

    verified_flags = []

    for _, row in csv_df.iterrows():
        emp_name = str(row[emp_col]).strip().upper()
        raw_date_str = row['_ParsedDate'].strftime('%m/%d/%Y')
        short_date_str = row['_ParsedDate'].strftime('%m/%d/%y')

        name_parts = [p for p in re.findall(r'[A-Z]{3,}', emp_name) if p not in ["THE", "AND", "SERVER", "BUSSER", "HOST", "KITCHEN"]]

        is_verified = False
        if name_parts:
            for part in name_parts:
                # Check tuple match
                if (part, raw_date_str) in pdf_records or (part, short_date_str) in pdf_records:
                    is_verified = True
                    break
                # Fallback check in full text
                if part in pdf_text_upper and (raw_date_str in pdf_text_upper or short_date_str in pdf_text_upper):
                    is_verified = True
                    break

        verified_flags.append(is_verified)

    csv_df["_Verified"] = verified_flags

    # Assign Bi-Weekly Pay Period Ending (PPE) Dates (Grouping into 14-day blocks)
    # Group dates by 14-day frequency ending on the max date
    min_date = csv_df['_ParsedDate'].min()
    csv_df['_PPE_Group'] = csv_df['_ParsedDate'].apply(lambda d: (min_date + pd.Timedelta(days=13 * ((d - min_date).days // 14) + 13)).strftime('%m/%d/%Y'))

    # -------------------------------------------------------------
    # 1. SECTION 1 TABLE: Aggregate by Pay Period Ending (PPE)
    # -------------------------------------------------------------
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

    # -------------------------------------------------------------
    # 2. SECTION 2 TABLE: Manager Attribution
    # -------------------------------------------------------------
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

def build_styled_docx(csv_df, pdf_file):
    rows_s1, rows_s2, processed_df = process_audit_reconciliation(csv_df, pdf_file)

    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    NAVY = "1B365D"
    STEEL_BLUE = "5C768D"
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

    # -------------------------------------------------------------
    # DOCUMENT HEADER
    # -------------------------------------------------------------
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

    # -------------------------------------------------------------
    # SECTION 1: EXECUTIVE SUMMARY
    # -------------------------------------------------------------
    add_section_header("1. Executive Summary & Period Audit Findings")
    
    p_exec = doc.add_paragraph(
        "A comprehensive wage and hour compliance audit was conducted reconciling electronic POS system edit logs "
        "against physical paper documentation packets. Under California Labor Code Sections 226.7 and 512, any manual "
        "adjustment to an employee's timecard requires a fully executed, employee-signed paper form."
    )
    p_exec.style.font.name = "Calibri"
    p_exec.style.font.size = Pt(10.5)

    headers_s1 = ["Pay Period Ending (PPE)", "Total System Edits", "Forms Verified", "Forms Missing", "Compliance Rate"]
    create_and_format_table(headers_s1, rows_s1)

    # -------------------------------------------------------------
    # SECTION 2: MANAGER BREAKDOWN
    # -------------------------------------------------------------
    add_section_header("2. Manager Attribution & Compliance Performance")
    
    p_mgr = doc.add_paragraph("The table below attributes system timecard edits executed during the audit period to the specific editing manager logged in the POS audit system:")
    p_mgr.style.font.name = "Calibri"
    p_mgr.style.font.size = Pt(10.5)

    headers_s2 = ["Editing Manager / Editor", "Total Edits Executed", "Supported by Form", "Missing Form", "Manager Compliance %"]
    create_and_format_table(headers_s2, rows_s2)

    # -------------------------------------------------------------
    # SECTION 3: DETAILED FINDINGS BY PAY PERIOD
    # -------------------------------------------------------------
    add_section_header("3. Detailed Findings & Itemized System Audit Log")

    raw_cols = [c for c in processed_df.columns if not c.startswith("_")][:5]
    display_df = processed_df[raw_cols].copy()
    display_df["Audit Status"] = processed_df["_Verified"].apply(lambda x: "MATCH (Signed Form Present)" if x else "DISCREPANCY (Missing Form)")

    headers_s3 = raw_cols + ["Audit Status"]
    csv_rows = display_df.fillna("N/A").values.tolist()

    create_and_format_table(headers_s3, csv_rows)

    # -------------------------------------------------------------
    # SECTION 4: RISK ANALYSIS & ACTION PLAN
    # -------------------------------------------------------------
    add_section_header("4. Risk Analysis & Corrective Action Plan")

    risk_points = [
        ("1. Documentation Exposure:", f"Out of {total_recs} manual edits executed in the system, {tot_mis} edits ({(tot_mis/total_recs*100):.1f}%) lack required employee-signed paper forms. Post-shift meal break additions and time reductions present severe liability under California Labor Code Section 226.7."),
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
    st.success("Files loaded successfully!")
    if st.button("Generate & Download Reconciled Word Report"):
        with st.spinner("Aggregating pay periods and performing reconciliation..."):
            csv_data = pd.read_csv(uploaded_csv)
            docx_file = build_styled_docx(csv_data, uploaded_pdf)
            
            st.download_button(
                label="Click Here to Download Reconciled .docx Report",
                data=docx_file,
                file_name="Punch_Edit_Audit_And_Compliance_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
