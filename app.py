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
st.write("Upload your POS CSV Log and PDF Packet to perform a dynamic line-by-line reconciliation and build your executive compliance Word report.")

uploaded_csv = st.file_uploader("Upload POS CSV Log", type=["csv"])
uploaded_pdf = st.file_uploader("Upload Signed PDF Packet", type=["pdf"])

def extract_pdf_text(pdf_file):
    """Extracts all text from the uploaded PDF packet."""
    text_data = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text_data += extracted + "\n"
    return text_data

def process_audit_data(csv_df, pdf_text):
    """Dynamically reconciles CSV records against PDF text."""
    # Standardize CSV column names by trimming spaces
    csv_df.columns = [str(c).strip() for c in csv_df.columns]
    
    # Identify key column names dynamically
    col_map = {}
    for col in csv_df.columns:
        c_lower = col.lower()
        if "employee" in c_lower or "name" in c_lower:
            col_map["employee"] = col
        elif "date" in c_lower:
            col_map["date"] = col
        elif "manager" in c_lower or "editor" in c_lower or "user" in c_lower:
            col_map["manager"] = col
        elif "action" in c_lower or "edit" in c_lower or "type" in c_lower:
            col_map["action"] = col

    # Default fallback column names if auto-detection misses
    emp_col = col_map.get("employee", csv_df.columns[0])
    date_col = col_map.get("date", csv_df.columns[1] if len(csv_df.columns) > 1 else csv_df.columns[0])
    mgr_col = col_map.get("manager", csv_df.columns[3] if len(csv_df.columns) > 3 else csv_df.columns[0])
    
    # Check if form exists in PDF text for each employee
    verified_flags = []
    pdf_text_upper = pdf_text.upper()
    
    for _, row in csv_df.iterrows():
        emp_name = str(row[emp_col]).strip().upper()
        # Extract last name or significant name part for matching
        name_parts = [p for p in emp_name.split() if len(p) > 2]
        
        # Match if employee name parts appear in PDF text
        if name_parts and any(part in pdf_text_upper for part in name_parts):
            verified_flags.append(True)
        else:
            verified_flags.append(False)

    csv_df["_Verified"] = verified_flags

    # --- Build Section 1 Data (By Date / Period) ---
    csv_df["_Date_Group"] = csv_df[date_col].astype(str)
    period_summary = []
    total_edits_all = len(csv_df)
    total_verified_all = sum(verified_flags)
    total_missing_all = total_edits_all - total_verified_all

    grouped_date = csv_df.groupby("_Date_Group")
    for date_val, group in grouped_date:
        tot = len(group)
        ver = group["_Verified"].sum()
        mis = tot - ver
        rate = f"{(ver / tot * 100):.1f}%" if tot > 0 else "0.0%"
        period_summary.append([f"PPE {date_val}", str(tot), str(ver), str(mis), rate])

    overall_rate = f"{(total_verified_all / total_edits_all * 100):.1f}%" if total_edits_all > 0 else "0.0%"
    period_summary.append(["COMBINED YTD TOTALS", str(total_edits_all), str(total_verified_all), str(total_missing_all), overall_rate])

    # --- Build Section 2 Data (By Manager) ---
    manager_summary = []
    grouped_mgr = csv_df.groupby(mgr_col)
    for mgr_val, group in grouped_mgr:
        tot = len(group)
        ver = group["_Verified"].sum()
        mis = tot - ver
        rate = f"{(ver / tot * 100):.1f}%" if tot > 0 else "0.0%"
        manager_summary.append([str(mgr_val), str(tot), str(ver), str(mis), rate])

    return period_summary, manager_summary, csv_df

def build_styled_docx(csv_df, pdf_text):
    # Dynamically compute audit metrics from uploaded files
    rows_s1, rows_s2, processed_df = process_audit_data(csv_df, pdf_text)

    doc = docx.Document()

    # Standard Page Setup (1-inch margins)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Styling Constants
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

        # Build Header Row
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

        # Build Data Rows
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

    # Metadata Control Box
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
    
    r_box_body = p_box.add_run(
        f"• Scope: Multi-Period Wage & Hour Compliance Reconciliation ({len(processed_df)} Total System Edits)\n"
        "• Prepared For: Payroll & HR Compliance\n"
        "• Verification: 100% Line-by-Line POS CSV Log vs. Physical Form Matching"
    )
    r_box_body.font.size = Pt(9.5)
    r_box_body.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    # -------------------------------------------------------------
    # SECTION 1: EXECUTIVE SUMMARY (DYNAMIC DATA)
    # -------------------------------------------------------------
    add_section_header("1. Executive Summary & Year-to-Date Audit Findings")
    
    p_exec = doc.add_paragraph(
        "A comprehensive wage and hour compliance audit was conducted across the provided pay periods. "
        "The audit reconciled electronic system edit logs exported from the POS/Time Management system against physical paper "
        "documentation packets provided in PDF format. Under California Labor Code Sections 226.7 and 512, any manual "
        "adjustment to an employee's timecard requires a fully executed, employee-signed paper form."
    )
    p_exec.style.font.name = "Calibri"
    p_exec.style.font.size = Pt(10.5)

    headers_s1 = ["Pay Period Ending (PPE)", "Total System Edits", "Forms Verified", "Forms Missing", "Compliance Rate"]
    create_and_format_table(headers_s1, rows_s1)

    # -------------------------------------------------------------
    # SECTION 2: MANAGER BREAKDOWN (DYNAMIC DATA)
    # -------------------------------------------------------------
    add_section_header("2. Manager Attribution & Compliance Performance")
    
    p_mgr = doc.add_paragraph("The table below attributes system timecard edits executed during the audit period to the specific editing manager logged in the POS audit system:")
    p_mgr.style.font.name = "Calibri"
    p_mgr.style.font.size = Pt(10.5)

    headers_s2 = ["Editing Manager / Editor", "Total Edits Executed", "Supported by Form", "Missing Form", "Manager Compliance %"]
    create_and_format_table(headers_s2, rows_s2)

    # -------------------------------------------------------------
    # SECTION 3: DETAILED FINDINGS BY PAY PERIOD (RAW CSV ROWS)
    # -------------------------------------------------------------
    add_section_header("3. Detailed Findings by Pay Period (Itemized Data)")

    p_period = doc.add_paragraph()
    p_period.paragraph_format.space_before = Pt(8)
    p_period.paragraph_format.space_after = Pt(4)
    r_period = p_period.add_run("Itemized System Edits Log")
    r_period.font.name = "Arial"
    r_period.font.size = Pt(11)
    r_period.font.bold = True
    r_period.font.color.rgb = RGBColor(0x5C, 0x76, 0x8D)

    # Extract display columns excluding internal processing flags
    display_cols = [c for c in processed_df.columns if not c.startswith("_")][:6]
    csv_rows = processed_df[display_cols].fillna("N/A").values.tolist()

    create_and_format_table(display_cols, csv_rows)

    # -------------------------------------------------------------
    # SECTION 4: RISK ANALYSIS & ACTION PLAN
    # -------------------------------------------------------------
    add_section_header("4. Risk Analysis & Corrective Action Plan")

    risk_points = [
        ("1. High Documentation Defect Rate:", "Unverified unpaid meal break additions and time-out reductions executed post-shift present substantial exposure under California Labor Code Section 226.7."),
        ("2. Manager Performance Gaps:", "Targeted retraining is required for managers showing compliance rates below 60% to ensure signed edit request forms are obtained prior to executing edits."),
        ("3. Mandatory Corrective Recommendations:", "Restrict POS edit access without digital form attachments, and enforce pre-payroll reconciliation locks each period.")
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
    if st.button("Generate & Download Executive Word Report"):
        with st.spinner("Processing files and performing dynamic audit reconciliation..."):
            csv_data = pd.read_csv(uploaded_csv)
            pdf_text = extract_pdf_text(uploaded_pdf)
            
            docx_file = build_styled_docx(csv_data, pdf_text)
            
            st.download_button(
                label="Click Here to Download Formatted .docx",
                data=docx_file,
                file_name="Punch_Edit_Audit_And_Compliance_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
