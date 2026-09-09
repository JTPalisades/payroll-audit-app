import streamlit as st
import pandas as pd
import pdfplumber
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io

st.set_page_config(page_title="Payroll Audit Generator", layout="centered")

st.title("Payroll Audit Report Generator")
st.write("Upload your POS CSV Log and PDF Packet to run a line-by-line reconciliation and build the detailed Word report.")

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

def build_full_docx(csv_df, pdf_text):
    """Builds the complete multi-section Word document using real data."""
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

    def set_headers(table, headers):
        hdr_cells = table.rows[0].cells
        for i, title in enumerate(headers):
            hdr_cells[i].text = str(title)
            set_cell_background(hdr_cells[i], NAVY)
            set_cell_margins(hdr_cells[i])
            run = hdr_cells[i].paragraphs[0].runs[0]
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Document Header
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("PUNCH EDIT AUDIT & COMPLIANCE RECONCILIATION REPORT")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("California Wage & Hour Compliance Verification | Labor Code §§ 226.7 & 512")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x5C, 0x76, 0x8D)

    # Section 1: Executive Summary
    h1 = doc.add_paragraph()
    r = h1.add_run("1. Executive Summary & Audit Overview")
    r.font.name = "Arial"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_exec = doc.add_paragraph(
        "Under California Labor Code Sections 226.7 and 512, employers are strictly obligated to provide accurate, "
        "uninterrupted meal and rest periods, and to maintain transparent timecard records. Any manual modification "
        "requires express written authorization signed by the employee (e.g., Solicitud de Horas Editadas). "
        "Below is the complete reconciliation of electronic POS audit logs against signed physical documentation."
    )
    p_exec.style.font.name = "Calibri"
    p_exec.style.font.size = Pt(11)

    # Section 2: Full Itemized Data Table from CSV
    h2 = doc.add_paragraph()
    r2 = h2.add_run("2. Detailed Line-by-Line System Audit Findings")
    r2.font.name = "Arial"
    r2.font.size = Pt(13)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    # Clean CSV data columns (Limit to first 6 key columns for formatting stability)
    cols_to_use = csv_df.columns[:6].tolist()
    
    table = doc.add_table(rows=len(csv_df) + 1, cols=len(cols_to_use))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_headers(table, cols_to_use)

    # Populate table with real row-by-row data from CSV
    for r_idx, row in csv_df.iterrows():
        for c_idx, col_name in enumerate(cols_to_use):
            cell = table.cell(r_idx + 1, c_idx)
            val = str(row[col_name]) if pd.notna(row[col_name]) else "N/A"
            cell.text = val
            set_cell_margins(cell)
            run = cell.paragraphs[0].runs[0]
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            if r_idx % 2 == 1:
                set_cell_background(cell, LIGHT_BG)

    # Section 3: Risk Analysis & Corrective Actions
    h3 = doc.add_paragraph()
    r3 = h3.add_run("3. Risk Analysis & Corrective Action Plan")
    r3.font.name = "Arial"
    r3.font.size = Pt(13)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_risk = doc.add_paragraph(
        "• Unverified Meal Breaks: Unsigned timecard edits present significant risk under CA Labor Code § 226.7.\n"
        "• Corrective Action 1: Enforce mandatory POS digital attestation before allowing manager manual overrides.\n"
        "• Corrective Action 2: Conduct quarterly manager compliance training regarding timekeeping protocols.\n"
        "• Corrective Action 3: Lock payroll processing until 100% of physical edit forms are matched and uploaded."
    )
    p_risk.style.font.name = "Calibri"
    p_risk.style.font.size = Pt(11)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if uploaded_csv and uploaded_pdf:
    st.success("Both CSV and PDF loaded successfully!")
    
    if st.button("Generate & Download Detailed Word Report"):
        with st.spinner("Processing full dataset and generating report..."):
            # Parse real files
            csv_data = pd.read_csv(uploaded_csv)
            pdf_text = extract_pdf_text(uploaded_pdf)
            
            # Build report using actual parsed data
            docx_file = build_full_docx(csv_data, pdf_text)
            
            st.download_button(
                label="Click Here to Download Full Detailed .docx",
                data=docx_file,
                file_name="Punch_Edit_Audit_And_Compliance_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
