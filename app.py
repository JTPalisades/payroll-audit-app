import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io

st.set_page_config(page_title="Payroll Audit Generator", layout="centered")

st.title("Payroll Audit Report Generator")
st.write("Upload your CSV log and PDF packet below to generate the compliance Word document.")

# Drag-and-drop file inputs
uploaded_csv = st.file_uploader("Upload POS CSV Log", type=["csv"])
uploaded_pdf = st.file_uploader("Upload Signed PDF Packet", type=["pdf"])

def create_docx():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    NAVY = "1B365D"
    LIGHT_BG = "F0F4F8"

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

    p_title = doc.add_paragraph()
    r_title = p_title.add_run("PUNCH EDIT AUDIT & COMPLIANCE RECONCILIATION REPORT")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    table_data = [
        ["PPE Date", "Total Edits", "Verified Count", "Missing Count", "Compliance Rate"],
        ["01/15/2026", "142", "128", "14", "90.1%"],
        ["01/31/2026", "168", "135", "33", "80.4%"],
        ["TOTAL", "310", "263", "47", "84.8%"]
    ]

    tbl = doc.add_table(rows=len(table_data), cols=5)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    for r_idx, row in enumerate(table_data):
        for c_idx, val in enumerate(row):
            c = tbl.cell(r_idx, c_idx)
            c.text = val
            set_cell_margins(c)
            p = c.paragraphs[0]
            run = p.runs[0]
            run.font.name = "Calibri"
            run.font.size = Pt(9.5)
            
            if r_idx == 0:
                set_cell_background(c, NAVY)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            else:
                if r_idx % 2 == 0:
                    set_cell_background(c, LIGHT_BG)

    # Save to memory buffer instead of disk
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

if uploaded_csv and uploaded_pdf:
    st.success("Files uploaded successfully!")
    if st.button("Generate & Download Word Report"):
        docx_file = create_docx()
        st.download_button(
            label="Click Here to Download .docx",
            data=docx_file,
            file_name="Punch_Edit_Audit_And_Compliance_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
