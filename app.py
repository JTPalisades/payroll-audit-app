import io
import re
import pandas as pd
import pdfplumber
import streamlit as st
from docx import Document


def extract_pdf_pages(pdf_files) -> list[str]:
    """
    Extracts text page-by-page from uploaded PDF files.
    Returning pages individually allows matching names and dates on the same page.
    """
    pages_text = []
    for pdf_file in pdf_files:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages_text.append(text)
    return pages_text


def generate_date_variants(date_val) -> list[str]:
    """Generates common string variations of a date to match against PDF text."""
    if pd.isna(date_val):
        return []

    try:
        parsed_date = pd.to_datetime(date_val)
    except Exception:
        return [str(date_val).strip()]

    # Format into standard variants: 09/15/2026, 9/15/2026, 09/15/26, Sep 15, 2026, 15-Sep-2026
    variants = {
        parsed_date.strftime("%m/%d/%Y"),
        parsed_date.strftime("%n/%e/%Y").replace(" ", ""),  # Single-digit month/day
        parsed_date.strftime("%m/%d/%y"),
        parsed_date.strftime("%b %d, %Y"),
        parsed_date.strftime("%B %d, %Y"),
        parsed_date.strftime("%Y-%m-%d"),
    }
    return list(variants)


def analyze_edits_strict(df: pd.DataFrame, pdf_pages: list[str]) -> pd.DataFrame:
    """Analyzes manager edits and checks for employee name AND edit date on the same PDF page."""
    # Column auto-detection
    mgr_col = next((c for c in df.columns if "edit" in c and "by" in c or "manager" in c), df.columns[0])
    emp_col = next((c for c in df.columns if "employee" in c or "name" in c), df.columns[1])
    date_col = next((c for c in df.columns if "date" in c or "time" in c or "shift" in c), None)

    # Filter out system entries
    system_identifiers = ["system", "auto", "ghost", "house", "toast"]
    filtered_df = df[
        ~df[emp_col].astype(str).str.lower().isin(system_identifiers)
        & ~df[mgr_col].astype(str).str.lower().isin(system_identifiers)
    ].dropna(subset=[mgr_col]).copy()

    results = []

    for idx, row in filtered_df.iterrows():
        manager = str(row[mgr_col]).strip()
        employee = str(row[emp_col]).strip()
        date_raw = row[date_col] if date_col else None

        # Extract name tokens
        emp_tokens = [part for part in re.split(r"\s+|,", employee) if len(part) > 1]
        date_variants = generate_date_variants(date_raw) if date_raw is not None else []

        has_valid_form = False

        # Scan page by page
        for page_text in pdf_pages:
            # Check 1: Does the page contain the employee's name?
            name_match = all(
                re.search(r"\b" + re.escape(token) + r"\b", page_text, re.IGNORECASE)
                for token in emp_tokens
            ) if emp_tokens else False

            # Check 2: Does the page contain any variant of the edit date?
            date_match = any(
                re.search(r"\b" + re.escape(d_var) + r"\b", page_text, re.IGNORECASE)
                for d_var in date_variants
            ) if date_variants else True  # Fall back to True if no date column exists

            if name_match and date_match:
                has_valid_form = True
                break  # Stop checking pages once matched

        results.append({
            "Manager": manager,
            "Employee": employee,
            "Edit_Date": str(date_raw) if date_raw is not None else "N/A",
            "Has_Signed_Form": has_valid_form
        })

    res_df = pd.DataFrame(results)

    # Summarize by Manager
    summary = (
        res_df.groupby("Manager")
        .agg(
            Total_Edits=("Has_Signed_Form", "count"),
            Forms_Present=("Has_Signed_Form", "sum")
        )
        .reset_index()
    )

    summary["Forms_Missing"] = summary["Total_Edits"] - summary["Forms_Present"]
    summary["Compliance_Pct"] = (
        (summary["Forms_Present"] / summary["Total_Edits"]) * 100
    ).round(1)

    return summary
