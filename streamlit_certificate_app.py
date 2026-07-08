import os
import re
import zipfile
import shutil
import tempfile
import subprocess
from pathlib import Path

import pandas as pd
import streamlit as st
from pptx import Presentation
from typing import Tuple


# -----------------------------
# Helper functions
# -----------------------------
REQUIRED_COLUMNS = [
    "Name of Student",
    "Subject Name",
    "Total Marks",
    "Grades",
    "Enrollment No.",
]

PLACEHOLDERS = {
    "{{Name}}": "Name of Student",
    "{{Course}}": "Subject Name",
    "{{Marks}}": "Total Marks",
    "{{Grade}}": "Grades",
    "{{Enrollment}}": "Enrollment No.",
}


def safe_filename(text: str) -> str:
    """Remove characters that are not allowed in Windows/Mac file names."""
    text = str(text).strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text)
    return text or "certificate"


def replace_text(shape, replacements: dict) -> None:
    """Replace placeholder text inside a normal text shape."""
    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)

        if not full_text:
            continue

        updated = full_text
        for key, value in replacements.items():
            updated = updated.replace(key, str(value))

        if updated != full_text and paragraph.runs:
            paragraph.runs[0].text = updated
            for run in paragraph.runs[1:]:
                run.text = ""


def process_shape(shape, replacements: dict) -> None:
    """Process normal shapes and grouped shapes recursively."""
    if shape.shape_type == 6:  # Group shape
        for shp in shape.shapes:
            process_shape(shp, replacements)
        return

    replace_text(shape, replacements)


def create_zip(source_folder: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in source_folder.iterdir():
            if file.is_file():
                zipf.write(file, arcname=file.name)


def convert_pptx_to_pdf(ppt_folder: Path, pdf_folder: Path) -> Tuple[bool, str]:
    """Convert PPTX files to PDF using LibreOffice if available."""
    pdf_folder.mkdir(parents=True, exist_ok=True)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False, "LibreOffice is not installed on this system. PDF conversion was skipped."

    ppt_files = list(ppt_folder.glob("*.pptx"))
    for ppt_file in ppt_files:
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_folder),
                str(ppt_file),
            ],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

    return True, "PDF conversion completed."


def generate_certificates(
    df: pd.DataFrame,
    template_path: Path,
    output_folder: Path,
    date_issued: str,
    progress_bar=None,
    status_box=None,
) -> tuple[int, int, list[dict]]:
    output_folder.mkdir(parents=True, exist_ok=True)

    passed = 0
    failed = 0
    failed_students = []
    total = len(df)

    if total == 0:
        return 0, 0, []

    for index, row in df.iterrows():
        grade = str(row["Grades"]).strip().upper()

        if grade == "F":
            failed += 1
            failed_students.append(row.to_dict())
            if status_box:
                status_box.info(f"Skipped: {row['Name of Student']} because Grade is F")
        else:
            prs = Presentation(str(template_path))

            replacements = {
                "{{Name}}": row["Name of Student"],
                "{{Course}}": row["Subject Name"],
                "{{Marks}}": row["Total Marks"],
                "{{Grade}}": row["Grades"],
                "{{DateIssued}}": date_issued,
                "{{Enrollment}}": row["Enrollment No."],
            }

            for slide in prs.slides:
                for shape in slide.shapes:
                    process_shape(shape, replacements)

            filename = safe_filename(f"{row['Enrollment No.']}_{row['Name of Student']}.pptx")
            prs.save(output_folder / filename)
            passed += 1

            if status_box:
                status_box.success(f"Generated: {filename}")

        if progress_bar:
            progress_bar.progress((index + 1) / total)

    return passed, failed, failed_students


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(
    page_title="Certificate Generator",
    page_icon="🎓",
    layout="wide",
)

st.title("🎓 Excel to PPT Certificate Generator")
st.caption("Upload an Excel file and a PowerPoint certificate template to generate personalized certificates.")

with st.sidebar:
    st.header("⚙️ Settings")
    date_issued = st.text_input("Date Issued", value="07 July 2026")
    convert_pdf = st.checkbox("Also convert PPT certificates to PDF", value=False)

    st.markdown("---")
    st.subheader("Template Placeholders")
    st.write("Use these placeholders inside your PPT template:")
    st.code(
        "{{Name}}\n{{Course}}\n{{Marks}}\n{{Grade}}\n{{DateIssued}}\n{{Enrollment}}",
        language="text",
    )

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])

with col2:
    ppt_template = st.file_uploader("Upload PPT Template", type=["pptx"])

st.markdown("---")

if excel_file:
    try:
        preview_df = pd.read_excel(excel_file).fillna("")
        st.subheader("📊 Excel Preview")
        st.dataframe(preview_df.head(10), use_container_width=True)

        missing_columns = [col for col in REQUIRED_COLUMNS if col not in preview_df.columns]
        if missing_columns:
            st.error("Missing required columns: " + ", ".join(missing_columns))
        else:
            total_students = len(preview_df)
            pass_count = len(preview_df[preview_df["Grades"].astype(str).str.strip().str.upper() != "F"])
            fail_count = total_students - pass_count

            m1, m2, m3 = st.columns(3)
            m1.metric("Total Students", total_students)
            m2.metric("Certificates to Generate", pass_count)
            m3.metric("Skipped Grade F", fail_count)

    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        preview_df = None
else:
    preview_df = None

button_disabled = not excel_file or not ppt_template or preview_df is None

if st.button("🚀 Generate Certificates", disabled=button_disabled, type="primary"):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        excel_path = tmpdir / "students.xlsx"
        template_path = tmpdir / "template.pptx"
        output_folder = tmpdir / "Certificates"
        pdf_folder = tmpdir / "PDF_Certificates"
        ppt_zip = tmpdir / "Certificates.zip"
        pdf_zip = tmpdir / "PDF_Certificates.zip"
        skipped_file = tmpdir / "Skipped_Students.xlsx"

        excel_path.write_bytes(excel_file.getvalue())
        template_path.write_bytes(ppt_template.getvalue())

        df = pd.read_excel(excel_path).fillna("")

        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_columns:
            st.error("Cannot generate. Missing columns: " + ", ".join(missing_columns))
        else:
            st.subheader("🔄 Generation Progress")
            progress_bar = st.progress(0)
            status_box = st.empty()

            passed, failed, failed_students = generate_certificates(
                df=df,
                template_path=template_path,
                output_folder=output_folder,
                date_issued=date_issued,
                progress_bar=progress_bar,
                status_box=status_box,
            )

            create_zip(output_folder, ppt_zip)

            st.success("Certificate generation completed.")

            r1, r2, r3 = st.columns(3)
            r1.metric("Total Students", len(df))
            r2.metric("PPT Certificates", passed)
            r3.metric("Skipped", failed)

            st.download_button(
                label="⬇️ Download PPT Certificates ZIP",
                data=ppt_zip.read_bytes(),
                file_name="Certificates.zip",
                mime="application/zip",
            )

            if failed_students:
                pd.DataFrame(failed_students).to_excel(skipped_file, index=False)
                st.download_button(
                    label="⬇️ Download Skipped Students Excel",
                    data=skipped_file.read_bytes(),
                    file_name="Skipped_Students.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            if convert_pdf:
                with st.spinner("Converting PPT files to PDF..."):
                    success, message = convert_pptx_to_pdf(output_folder, pdf_folder)

                if success and any(pdf_folder.glob("*.pdf")):
                    create_zip(pdf_folder, pdf_zip)
                    st.success(message)
                    st.download_button(
                        label="⬇️ Download PDF Certificates ZIP",
                        data=pdf_zip.read_bytes(),
                        file_name="PDF_Certificates.zip",
                        mime="application/zip",
                    )
                else:
                    st.warning(message)

st.markdown("---")
st.info(
    "Note: For PDF conversion, LibreOffice must be installed on your computer or server. "
    "PPT generation works with Python packages only."
)
