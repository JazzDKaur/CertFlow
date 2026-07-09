import os
import re
import zipfile
import shutil
import tempfile
import subprocess
import calendar
import hashlib
from datetime import date
from pathlib import Path
from typing import Tuple, Optional

import pandas as pd
import streamlit as st
from pptx import Presentation



# -----------------------------
# Helper functions
# -----------------------------
REQUIRED_COLUMNS = [
    "Name of Student",
    "Subject Name",
    "Total Marks",
    "Grades",
    "Enrollment No.",
    "Exam Cycle",
]

PLACEHOLDERS = {
    "{{Name}}": "Name of Student",
    "{{Course}}": "Subject Name",
    "{{Marks}}": "Total Marks",
    "{{Grade}}": "Grades",
    "{{Enrollment}}": "Enrollment No.",
    "{{DateIssued}}": "Calculated from Exam Cycle",
}


MONTH_LOOKUP = {
    "JAN": 1,
    "JANUARY": 1,
    "FEB": 2,
    "FEBRUARY": 2,
    "MAR": 3,
    "MARCH": 3,
    "APR": 4,
    "APRIL": 4,
    "MAY": 5,
    "JUN": 6,
    "JUNE": 6,
    "JUL": 7,
    "JULY": 7,
    "AUG": 8,
    "AUGUST": 8,
    "SEP": 9,
    "SEPT": 9,
    "SEPTEMBER": 9,
    "OCT": 10,
    "OCTOBER": 10,
    "NOV": 11,
    "NOVEMBER": 11,
    "DEC": 12,
    "DECEMBER": 12,
}

def install_custom_fonts():
    """
    Install custom fonts for Streamlit Cloud/Linux.
    On Windows local system, skip fc-cache because Windows does not have it.
    """

    import os
    import shutil
    import subprocess
    from pathlib import Path

    fonts_source = Path("fonts")

    if not fonts_source.exists():
        return

    # If running on Windows, do not run fc-cache
    if os.name == "nt":
        return

    fonts_target = Path.home() / ".local" / "share" / "fonts"
    fonts_target.mkdir(parents=True, exist_ok=True)

    for pattern in ("*.ttf", "*.TTF", "*.otf", "*.OTF"):
        for font_file in fonts_source.glob(pattern):
            shutil.copy2(font_file, fonts_target / font_file.name)

    subprocess.run(
        ["fc-cache", "-f", "-v"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


def safe_filename(text: str) -> str:
    """Remove characters that are not allowed in Windows/Mac/Linux file names."""
    text = str(text).strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text)
    return text or "certificate"


def get_ordinal_suffix(day: int) -> str:
    """Return st, nd, rd, or th suffix for a day number."""
    if 11 <= day % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


def format_certificate_date(date_value: date) -> str:
    """Format date as dd/mm/yyyy."""
    return date_value.strftime("%d/%m/%Y")


def last_day_of_month(year: int, month: int) -> date:
    """Return the last date of a given month."""
    return date(year, month, calendar.monthrange(year, month)[1])


def get_last_completed_month_end(today: Optional[date] = None) -> date:
    """Return the last date of the previous completed month."""
    today = today or date.today()
    if today.month == 1:
        return last_day_of_month(today.year - 1, 12)
    return last_day_of_month(today.year, today.month - 1)


def parse_exam_cycle(exam_cycle: str) -> tuple[int, int]:
    """
    Parse exam cycle values like Jan26, JAN26, Jan-26, Jan 26, January 2026.
    Returns (month, year).
    """
    value = str(exam_cycle).strip()
    if not value:
        raise ValueError("Exam Cycle is blank")

    cleaned = re.sub(r"[^A-Za-z0-9]", "", value).upper()
    match = re.match(r"^([A-Z]+)(\d{2}|\d{4})$", cleaned)

    if not match:
        raise ValueError(f"Invalid Exam Cycle format: {exam_cycle}. Use values like Jan26, Apr26, Sep26.")

    month_text, year_text = match.groups()
    if month_text not in MONTH_LOOKUP:
        raise ValueError(f"Invalid month in Exam Cycle: {exam_cycle}")

    year = int(year_text)
    if len(year_text) == 2:
        year += 2000

    return MONTH_LOOKUP[month_text], year


def calculate_date_issued(exam_cycle: str, today: Optional[date] = None) -> str:
    """
    Calculate certificate issue date from Exam Cycle.
    If the exam cycle is in the future or current running month, use the last completed month end.
    """
    today = today or date.today()
    month, year = parse_exam_cycle(exam_cycle)
    exam_cycle_end = last_day_of_month(year, month)
    last_completed_end = get_last_completed_month_end(today)

    issue_date = min(exam_cycle_end, last_completed_end)
    return format_certificate_date(issue_date)


def add_date_issued_preview(df: pd.DataFrame) -> pd.DataFrame:
    """Add a calculated Date Issued column for preview and generation."""
    df = df.copy()
    df["Date Issued"] = df["Exam Cycle"].apply(calculate_date_issued)
    return df


def replace_text(shape, replacements: dict) -> None:
    """
    Replace placeholder text while keeping the same font/style as the PPT template.
    No custom font is applied.
    """
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
            # Put replaced text in the first existing run.
            # This preserves the template font/style of that placeholder run.
            paragraph.runs[0].text = updated

            # Clear extra runs to avoid duplicate/old placeholder text.
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
        for file in sorted(source_folder.iterdir()):
            if file.is_file():
                zipf.write(file, arcname=file.name)


def convert_pptx_to_pdf(ppt_folder: Path, pdf_folder: Path) -> Tuple[bool, str]:
    """
    Convert all PPTX files in a folder to PDF using LibreOffice.
    Works on Windows, Linux and Streamlit Cloud.
    """

    import os
    import shutil
    import subprocess

    pdf_folder.mkdir(parents=True, exist_ok=True)

    # Try to locate LibreOffice
    possible_paths = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    soffice = None
    for path in possible_paths:
        if path and os.path.exists(path):
            soffice = path
            break

    if soffice is None:
        return (
            False,
            "LibreOffice (soffice.exe) was not found. "
            "Install LibreOffice or add soffice.exe to your PATH."
        )

    ppt_files = list(ppt_folder.glob("*.pptx"))

    if not ppt_files:
        return False, "No PPTX files found."

    errors = []
    converted = 0

    for ppt_file in ppt_files:

        command = [
            soffice,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_folder),
            str(ppt_file),
        ]

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        expected_pdf = pdf_folder / f"{ppt_file.stem}.pdf"

        if expected_pdf.exists():
            converted += 1
        else:
            errors.append(
                f"{ppt_file.name}\n"
                f"Return Code: {result.returncode}\n"
                f"STDOUT: {result.stdout}\n"
                f"STDERR: {result.stderr}"
            )

    if converted == 0:
        return (
            False,
            "PDF conversion failed.\n\n"
            + "\n\n".join(errors[:3])
        )

    if errors:
        return (
            True,
            f"{converted} PDFs created.\nSome files failed:\n"
            + "\n\n".join(errors[:3])
        )

    return True, f"Successfully created {converted} PDF(s)."

def generate_certificates(
    df: pd.DataFrame,
    template_path: Path,
    output_folder: Path,
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

    for count, (index, row) in enumerate(df.iterrows(), start=1):
        grade = str(row["Grades"]).strip().upper()

        if grade == "F":
            failed += 1
            failed_students.append(row.to_dict())
            if status_box:
                status_box.info(f"Skipped: {row['Name of Student']} because Grade is F")
        else:
            prs = Presentation(str(template_path))

            date_issued = row.get("Date Issued") or calculate_date_issued(row["Exam Cycle"])

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
                status_box.success(f"Generated: {filename} | Date Issued: {date_issued}")

        if progress_bar:
            progress_bar.progress(count / total)

    return passed, failed, failed_students


def files_to_bytes_dict(folder: Path, pattern: str) -> dict[str, bytes]:
    """Read generated files into memory so downloads still work after temp folder is deleted."""
    file_dict = {}
    for file in sorted(folder.glob(pattern)):
        if file.is_file():
            file_dict[file.name] = file.read_bytes()
    return file_dict


def download_state_key(key_prefix: str, file_name: str) -> str:
    digest = hashlib.md5(file_name.encode("utf-8")).hexdigest()
    return f"downloaded_{key_prefix}_{digest}"


def mark_downloaded(state_key: str) -> None:
    st.session_state[state_key] = True


def render_file_name(file_name: str, downloaded: bool) -> None:
    if downloaded:
        st.markdown(
            f"<div style='color:#168038; font-weight:700;'>✅ {file_name}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.write(file_name)


def render_individual_downloads(title: str, files_dict: dict[str, bytes], mime: str, key_prefix: str):
    if not files_dict:
        return

    expander_key = f"expander_open_{key_prefix}"

    # Streamlit does not expose native expander open/close state.
    # This toggle keeps the section open across reruns until the user turns it off.
    st.toggle(f"Show {title}", key=expander_key)

    if not st.session_state.get(expander_key, False):
        return

    with st.expander(title, expanded=True):
        search = st.text_input("🔍 Search by name or enrollment number", key=f"search_{key_prefix}")
        filtered_files = {
            name: data
            for name, data in files_dict.items()
            if search.lower() in name.lower()
        }

        st.caption(f"Showing {len(filtered_files)} of {len(files_dict)} files")

        if not filtered_files:
            st.warning("No matching certificate found.")
            return

        for i, (file_name, file_data) in enumerate(filtered_files.items(), start=1):
            state_key = download_state_key(key_prefix, file_name)
            downloaded = st.session_state.get(state_key, False)

            c1, c2 = st.columns([4, 1])
            with c1:
                render_file_name(file_name, downloaded)
            with c2:
                st.download_button(
                    label="✅ Downloaded" if downloaded else "Download",
                    data=file_data,
                    file_name=file_name,
                    mime=mime,
                    key=f"{key_prefix}_{i}_{file_name}",
                    on_click=mark_downloaded,
                    args=(state_key,),
                )


def init_session_state():
    defaults = {
        "generated": False,
        "ppt_zip_bytes": None,
        "pdf_zip_bytes": None,
        "skipped_excel_bytes": None,
        "ppt_files": {},
        "pdf_files": {},
        "summary": {},
        "pdf_message": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def clear_downloaded_statuses():
    for key in list(st.session_state.keys()):
        if str(key).startswith("downloaded_"):
            del st.session_state[key]


def clear_previous_results():
    st.session_state.generated = False
    st.session_state.ppt_zip_bytes = None
    st.session_state.pdf_zip_bytes = None
    st.session_state.skipped_excel_bytes = None
    st.session_state.ppt_files = {}
    st.session_state.pdf_files = {}
    st.session_state.summary = {}
    st.session_state.pdf_message = ""
    clear_downloaded_statuses()


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(
    page_title="Certificate Generator",
    page_icon="🎓",
    layout="wide",
)

init_session_state()
install_custom_fonts()

st.title("🎓 Excel to PPT Certificate Generator")
st.caption("Upload an Excel file and a PowerPoint certificate template to generate personalized certificates.")

with st.sidebar:
    st.header("⚙️ Settings")
    convert_pdf = st.checkbox("Also convert PPT certificates to PDF", value=False)

    st.info(
        "Date Issued is now calculated automatically from the Excel column `Exam Cycle`. "
        "Example: Jan26 → 31/01/2026. Future/current cycles use the last completed month."
    )

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

        missing_columns = [col for col in REQUIRED_COLUMNS if col not in preview_df.columns]
        if missing_columns:
            st.error("Missing required columns: " + ", ".join(missing_columns))
            st.subheader("📊 Excel Preview")
            st.dataframe(preview_df.head(10), use_container_width=True)
        else:
            try:
                preview_df = add_date_issued_preview(preview_df)
            except Exception as e:
                st.error(f"Could not calculate Date Issued from Exam Cycle: {e}")

            st.subheader("📊 Excel Preview")
            st.dataframe(preview_df.head(10), use_container_width=True)

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
    clear_previous_results()

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
            try:
                df = add_date_issued_preview(df)
            except Exception as e:
                st.error(f"Cannot generate. Invalid Exam Cycle data: {e}")
                st.stop()

            st.subheader("🔄 Generation Progress")
            progress_bar = st.progress(0)
            status_box = st.empty()

            passed, failed, failed_students = generate_certificates(
                df=df,
                template_path=template_path,
                output_folder=output_folder,
                progress_bar=progress_bar,
                status_box=status_box,
            )

            # Store individual PPT files in memory
            st.session_state.ppt_files = files_to_bytes_dict(output_folder, "*.pptx")

            # Create and store PPT ZIP in memory
            create_zip(output_folder, ppt_zip)
            st.session_state.ppt_zip_bytes = ppt_zip.read_bytes()

            # Store skipped students Excel in memory
            if failed_students:
                pd.DataFrame(failed_students).to_excel(skipped_file, index=False)
                st.session_state.skipped_excel_bytes = skipped_file.read_bytes()

            # Optional PDF conversion
            if convert_pdf:
                with st.spinner("Converting PPT files to PDF..."):
                    success, message = convert_pptx_to_pdf(output_folder, pdf_folder)
                st.session_state.pdf_message = message

                if success and any(pdf_folder.glob("*.pdf")):
                    st.session_state.pdf_files = files_to_bytes_dict(pdf_folder, "*.pdf")
                    create_zip(pdf_folder, pdf_zip)
                    st.session_state.pdf_zip_bytes = pdf_zip.read_bytes()
                else:
                    st.session_state.pdf_files = {}
                    st.session_state.pdf_zip_bytes = None
            else:
                st.session_state.pdf_message = "PDF conversion was not selected."

            st.session_state.summary = {
                "total": len(df),
                "ppt_created": passed,
                "skipped": failed,
                "pdf_created": len(st.session_state.pdf_files),
            }
            st.session_state.generated = True

# -----------------------------
# Download Center
# -----------------------------
if st.session_state.generated:
    st.success("Certificate generation completed.")

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Total Students", st.session_state.summary.get("total", 0))
    r2.metric("PPT Certificates", st.session_state.summary.get("ppt_created", 0))
    r3.metric("PDF Certificates", st.session_state.summary.get("pdf_created", 0))
    r4.metric("Skipped", st.session_state.summary.get("skipped", 0))

    st.markdown("---")
    st.header("📥 Download Center")

    st.subheader("Option 1: Bulk ZIP Downloads")
    z1, z2, z3 = st.columns(3)

    if st.session_state.ppt_zip_bytes:
        z1.download_button(
            label="⬇️ Download PPT Certificates ZIP",
            data=st.session_state.ppt_zip_bytes,
            file_name="Certificates.zip",
            mime="application/zip",
            key="download_ppt_zip",
        )

    if st.session_state.pdf_zip_bytes:
        z2.download_button(
            label="⬇️ Download PDF Certificates ZIP",
            data=st.session_state.pdf_zip_bytes,
            file_name="PDF_Certificates.zip",
            mime="application/zip",
            key="download_pdf_zip",
        )
    elif convert_pdf:
        z2.warning(st.session_state.pdf_message or "PDF ZIP not available.")

    if st.session_state.skipped_excel_bytes:
        z3.download_button(
            label="⬇️ Download Skipped Students Excel",
            data=st.session_state.skipped_excel_bytes,
            file_name="Skipped_Students.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_skipped_excel",
        )

    st.markdown("---")

    render_individual_downloads(
        title="Option 2: Download Individual PPT Certificates",
        files_dict=st.session_state.ppt_files,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key_prefix="ppt_individual",
    )

    render_individual_downloads(
        title="Option 3: Download Individual PDF Certificates",
        files_dict=st.session_state.pdf_files,
        mime="application/pdf",
        key_prefix="pdf_individual",
    )

    if not st.session_state.pdf_files:
        st.info("Individual PDF downloads will appear only when PDF conversion is selected and successful.")

st.markdown("---")
st.info(
    "Note: For PDF conversion on Streamlit Cloud, add `libreoffice` in packages.txt. "
    "PPT generation works with Python packages only."
)
