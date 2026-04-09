import os
import sys
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st


# ============================================================
# CONFIG
# ============================================================
APP_DIR = Path(__file__).resolve().parent
DEFAULT_PROJECT_DIR = APP_DIR

STEP_FILES = {
    "Step 1 - Load PDF and crop": "1_Step_load_Crop.py",
    "Step 2 - Extract raw": "2_Step_Extract_Raw.py",
    "Step 3 - Clean and compare": "3_Step_Cleaning_Comparison.py",
    "Step 4 - Mismatch report": "4_Step_Mismatch_Report.py",
    "Step 5 - Automated email draft": "5_Step_automated_Email.py",
    "Run full pipeline": "6_Pipeline.py",
}

OUTPUT_FILES = {
    "Raw data": "Data_1_RawData.xlsx",
    "Clean data": "Data_2_Clean.xlsx",
    "Mismatch report": "Data_3_Mismatch_Report.xlsx",
}

INPUT_PDF_DIRNAME = "1_NyaRitningar"
OUTPUT_IMAGE_DIRNAME = "2_output_images"
CROP_DIRNAME = "3_cropsImages_flat"
LOG_DIRNAME = "logs"
PIPELINE_LOG = "pipeline_log.txt"


# ============================================================
# PAGE SETUP
# ============================================================
st.set_page_config(page_title="Checkdat Stampel AI", layout="wide")
st.title("Checkdat Stampel AI")
st.caption("Run your existing step scripts from a Streamlit interface.")


# ============================================================
# HELPERS
# ============================================================
def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def project_path(project_dir: Path, name: str) -> Path:
    return project_dir / name


def ensure_folder(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def read_text_file(path: Path) -> str:
    if not path.exists():
        return ""
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except Exception as e:
        return f"Could not read file: {e}"


def save_uploaded_pdfs(files, target_dir: Path, clear_first: bool) -> list[str]:
    ensure_folder(target_dir)

    saved = []
    if clear_first:
        for item in target_dir.glob("**/*"):
            if item.is_file():
                try:
                    item.unlink()
                except Exception:
                    pass

    for up in files:
        out_path = target_dir / up.name
        with open(out_path, "wb") as f:
            f.write(up.getbuffer())
        saved.append(up.name)

    return saved


def run_script(script_path: Path, working_dir: Path) -> tuple[bool, str]:
    if not script_path.exists():
        return False, f"Script not found: {script_path}"

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"

    cmd = [sys.executable, str(script_path)]

    result = subprocess.run(
        cmd,
        cwd=str(working_dir),
        text=True,
        capture_output=True,
        env=env,
    )

    log_text = []
    log_text.append(f"[{now_str()}] RUN: {' '.join(cmd)}")
    log_text.append(f"RETURN CODE: {result.returncode}")

    if result.stdout:
        log_text.append("\nSTDOUT:\n" + result.stdout)
    if result.stderr:
        log_text.append("\nSTDERR:\n" + result.stderr)

    ok = result.returncode == 0
    return ok, "\n".join(log_text)


def output_file_exists(project_dir: Path, filename: str) -> bool:
    return (project_dir / filename).exists()


def dataframe_preview_from_excel(path: Path) -> pd.DataFrame | None:
    if not path.exists():
        return None
    try:
        xls = pd.ExcelFile(path)
        first_sheet = xls.sheet_names[0]
        return pd.read_excel(path, sheet_name=first_sheet)
    except Exception:
        return None


def render_download_button(path: Path, label: str):
    if not path.exists():
        return
    with open(path, "rb") as f:
        st.download_button(
            label=label,
            data=f.read(),
            file_name=path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def folder_metrics(project_dir: Path) -> dict:
    pdf_dir = project_dir / INPUT_PDF_DIRNAME
    img_dir = project_dir / OUTPUT_IMAGE_DIRNAME
    crop_dir = project_dir / CROP_DIRNAME

    pdf_count = len(list(pdf_dir.glob("**/*.pdf"))) if pdf_dir.exists() else 0
    img_count = len([p for p in img_dir.glob("**/*") if p.is_file()]) if img_dir.exists() else 0
    crop_count = len([p for p in crop_dir.glob("**/*") if p.is_file()]) if crop_dir.exists() else 0

    return {
        "pdf_count": pdf_count,
        "img_count": img_count,
        "crop_count": crop_count,
    }


# ============================================================
# SESSION STATE
# ============================================================
if "last_log" not in st.session_state:
    st.session_state.last_log = ""


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("Settings")

    project_dir_input = st.text_input(
        "Project folder",
        value=str(DEFAULT_PROJECT_DIR),
        help="Folder that contains your step scripts and output files.",
    )
    project_dir = Path(project_dir_input)

    st.markdown("---")
    st.subheader("Expected folders")
    st.write(f"Input PDFs: `{INPUT_PDF_DIRNAME}`")
    st.write(f"Output images: `{OUTPUT_IMAGE_DIRNAME}`")
    st.write(f"Crops: `{CROP_DIRNAME}`")
    st.write(f"Logs: `{LOG_DIRNAME}`")

    st.markdown("---")
    st.subheader("Quick status")
    st.write("Project exists:" , project_dir.exists())
    for label, script_name in STEP_FILES.items():
        if label == "Run full pipeline":
            continue
        st.write(f"{script_name}:", (project_dir / script_name).exists())


# ============================================================
# TOP STATUS
# ============================================================
metrics = folder_metrics(project_dir) if project_dir.exists() else {"pdf_count": 0, "img_count": 0, "crop_count": 0}

c1, c2, c3 = st.columns(3)
c1.metric("PDF files", metrics["pdf_count"])
c2.metric("Rendered pages", metrics["img_count"])
c3.metric("Cropped stamps", metrics["crop_count"])


# ============================================================
# PDF UPLOAD SECTION
# ============================================================
st.subheader("1. Upload PDFs")

pdf_target_dir = project_dir / INPUT_PDF_DIRNAME
ensure_folder(pdf_target_dir)

uploaded_pdfs = st.file_uploader(
    "Upload one or more PDF drawings",
    type=["pdf"],
    accept_multiple_files=True,
)
clear_before_upload = st.checkbox("Clear existing PDFs before upload", value=False)

col_upload_a, col_upload_b = st.columns([1, 2])
with col_upload_a:
    if st.button("Save uploaded PDFs", use_container_width=True):
        if not uploaded_pdfs:
            st.warning("Please upload at least one PDF first.")
        else:
            saved = save_uploaded_pdfs(uploaded_pdfs, pdf_target_dir, clear_before_upload)
            st.success(f"Saved {len(saved)} PDF file(s) to {pdf_target_dir}")
with col_upload_b:
    if pdf_target_dir.exists():
        current_pdfs = sorted([p.name for p in pdf_target_dir.glob("**/*.pdf")])
        if current_pdfs:
            st.write("Current PDFs in input folder:")
            st.dataframe(pd.DataFrame({"PDF": current_pdfs}), use_container_width=True, hide_index=True)
        else:
            st.info("No PDFs found in the input folder yet.")


# ============================================================
# RUN STEPS
# ============================================================
st.subheader("2. Run steps")

run_cols = st.columns(3)
buttons = [
    ("Step 1 - Load PDF and crop", run_cols[0]),
    ("Step 2 - Extract raw", run_cols[1]),
    ("Step 3 - Clean and compare", run_cols[2]),
    ("Step 4 - Mismatch report", run_cols[0]),
    ("Step 5 - Automated email draft", run_cols[1]),
    ("Run full pipeline", run_cols[2]),
]

for label, col in buttons:
    with col:
        if st.button(label, use_container_width=True):
            script_name = STEP_FILES[label]
            script_path = project_dir / script_name
            ok, log_text = run_script(script_path, project_dir)
            st.session_state.last_log = log_text
            if ok:
                st.success(f"{label} finished successfully.")
            else:
                st.error(f"{label} failed. Check the log below.")


# ============================================================
# LOG VIEWER
# ============================================================
st.subheader("3. Logs")

log_tab1, log_tab2 = st.tabs(["Latest run", "pipeline_log.txt"])

with log_tab1:
    st.text_area("Latest command output", value=st.session_state.last_log, height=320)

with log_tab2:
    pipeline_log_path = project_dir / PIPELINE_LOG
    st.text_area("Project pipeline log", value=read_text_file(pipeline_log_path), height=320)

step_log_dir = project_dir / LOG_DIRNAME
if step_log_dir.exists():
    step_logs = sorted(step_log_dir.glob("*.log"))
    if step_logs:
        selected_log = st.selectbox("Open step log", [p.name for p in step_logs])
        if selected_log:
            st.text_area(
                "Selected step log",
                value=read_text_file(step_log_dir / selected_log),
                height=260,
            )


# ============================================================
# OUTPUTS
# ============================================================
st.subheader("4. Output files")

out_cols = st.columns(3)
for idx, (label, filename) in enumerate(OUTPUT_FILES.items()):
    path = project_dir / filename
    with out_cols[idx]:
        st.markdown(f"**{label}**")
        st.write(path.exists())
        render_download_button(path, f"Download {filename}")

preview_tabs = st.tabs(list(OUTPUT_FILES.keys()))
for tab, (label, filename) in zip(preview_tabs, OUTPUT_FILES.items()):
    with tab:
        path = project_dir / filename
        if not path.exists():
            st.info(f"{filename} does not exist yet.")
            continue

        df_preview = dataframe_preview_from_excel(path)
        if df_preview is None:
            st.warning("Could not preview this file, but you can still download it.")
        else:
            st.dataframe(df_preview, use_container_width=True)
            st.caption(f"Rows: {len(df_preview)} | Columns: {len(df_preview.columns)}")


# ============================================================
# FILE / FOLDER OVERVIEW
# ============================================================
st.subheader("5. Project overview")

overview_col1, overview_col2 = st.columns(2)

with overview_col1:
    st.markdown("**Scripts**")
    script_df = pd.DataFrame(
        {
            "Script": list(STEP_FILES.values()),
            "Exists": [(project_dir / name).exists() for name in STEP_FILES.values()],
        }
    )
    st.dataframe(script_df, use_container_width=True, hide_index=True)

with overview_col2:
    st.markdown("**Outputs**")
    out_df = pd.DataFrame(
        {
            "File": list(OUTPUT_FILES.values()),
            "Exists": [(project_dir / name).exists() for name in OUTPUT_FILES.values()],
        }
    )
    st.dataframe(out_df, use_container_width=True, hide_index=True)


# ============================================================
# FOOTER NOTES
# ============================================================
st.markdown("---")
st.markdown(
    """
    **Notes**

    - This app assumes your existing step scripts stay in the same project folder.
    - Your current scripts still use hard-coded Windows paths. Because those paths point to the same project folder,
      this app will work best when `streamlit_app.py` is placed inside `C:\\Users\\admin\\Checkdat_Stampel_AI`.
    - Step 4 and Step 5 require Microsoft Excel / Outlook on Windows because they use `win32com`.
    - Step 2 requires your current OCR/YOLO dependencies and model file to already be installed and available.
    """
)
