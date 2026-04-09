# 8_Pipeline.py
import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

PROJECT_DIR = Path(__file__).resolve().parent

STEP1_SCRIPT = PROJECT_DIR / "1_Step_load_Crop.py"

SCRIPTS = [
    PROJECT_DIR / "2_Step_Extract_Raw.py",
    PROJECT_DIR / "3_Step_Cleaning_Comparison.py",
    PROJECT_DIR / "4_Step_Mismatch_Report.py",
    PROJECT_DIR / "5_Step_automated_Email.py",
]

LOG_FILE = PROJECT_DIR / "pipeline_log.txt"
STEP_LOG_DIR = PROJECT_DIR / "logs"
STEP_LOG_DIR.mkdir(parents=True, exist_ok=True)


def log(msg: str):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run_cmd(cmd, cwd: Path, step_name: str):
    log(f"RUN ({step_name}): {' '.join(map(str, cmd))}")

    env = dict(**os.environ)
    env["PYTHONUTF8"] = "1"  # helps encoding in Windows console/logs

    result = subprocess.run(
        cmd,
        cwd=str(cwd),
        text=True,
        capture_output=True,
        env=env
    )

    # Write full step log
    step_log_path = STEP_LOG_DIR / f"{step_name}.log"
    with open(step_log_path, "w", encoding="utf-8") as f:
        f.write("CMD:\n" + " ".join(map(str, cmd)) + "\n\n")
        f.write("RETURN CODE:\n" + str(result.returncode) + "\n\n")
        f.write("STDOUT:\n" + (result.stdout or "") + "\n\n")
        f.write("STDERR:\n" + (result.stderr or "") + "\n")

    # Keep pipeline_log smaller (only first 4000 chars of each)
    if result.stdout:
        log(f"[{step_name}] STDOUT (first 4000 chars):\n{result.stdout[:4000]}")
    if result.stderr:
        log(f"[{step_name}] STDERR (first 4000 chars):\n{result.stderr[:4000]}")

    if result.returncode != 0:
        raise RuntimeError(
            f"{step_name} failed (exit {result.returncode}). "
            f"Open full log: {step_log_path}"
        )


def run_script(py_path: Path):
    if not py_path.exists():
        raise FileNotFoundError(f"Script not found: {py_path}")
    cmd = [sys.executable, str(py_path)]
    run_cmd(cmd, cwd=PROJECT_DIR, step_name=py_path.stem)


def main():
    log("=== PIPELINE START ===")

    log("Step 1/7: Running 1_Step_load_data.py")
    run_script(STEP1_SCRIPT)

    for i, script in enumerate(SCRIPTS, start=2):
        log(f"Step {i}/7: Running {script.name}")
        run_script(script)

    log("=== PIPELINE DONE (SUCCESS) ===")


if __name__ == "__main__":
    import os  # needed for env in run_cmd
    try:
        main()
    except Exception as e:
        log(f"❌ PIPELINE FAILED: {e}")
        raise
