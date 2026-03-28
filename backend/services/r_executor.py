"""
R Executor Service
Runs R code from Claude's actions and returns the output.
Used when the request comes from Excel and needs R processing,
or to validate R code before sending to the RStudio panel.
"""

import subprocess
import tempfile
import os


def run_r_code(code: str, timeout: int = 30) -> dict:
    """
    Executes R code in a subprocess and returns stdout/stderr.
    Returns: {"success": bool, "output": str, "error": str}
    """
    # Write code to a temp file
    with tempfile.NamedTemporaryFile(mode="w", suffix=".R", delete=False) as f:
        f.write(code)
        tmp_path = f.name

    try:
        result = subprocess.run(
            ["Rscript", tmp_path],
            capture_output=True,
            text=True,
            timeout=timeout
        )
        return {
            "success": result.returncode == 0,
            "output": result.stdout.strip(),
            "error": result.stderr.strip() if result.returncode != 0 else ""
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "output": "", "error": "R script timed out after 30s"}
    except FileNotFoundError:
        return {"success": False, "output": "", "error": "R not found. Is R installed?"}
    finally:
        os.unlink(tmp_path)
