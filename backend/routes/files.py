"""
Files Route — reads CSV/TSV files from the server filesystem.
Used by the Excel add-in's import_csv action to pull R-generated
data into the workbook.
"""

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
import csv
import os
from pathlib import Path

router = APIRouter()

# Common directories where R/shell scripts save files
SEARCH_DIRS = [
    ".",                          # current working dir (backend/)
    "/tmp",                       # common temp dir
    os.path.expanduser("~"),      # home directory
    "/app",                       # Railway app root
    os.path.expanduser("~/Documents"),
    os.path.expanduser("~/Downloads"),
]

class ReadCSVRequest(BaseModel):
    path: str
    delimiter: str = ","
    max_rows: int = 5000  # safety cap

def _find_file(path: str) -> str | None:
    """Search for a file in common directories if not found at the given path."""
    # 1. Try exact path first
    expanded = os.path.expanduser(path)
    if os.path.isfile(expanded):
        return expanded

    # 2. If it's a relative path, search common directories
    if not os.path.isabs(path):
        for d in SEARCH_DIRS:
            candidate = os.path.join(d, path)
            if os.path.isfile(candidate):
                return candidate

    # 3. Try just the filename in case a full path was given but wrong directory
    basename = os.path.basename(path)
    if basename != path:
        for d in SEARCH_DIRS:
            candidate = os.path.join(d, basename)
            if os.path.isfile(candidate):
                return candidate

    # 4. Fuzzy match: case-insensitive, ignore extension, allow stem-prefix match.
    #    e.g. "loandata" or "loandata.csv" → "LoanData.csv"
    target_stem = os.path.splitext(basename)[0].lower().replace(" ", "").replace("_", "")
    for d in SEARCH_DIRS:
        try:
            for entry in os.listdir(d):
                if not entry.lower().endswith((".csv", ".tsv", ".txt")):
                    continue
                entry_stem = os.path.splitext(entry)[0].lower().replace(" ", "").replace("_", "")
                if entry_stem == target_stem or entry_stem.startswith(target_stem):
                    full = os.path.join(d, entry)
                    if os.path.isfile(full):
                        return full
        except (FileNotFoundError, PermissionError):
            continue

    return None

@router.post("/read-csv")
async def read_csv(request: ReadCSVRequest):
    """Read a CSV file and return its contents as a 2D array."""
    found_path = _find_file(request.path)

    if not found_path:
        # List what we searched for debugging
        searched = [os.path.join(d, os.path.basename(request.path)) for d in SEARCH_DIRS]
        raise HTTPException(
            status_code=404,
            detail=f"File not found: {request.path}. Searched: {', '.join(searched)}"
        )

    try:
        rows = []
        with open(found_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=request.delimiter)
            for i, row in enumerate(reader):
                if i >= request.max_rows:
                    break
                # Convert numeric strings to numbers for Excel
                converted = []
                for cell in row:
                    cell = cell.strip()
                    if cell == "":
                        converted.append("")
                        continue
                    # Try int first, then float
                    try:
                        converted.append(int(cell))
                        continue
                    except ValueError:
                        pass
                    try:
                        converted.append(float(cell))
                        continue
                    except ValueError:
                        pass
                    converted.append(cell)
                rows.append(converted)

        return {
            "data": rows,
            "rows": len(rows),
            "cols": max(len(r) for r in rows) if rows else 0,
            "source": found_path,
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error reading file: {str(e)}")
