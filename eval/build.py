#!/usr/bin/env python3
"""Build the eval corpus, and recalculate it with LibreOffice.

    models/mNN_*.py  ──►  build/<name>.xlsx        formulas, no cached values
                                │                  (exactly what openpyxl — and
                                │                   therefore an agent — emits)
                                ▼
                     soffice --headless --convert-to xlsx
                                │
                                ▼
                          build/recalc/<name>.xlsx  every value filled in

The pair is the whole experiment. The first file is the artefact under test: a
workbook whose every formula cell has an empty `<v>`, which is why it renders as
a skeleton in any viewer that only reads the cache. The second is the oracle.

Usage:
    python3 eval/build.py            build everything
    python3 eval/build.py m03 m10    build only the models whose name matches
    python3 eval/build.py --no-recalc
"""

from __future__ import annotations

import importlib
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
RECALC = BUILD / "recalc"

sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "models"))

MODELS = [
    "m01_budget",
    "m02_valuation_dcf",
    "m03_lbo",
    "m04_three_statement",
    "m05_workflow_approvals",
    "m06_sales_data",
    "m07_cohort_retention",
    "m08_loan_amortization",
    "m09_inventory_planning",
    "m10_edge_cases",
]

SOFFICE_CANDIDATES = [
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
]


def find_soffice() -> str | None:
    for candidate in SOFFICE_CANDIDATES:
        resolved = shutil.which(candidate) or (candidate if Path(candidate).exists() else None)
        if resolved:
            return resolved
    return None


def count_formulas(wb) -> tuple[int, int, int]:
    formulas = cells = sheets = 0
    for ws in wb.worksheets:
        sheets += 1
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                cells += 1
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas += 1
    return sheets, cells, formulas


def recalc(soffice: str, src: Path) -> tuple[Path | None, float]:
    RECALC.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    try:
        subprocess.run(
            [soffice, "--headless", "--norestore", "--convert-to",
             "xlsx:Calc MS Excel 2007 XML", "--outdir", str(RECALC), str(src)],
            check=True, capture_output=True, timeout=600,
        )
    except subprocess.CalledProcessError as exc:
        print(f"      LibreOffice failed: {exc.stderr.decode()[:300]}")
        return None, time.time() - t0
    except subprocess.TimeoutExpired:
        print("      LibreOffice timed out")
        return None, time.time() - t0
    out = RECALC / src.name
    return (out if out.exists() else None), time.time() - t0


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    do_recalc = "--no-recalc" not in sys.argv
    selected = [m for m in MODELS if not args or any(a in m for a in args)]
    if not selected:
        sys.exit(f"no model matches {args}; known models: {', '.join(MODELS)}")

    from common import namespace_workbook

    BUILD.mkdir(parents=True, exist_ok=True)
    soffice = find_soffice() if do_recalc else None
    if do_recalc and not soffice:
        print("LibreOffice not found — writing fixtures without an oracle.\n")

    total_formulas = 0
    for name in selected:
        module = importlib.import_module(name)
        t0 = time.time()
        wb = module.build()
        namespace_workbook(wb)
        sheets, cells, formulas = count_formulas(wb)
        dest = BUILD / f"{name}.xlsx"
        wb.save(dest)
        gen = time.time() - t0
        total_formulas += formulas

        line = (f"  {name:26s} {sheets} sheets  {cells:6,d} cells  "
                f"{formulas:6,d} formulas  {gen:5.1f}s")
        if soffice:
            out, secs = recalc(soffice, dest)
            line += f"  → recalc {secs:5.1f}s" if out else "  → RECALC FAILED"
        print(line)

    print(f"\n{len(selected)} workbooks · {total_formulas:,} formulas → {BUILD}")
    if soffice:
        print(f"oracle → {RECALC}")
    print("\nNow run:  npm run eval")


if __name__ == "__main__":
    main()
