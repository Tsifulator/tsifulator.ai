# tsifl regression harness

Purpose: lock in known-good behavior so future prompt/backend changes can't silently regress projects that already worked.

## Architecture

```
tests/regression/
├── cases/                  # one directory per test case
│   └── 001-placerhills-09/
│       ├── request.json    # POST payload for /chat
│       ├── images/         # optional screenshots attached to the request
│       ├── start.xlsx      # workbook state before the request (for reference)
│       ├── expected.xlsx   # gold-standard output (for reference/inspection)
│       └── rubric.yaml     # assertions the returned actions must satisfy
├── rubrics/                # per-app rubric evaluators
│   ├── excel.py
│   ├── word.py             # (future)
│   ├── powerpoint.py       # (future)
│   └── rstudio.py          # (future)
├── run_tests.py            # main runner
├── capture_case.py         # snapshot a live request → new test case
└── requirements.txt
```

## What gets asserted

Rubrics are **not** exact-match. LLM output varies between runs — brittle equality checks would flap. Instead the rubric asserts:

- **Required action types** — e.g. "must include at least one `create_named_range`"
- **Required cell references** — e.g. "some action must write to `Calculator!C18`"
- **Formula content patterns** — e.g. "`Sales Forecast!F5` formula must contain `E5:E26/C5:C26` and `*2`"
- **Forbidden patterns** — e.g. "no action may target sheets outside `[Calculator, Price Solver, Sales Forecast]`"
- **Non-destructive guarantees** — e.g. "no action may overwrite `Price Solver!B12:B17` labels"

This catches the bugs we've actually seen — phantom sheets, wrong formulas, destructive label edits — without being flaky.

## Running

```bash
cd tests/regression
pip install -r requirements.txt

# run all cases against deployed Railway backend
python run_tests.py

# run a single case
python run_tests.py --case 001-placerhills-09

# run against local backend
python run_tests.py --backend http://localhost:8000
```

Exit code `0` = all pass. Non-zero = at least one case failed; report is printed.

## Adding new cases

After a successful live run in any tsifl add-in:

```bash
python capture_case.py --name 002-comps-table
# Walks through: request payload, attached images, expected workbook,
# then generates a starter rubric.yaml you can tune.
```

## CI

A GitHub Action runs this suite on every push to `main`. A regression blocks the deploy.
