# takecourses

A light local web tool for comparing academic degree plans. Run one Python
script and it starts a tiny local-only HTTP server on `127.0.0.1:8765` and
opens Chrome to it automatically. The page lets you look at the `.xlsx`
degree-plan files sitting in this same folder (each file represents one
person's academic plan, e.g. `מבנה תואר - ישראל ישראלי.xlsx`), pick a subset
of them via checkboxes, pick exactly one of the six semesters, and see a
live histogram of how many of the selected files include each course in
that semester. Every request re-reads the underlying `.xlsx` files fresh
from disk — nothing is cached, so editing a file and reloading always
reflects the latest content.

## Requirements

- Python 3 with `openpyxl` installed. On this machine that's already covered
  by the `py` launcher (Windows, Python 3.9.13).
- No other dependencies, and no internet connection is needed at any point —
  the chart is a hand-rolled SVG bar chart, not a CDN library.

## How to run

Either:

- Double-click `run.bat`, or
- Run `py server.py` from this folder in a terminal.

Either way, it opens Chrome automatically at `http://127.0.0.1:8765/`.

To stop it, close the terminal window, or press Ctrl+C in it.

## How to use

1. Drop `.xlsx` plan files (following the template — see
   `docs/ARCHITECTURE.md` for the exact structure) into this folder. The
   file list in the browser picks them up automatically within about 3
   seconds — no restart needed.
2. Check the ones you want to compare.
3. Pick a semester.
4. Watch the chart update live.

Both your file selection and semester choice are remembered in the browser
(`localStorage`) and restored automatically the next time you open the page
— you don't need to re-pick your usual comparison set every time.

## Testing

`tests/test_server.py` is a self-contained regression suite (stdlib
`unittest` + `openpyxl`, no other dependencies) that starts a real copy of
`server.py` as a subprocess and exercises it over actual HTTP, using
synthetic `.xlsx` fixtures it builds and tears down itself — it never
touches your real plan files. Run it after making any change to
`server.py`:

```
py tests\test_server.py
```

A clean run ends with `SUMMARY: N/N passed`. See `docs/ARCHITECTURE.md` for
what it covers and known/expected edge-case behavior it documents.

## Known limitations

- Only the `סיכום` sheet is read (or, as a fallback, a sheet containing the
  text `תצוגה לפי סמסטר`).
- Files must follow the template's `code, name, credits, semester-letter`
  row pattern.
- This is a local-only tool — it binds to `127.0.0.1` and is not meant to be
  exposed on a network.
- Credit-point values are read from each course row but not currently
  surfaced anywhere (API response or UI) — see "Possible future
  improvements" in `docs/ARCHITECTURE.md` if you want that added later.

## Project layout

```
takecourses/
  server.py        - local HTTP server + xlsx extraction logic (stdlib + openpyxl only)
  run.bat           - double-click launcher (runs `py server.py`)
  index.html        - the UI (vanilla JS, no dependencies, hand-rolled SVG bar chart)
  tests/
    test_server.py  - regression test suite, run with `py tests\test_server.py`
  README.md
  docs/
    ARCHITECTURE.md
  *.xlsx            - the user's degree-plan files (added/removed freely; auto-detected)
```

See `docs/ARCHITECTURE.md` for how the `.xlsx` extraction and the API work
under the hood.
