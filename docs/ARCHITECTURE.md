# Architecture

## Overview

`takecourses` is a local-only tool made of three pieces:

- `server.py` — a stdlib + `openpyxl`-only HTTP server bound to
  `127.0.0.1:8765`. It serves `index.html`, exposes a small JSON API, and
  does all the `.xlsx` extraction.
- `run.bat` — a double-click launcher that runs `py server.py`.
- `index.html` — the browser UI: vanilla JS, no dependencies, a
  hand-rolled SVG bar chart (no charting library, no CDN).

The server opens Chrome to `http://127.0.0.1:8765/` on startup.

## Project layout

```
takecourses/
  server.py        - local HTTP server + xlsx extraction logic (stdlib + openpyxl only)
  run.bat           - double-click launcher (runs `py server.py`)
  index.html        - the UI (vanilla JS, no dependencies, hand-rolled SVG bar chart)
  tests/
    test_server.py  - regression test suite (subprocess + real HTTP + synthetic openpyxl fixtures)
  README.md
  docs/
    ARCHITECTURE.md
  *.xlsx            - the user's degree-plan files (added/removed freely; auto-detected)
```

## Domain finding: how course data is extracted from the `.xlsx` files

Each source `.xlsx` has a sheet literally named `סיכום` (Hebrew: "Summary").
That sheet contains six rectangular blocks, one per semester, each block
tagged by a bare single Hebrew letter from `{א, ב, ג, ד, ה, ו}` (semesters
1–6, Israeli academic convention: א/ב = year 1 fall/spring, ג/ד = year 2,
ה/ו = year 3).

Each course is one row within a block, laid out as 4 cells left-to-right on
that row:

```
course_code, course_name, credits, semester_letter
```

Concrete real example from the shipped example file's sheet (row 5 of
`סיכום`):

| Cell | Value | Meaning |
|------|-------|---------|
| `B5` | `77101` | course code |
| `C5` | `'מכניקה ויחסות פרטית'` | course name |
| `D5` | `7` | credits |
| `E5` | `'א'` | semester letter |

So, relative to the letter-cell, the name is 2 cells to its left, the code
is 3 cells to its left, and the credits value is 1 cell to its left.

### Why extraction scans for the letter instead of hardcoding columns

Block *position* on the sheet is **not** fixed. In the example file,
semesters א/ג/ה occupy one row-band of column-groups and ב/ד/ו occupy
another. The *number of courses per block* also varies file to file (some
semesters have 3 courses, others 7+).

What **is** constant across all files following this template is the
4-cell left-to-right relative pattern on any given course row. So
extraction works by scanning every cell on the sheet for a value that is
exactly one of the bare semester letters (`א`, `ב`, `ג`, `ד`, `ה`, `ו`),
and then reading that cell's 3 leftward neighbors (code, name, credits) —
rather than by hardcoding fixed column letters like `B`/`C`/`D`/`E`.

### Header/total rows are skipped automatically

Header and total rows within the same blocks also carry a semester-letter
tag in that same column position. But their "name" position lands on a
cell that's part of a merged header range, which reads as empty/`None`
when accessed via `openpyxl`. Extraction naturally skips these rows by
requiring the name cell to be a non-empty string — no special-casing of
header/total rows is needed.

### Fallback sheet detection

If a file's `סיכום` sheet is missing or misnamed, the backend falls back to
scanning all sheets for one containing a cell with the literal text
`תצוגה לפי סמסטר` ("View by semester" — the block section header) and uses
that sheet instead.

### Sheets that are intentionally NOT used

Other sheets in the workbook (e.g. `התואר שלי`, `מדמ"ח`, `מתמטיקה`,
`פיסיקה`, and their `תואר שני -` counterparts) hold per-track detail views
with a similar-but-not-identical column layout. They are not read by this
tool — only `סיכום` (or its fallback) is read, because it's the one sheet
that consolidates every track into a single canonical per-semester view.

## No caching, by design

Every API request re-reads and re-parses the relevant `.xlsx` files fresh
from disk. There is no caching anywhere in the server. This is intentional:
the source files may be edited between requests (e.g. a user editing a
plan in Excel while the tool is open in the browser), and the tool is
meant to always reflect the latest content on disk without requiring a
restart.

## API contract

### `GET /api/files`

Returns a JSON array of `.xlsx` basenames present in the folder. Excel
lock files (names starting with `~$`) are excluded.

### `GET /api/histogram?semester=<letter>&files=<name1>&files=<name2>...`

`semester` is one of `א`, `ב`, `ג`, `ד`, `ה`, `ו`. `files` is repeated once
per selected filename.

Response:

```json
{
  "semester": "א",
  "results": [{"code": 77101, "name": "מכניקה ויחסות פרטית", "count": 3}],
  "errors": [{"file": "somefile.xlsx", "message": "why it failed"}]
}
```

- `results` is sorted by `count` descending — courses present in more of
  the selected files rank higher. `count` is the number of selected files
  that contain that course in the given semester, deduped per file (a
  course appearing twice in one file's block still counts once for that
  file).
- `errors` lists any selected files that failed to parse (missing file,
  missing/unreadable sheet, etc.), each with the filename and a short
  reason. A file's failure does not prevent the other selected files from
  contributing to `results`.
- An invalid or missing `semester` parameter returns HTTP 400.

This reference is meant to be complete enough that the backend or frontend
can be modified without re-deriving the sheet layout or API shape from
scratch.

## Concurrency

`server.py` uses `http.server.ThreadingHTTPServer` (not the plain
single-threaded `HTTPServer`), so the frontend's 3-second `/api/files` poll
and a user-triggered `/api/histogram` request can be served concurrently
instead of queueing behind each other. There's no shared mutable state
between requests (every request opens and closes its own `openpyxl`
workbook), so this is safe as-is.

## Known characterized edge case: duplicate filenames in one request

`_handle_histogram` does not dedupe the `files` query parameter. If the
same filename appears twice in one `/api/histogram` request, that file's
courses are counted twice. This is exercised and documented explicitly by
`test_14_duplicate_filename_in_query_characterization` in
`tests/test_server.py` — it's a characterization test (asserts the actual
current behavior), not a claim that this is desired behavior. The shipped
`index.html` UI cannot produce this today (each checkbox contributes at
most one `files=` value per distinct filename), so it's not a live bug, but
if the frontend or any other client is changed to build the query
differently, dedupe the list server-side first.

## Frontend gotcha: SVG `text-anchor` under an RTL page

The page root is `<html dir="rtl">`. CSS `direction` is an inherited
property, and modern browsers resolve SVG `text-anchor: start`/`end`
*relative to the element's `direction`* (a logical mapping), not to a fixed
physical left/right. Left un-handled, that flips which physical side
`text-anchor: end` anchors to, which — combined with the `<svg>` root's own
default `overflow: hidden` — silently clips course-name labels off the
right edge of the chart (this was hit and fixed during this project's
initial hardening pass; the symptom was "course names disappear on the
right"). The fix, in `renderChart()` in `index.html`: the generated `<svg>`
explicitly sets `direction: ltr` (both as a presentation attribute and an
inline style), which makes `text-anchor` unambiguous/physical again. This
does **not** affect how the Hebrew text itself renders — glyph
shaping/ordering is governed by the Unicode Bidi Algorithm on the
character data, not by the container's `direction` — it only fixes the
anchor-point math. If the chart is ever rewritten or a new SVG text element
is added, keep (or re-apply) this `direction: ltr` on the SVG root.

## `localStorage` persistence

Two keys, set/read in `index.html`'s inline script:

- `selectedSemester` — the last-chosen semester letter, restored on load
  (defaults to `'א'` if absent/invalid).
- `selectedFiles` — a JSON array of filenames that were checked, restored
  on load. `saveSelectedFiles()` is called from `refreshHistogramNow()`, so
  it's kept up to date on every checkbox toggle, select-all/clear-all, and
  poll-driven removal.

Gotcha to preserve if this logic is touched again: restoring checkbox
`checked` state alone (in `addFileRow`/`loadFilesInitial`/`pollFiles`) is
not sufficient to make the restored selection actually show a chart —
`loadFilesInitial()` must explicitly call `refreshHistogramNow()` after
populating rows if any ended up checked, otherwise the page loads with
boxes visibly checked but the chart still stuck on the "select at least one
file" empty state until the user manually touches something. This was a
real bug caught (by tracing, not live browser testing — see Testing below)
and fixed during this project's initial hardening pass; the fix is the
`if (getCheckedFiles().length > 0) { refreshHistogramNow(); }` call at the
end of `loadFilesInitial()`'s success branch.

## Testing

`tests/test_server.py` is a permanent, dependency-free (stdlib `unittest` +
`openpyxl`) regression suite. Design: it launches a real, port-patched copy
of `server.py` as a subprocess and drives it over actual HTTP (never
imports/calls server internals directly), using synthetic `.xlsx` fixtures
it builds with `openpyxl` at runtime in a temp directory — nothing is
committed as binary fixture files, and the real project `.xlsx` file(s) are
only ever read (GET requests), never modified. Run with `py
tests\test_server.py`; a clean run prints `SUMMARY: N/N passed` and exits
non-zero on any failure.

Coverage includes: extraction correctness against the real example file
(cross-checked against a raw cell-by-cell dump of its `סיכום` sheet — see
the docstring on `test_01` for the verified expected course lists per
semester), zero-course blocks, the fallback sheet-detection path, missing
summary sheet entirely, in-file duplicate courses, falsy/zero course codes,
merged-cell header rows, corrupted files, multi-file aggregation, `~$`
lock-file exclusion, invalid/missing `semester`, nonexistent requested
files, path-traversal attempts, the duplicate-filename characterization
(see above), unknown routes, Hebrew `Content-Length` byte-vs-character
correctness, and concurrent-request handling.

Re-run this suite after any change to `server.py`'s extraction logic or API
shape — it's the fastest way to confirm nothing regressed without manually
re-deriving expected values from the sheet again.

Note on scope: the automated suite covers the backend/API only. There is no
automated frontend/browser test — `index.html` has been verified by close
reading plus manual HTTP-level checks that mirror exactly what the UI's
`fetch()` calls send (see "Frontend gotcha" and "`localStorage`
persistence" above for two real bugs caught this way), but an actual
in-browser walkthrough (clicking checkboxes, watching the chart) has not
been performed by an automated tool in this project yet — do a manual
check in Chrome after any `index.html` change, especially to the chart
rendering.

## Possible future improvements (not applied yet)

Identified during this project's initial hardening pass but intentionally
left out of that round — listed here so a future session doesn't need to
rediscover them:

- **Surface credit points.** `extract_courses_for_semester` reads each
  course's credit value off the row but currently discards it before
  returning — only `code`/`name` make it into the API response. Adding
  `"credits": ...` to each result and showing it in the chart
  tooltip/label would give more planning context at low implementation
  risk.
- **Friendlier port-in-use error.** If port 8765 is already bound (e.g. a
  previous instance of the tool didn't exit cleanly), `main()` currently
  lets the raw `OSError` propagate as an unhandled traceback. Catching it
  and printing a clear "already running? check
  http://127.0.0.1:8765" message would be a small UX improvement.
- Lighter-weight ideas, not yet scoped: a search/filter box for the file
  list if the folder ever holds many plans; visually distinguishing
  courses common to *all* selected files vs. only some; an "export chart
  as image" action.
