# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file Playwright automation (`mlc_auto_match.py`, ~1100 lines) that drives
[The MLC Matching Tool](https://portal.themlc.com) in a real browser to match tracks
from an Excel/CSV catalog against MLC's database, in place of doing it by hand.

## Setup

```
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
npx playwright install chromium
```

Credentials via env vars only, never hardcoded:
```
export MLC_EMAIL="..."
export MLC_PASSWORD="..."
```
(Windows PowerShell: `$env:MLC_EMAIL="..."`)

## Running

```
python mlc_auto_match.py catalogo.xlsx --sheet "FEB 25" --output risultati.xlsx
```

Work the master catalog directly, in place (colors rows + writes status columns):
```
python mlc_auto_match.py "mlc spreadsheet.xlsx" --colora-catalogo
```

Work it in chunks (each run picks up where the last one left off):
```
python mlc_auto_match.py "mlc spreadsheet.xlsx" --colora-catalogo --limite 50
```

Key flags: `--headless` (skip on first run of any change — see below), `--skip-ambigui`
(don't stop for manual review, leave `ambiguous_work` rows for later),
`--rifai-tutto` (with `--colora-catalogo`, reprocess already-colored rows).

There is no test suite, linter, or build step in this repo — validation happens by
running against the real MLC portal (see "Verifying changes" below).

## Architecture: two sequential search stages per row

MLC's matching tool is not a single search; each catalog row goes through two
independent searches against the same page:

1. **Stage 1 — find the recording** (`stage1_search_isrc`): search by ISRC. Multiple
   groups under the same ISRC (e.g. DSP-variant duplicates) are normal — all get
   selected and confirmed together via `_select_all_groups`, clicking "Load More"
   first if present. No result → `no_match_recording`. Already Submitted/Accepted/
   Rejected → `already_submitted`.
2. **Stage 2 — find the matching work** (`stage2_search_work`): search by Title +
   Publisher Name fixed to `"LOO"` (matches "LOOSE CLUB EDITION" regardless of the
   exact publisher in the row). If that's empty or ambiguous and a writer surname is
   available, retries with Title + Writer Name (`_set_second_criteria`).

`process_row` drives both stages for one row and returns a `RowResult` with one of
the statuses documented in the README (`matched`, `no_match_recording`,
`already_submitted`, `no_match_work`, `submit_failed`, `ambiguous_recording`,
`ambiguous_work`, `manual_incomplete`, `error`).

MLC's title search matches on individual words, not phrases, so results routinely
include unrelated titles. `_scegli_opera`/`_norm_title` normalize and compare titles
(accents, case, punctuation) to decide automatically between: single exact-title
match (auto-select), multiple exact matches that are the same work under different
Song Codes (auto-select first), no exact match (`no_match_work`), or multiple
different homonymous works (queued for manual selection — `ambiguous_work`).

Ambiguous rows don't block the batch: `main()` runs straight through, queues them,
and only prompts for manual selection in the browser at the end (unless
`--skip-ambigui`).

## `--colora-catalogo` mode specifics

`carica_catalogo`/`_aggiorna_catalogo` treat the master spreadsheet itself as the
only file to consult: writes `MLC Stato`, `MLC Note`, `MLC Aggiornato` columns and
colors each row (green/yellow/orange, same scheme as the pre-existing manual
process). Automatic backup to `backup/` before any write. Resumable: rows with a
definitive outcome (colored, including pre-existing hand-colored ones) are skipped;
`submit_failed`/`error`/`manual_incomplete` rows are retried since those are
technical failures, not real outcomes. Same ISRC across multiple rows (co-writers)
is processed once and all its rows are colored together. Writes (report and
catalog) go through a temp file + `os.replace` (`_salva_atomico`) so they're atomic
against crashes/Ctrl+C — this matters because the report/catalog is rewritten after
*every row*, not just at the end. Ctrl+C is handled (`_installa_stop_pulito`) to
finish the current row cleanly before stopping; a second Ctrl+C exits immediately.

## Working on the Playwright automation itself

The selectors and waits encode a long list of verified-live quirks of the MLC portal
(see README "Assunzioni da verificare al primo run" for the full list with evidence).
The important ones to not accidentally regress:

- `search.1.searchTerm` is the **same test-id reused** for ISRC (Stage 1) and for the
  Publisher/Writer value (Stage 2). If a Stage 2 criteria change is interrupted, this
  slot can stay dirty and silently produce false `no_match_recording` results on
  later rows. There are two guards for this — don't remove them without replacing
  them: `run_search` verifies Stage 1's row 1 actually reads "Recording ISRC" before
  searching (full reload if not), and error recovery does a full page reload, not a
  sidebar click (sidebar click is a SPA no-op).
- Group selection buttons disappear once clicked, shifting indices — always click
  the first *remaining* one in a loop, never `nth(i)` for a fixed `i`.
- Never add fixed-time sleeps after "Search"; wait for the actual network response
  (`/search/unmatched-recordings` for Stage 1, `/search/works/catalog` for Stage 2),
  then for render.
- The single-result case in work extraction needs its own test path (DOM ancestor
  walk behaves differently with one result vs. two-plus) — this was a real bug that
  silently turned single-result matches into `no_match_work`.

### Verifying changes

There's no automated test suite — this is a live-portal automation. Run with
`--headless` **off** against the real account after any change, and specifically
re-check the single-result Stage 2 case and the multi-group Stage 1 case, since
those are the cases past bugs hid in. Never let the MLC password appear in chat,
screenshots, or logs.
