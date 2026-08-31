# CLAUDE.md — RCM–EC Reconciliation: Streamlit Demo Build

| | |
|---|---|
| **Purpose** | Standing orders for Claude Code. Build a Streamlit app that runs the RCM–EC requisition↔position reconciliation engine against real Excel extracts (the two SuccessFactors reports), chained day-over-day. Client-facing demo; synthetic library included. |
| **Parent build** | The delivered solution is the **M365-only flow** (SharePoint + Power Automate + Office Scripts), POC verified 27 Aug 2026. This Streamlit app is the **presentation skin** — it must reproduce that engine's behaviour exactly, never replace it. |
| **Owner** | Jagraj Singh (Genpact) · ANZ Project Union, T&C People Analytics, Seq 27 |
| **Handling** | Internal Genpact. All bundled data is SYNTHETIC. Never imply access to ANZ data. |

---

## 0. Hard rules (non-negotiable — apply to every phase)

1. **No AI anywhere.** The engine is deterministic: parsing, joins, string comparison, regex normalisers. No LLM calls, no ML libraries, no embeddings. If a task seems to want AI, the answer is more config, not a model. (Account constraint: all AI/GenAI initiatives paused pending ANZ's enterprise AI strategy.)
2. **Synthetic-by-default, and provably so.** Library mode shows a permanent `SYNTHETIC DATA — DEMO ONLY` banner. Upload mode shows `UPLOADED DATA — data-handling gate applies`. The app never ships with, caches, or hard-codes anything resembling real ANZ records. Every bundled workbook carries a `_SYNTHETIC` sheet.
3. **Config over code.** Field pairs, weights, and value normalisers are read from `config_field_map.xlsx` at runtime. Only rows with `status = confirmed` are compared. Changing the field list must never require a code change.
4. **Engine parity with the Office Script.** The rules in §4 are a contract, not a starting point. Do not "improve" matching logic (no fuzzy matching, no trimming beyond `canon()`, no extra exception types) — parity is what makes this demo an honest representation of the delivered M365 build.
5. **IDs are text.** Read every cell as a string (`dtype=str`; never let pandas infer numerics). Leading zeros on Position Numbers / cost centres must survive round-trips — wrong-looking IDs erode recruiter trust.
6. **Fail loudly.** Header validation errors halt the run and name the missing columns on screen. No silent partial reconciliation, ever.
7. **Golden tests gate everything.** `pytest` (§7) must pass before any UI work is shown or any claim of "working" is made. Do not weaken a test to make it pass; if a test fails, the engine is wrong.

**Non-goals:** no cloud deployment, no SharePoint/Graph integration, no scheduler, no email sending, no database. Single-machine `streamlit run app.py`.

---

## 1. What you are building

A Streamlit app with two modes:

- **Library mode (default):** scans `synthetic_library/extracts/` for dated folders (`YYYY-MM-DD/`), runs the engine over all days **in date order, chained** (each day's New/Recurring/Resolved computed against the previous day's findings), and lets the presenter walk the days: KPIs, filterable mismatch register, resolved panel, per-recruiter digest previews, config viewer with a live normaliser toggle.
- **Upload mode ("real-time excels"):** the user uploads today's two extracts (requisition report + position report), optionally a previous day's register workbook for chain status, and optionally a config workbook (else the bundled config is used). The engine runs immediately on upload and offers the produced register as a downloadable `.xlsx`. This is the path that goes live when real extracts clear the data-handling gate — same headers, same engine, zero code change.

---

## 2. Repository layout

```
rcm-ec-recon/
  CLAUDE.md                      # this file
  requirements.txt               # streamlit>=1.36, pandas>=2.0, openpyxl>=3.1
  .streamlit/config.toml         # Genpact dark theme (§6)
  app.py                         # Streamlit UI only — no engine logic in this file
  engine/
    __init__.py
    reconcile.py                 # pure engine: no streamlit, no I/O imports
    io_excel.py                  # workbook readers/writers + header validation
  synthetic_library/             # bundled — DO NOT regenerate or edit
    config/config_field_map.xlsx
    extracts/2026-08-20/ … 2026-08-26/   (RCM_Requisition_Report_*.xlsx, EC_Position_Report_*.xlsx)
  registers/                     # engine output per day (gitignored)
  tests/test_golden.py           # §7 acceptance tests
```

`engine/reconcile.py` must be importable and testable with no Streamlit installed.

---

## 3. Data contract — real headers (exact strings, case-sensitive)

**Requisition report (RCM extract).** Required columns:
`Requisition No` · `Position Number` · `Requisition Title (BL)` · `Recruitment Stage` · `Current Status` · `Recruiter (R) Name` · `Hiring Manager Name` · `Job Code` · `Job Code Label` · `Job Grade` · `Cost Center Number` · `Cost Center Name` · `Legal Entity Code` · `Legal Entity Name` · `Country` · `Employee Class` · `FBS Function` · `FBS LoB`

**Position report (EC extract).** Required columns:
`Position Number` · `Position Title` · `Job Code` · `Job Code Label` · `Pay Grade Level` · `Cost Center Code` · `Cost Center Name` · `Legal Entity Code` · `Legal Entity` · `Country` · `Position Type` · `Function Name` · `Line of Business Name` · `Line Manager Name`

Validation rule (mirror the Office Script): the requisition report must contain
`Position Number`, `Requisition No`, `Recruiter (R) Name`, `Recruitment Stage`, `Current Status` **plus every confirmed `req_col`**; the position report must contain `Position Number` **plus every confirmed `pos_col`**. Missing columns → hard error naming them.

**Config workbook `config_field_map.xlsx`:**
- Sheet `FieldMap`, columns `req_col | pos_col | rule | weight | status | notes`. 13 confirmed pairs + 2 pending (`FBS Division ↔ Operational Division Name`, `Location ↔ Location`). Pending rows are displayed greyed-out in the UI and **never compared**.
- Sheet `ValueNormalisers`, columns `report | column | pattern | replacement`. One row ships: `requisition | Country | \s*\([A-Z]{2}\)$ | ""` (strips "Australia (AU)" → "Australia"). Apply each normaliser as a Python `re.sub` to the named column of the named report **before** comparison.
- Extra sheets (`ReadMe`, `_SYNTHETIC`) are ignored by the loader.

Extra/unknown columns in any workbook are ignored. First data sheet of each extract is the report; identify it as the first sheet whose name is not `_SYNTHETIC`.

---

## 4. Engine specification (faithful port of Office Script `RunReconciliation`)

Constants: `OFFER_STAGES = {"offer", "offer approval"}` · `ACTIVE = {"open", "active"}` · severity map `high→HIGH, medium→MEDIUM, low→LOW`, default `MEDIUM`.

`canon(v)`: `str(v or "")`, trim, collapse internal whitespace to single spaces, lowercase. Used for **all** comparisons and key lookups. Display values keep original casing.

Run order:
1. Load both extracts and config. Validate headers (§3). Apply value normalisers.
2. Filter requisitions to `canon(Current Status) ∈ ACTIVE`.
3. Index positions by `canon(Position Number)`; a key seen more than once is marked duplicate.
4. Per active requisition, in this order, first match wins and short-circuits the row:
   - empty key → `MISSING KEY` (weight high)
   - duplicate key → `DUPLICATE POSITION` (weight high, position value literal `(multiple position rows)`)
   - no position row → `ORPHAN REQUISITION (no position row)` (weight high)
   - else compare every confirmed pair: equal after `canon` → skip; one side empty → `MISSING IN POSITION|REQUISITION REPORT`; both non-empty and different → `MISMATCH`. Field label is exactly `"{req_col} <> {pos_col}"`.
5. Severity = weight-mapped value, **escalated to HIGH** when `canon(Recruitment Stage) ∈ OFFER_STAGES`.
6. Chain key = `"{Requisition No}|{field}|{exception type}"`. Status = `Recurring` if the key existed in the previous run, else `New`. `Resolved` = previous keys absent from the current run (report requisition no, field, exception type).
7. Outputs per run: findings list (columns as in the register below), resolved list, per-recruiter digests (count, HIGH count, HTML table identical in shape to the Office Script digest), summary (run date, active count, pairs checked, exceptions, HIGH, resolved).

**Register workbook writer** (download in both modes; mirrors the M365 register):
- `Mismatch_Register`: `Requisition No, Position Number, Requisition Title, Recruitment Stage, Recruiter, Hiring Manager, Exception Type, Field, Requisition Value, Position Value, Severity, Status`
- `Run_Summary`: metric/value rows including the literal row `DATA | SYNTHETIC - demo only. Never imply access to ANZ data.` in library mode (upload mode: `DATA | UPLOADED - data-handling gate applies`)
- `Resolved_This_Run`: `Requisition No, Field, Exception Type, Status` (sheet present only when non-empty)
- All ID columns written with text number-format.

---

## 5. UI specification

Layout top-to-bottom: provenance banner (rule 2) → mode switch → day rail → KPI row → tabs.

- **Day rail (library mode):** one segment per dated folder in chronological order, showing date, exception count, `new · recurring · ✓resolved` mix; selecting a day drives everything below. Include a small caption: *"Each day's register starts as a copy of the previous day's — New/Recurring/Resolved needs no database."*
- **KPI row:** Active requisitions · Field pairs checked · Exceptions (with HIGH sub-count) · New/Recurring · Resolved this run.
- **Tab 1 — Mismatch register:** filters (severity, status, recruiter, free-text search) over a table styled so HIGH severity and Recurring/New are visually distinct; values shown as `req value ⟷ pos value`. Below it, a green "Resolved since previous run" panel. Empty states must be written, not blank.
- **Tab 2 — Recruiter digests:** recruiter list with counts; selecting one renders the engine's digest HTML in a white email-style card with a subject line (`RCM–EC reconciliation: N items need your review (YYYY-MM-DD)`). Caption: digests are what the scheduled M365 flow will email; the loop is not yet wired.
- **Tab 3 — Config & rules:** FieldMap table (pending rows greyed), the normaliser row, and a **normaliser toggle**: switching it off re-runs the day and surfaces the Country false-positive count with the line *"this is why ValueNormalisers exists."*
- **Upload mode:** three file inputs (requisition, position, previous register optional, config optional), run button, same KPI/register/digest rendering for the single run, register download button.
- Footer: *"Deterministic engine — no AI components. Logic mirrors Office Script RunReconciliation (M365-only build, verified 27 Aug 2026)."*

## 6. Theme

`.streamlit/config.toml`:

```toml
[theme]
base = "dark"
primaryColor = "#FFAD28"
backgroundColor = "#181C23"
secondaryBackgroundColor = "#282A27"
textColor = "#FFFFFF"
```

Accent usage: Sunset Orange `#FFAD28` for selection/interactive, Coral `#FF4F59` for HIGH severity and the synthetic banner, `#43C98A` for Resolved. Keep everything else quiet.

---

## 7. Golden acceptance tests (`tests/test_golden.py`) — MUST pass

Run the chained engine over the bundled library (normaliser ON) and assert exactly:

| Run date | Active reqs | Exceptions | New | Recurring | Resolved (IDs) |
|---|---|---|---|---|---|
| 2026-08-20 | 55 | 3 | 3 | 0 | — |
| 2026-08-21 | 56 | 4 | 1 — `101109` ORPHAN | 3 | — |
| 2026-08-24 | 56 | 2 | 0 | 2 | `101103`, `101109` |
| 2026-08-25 | 57 | 3 | 3 — `101118`, `101127`, `101140` | 0 | `101112`, `101131` |
| 2026-08-26 | 58 | 3 | 1 — `101144` | 2 — `101127`, `101140` | `101118` |

Point assertions: `101140` = Hiring Manager mismatch, HIGH, stage Offer. `101127` = FBS Function mismatch, MEDIUM, stage Screening. `101144` = Cost Center Number `45888` vs `45008`, HIGH, stage Offer. With the normaliser OFF, the 2026-08-26 run yields **61** findings, of which **58** are `Country <> Country` false positives. Also test: header validation raises with the missing column named when a required column is dropped; IDs round-trip as text.

Provenance of the goldens: the 2026-08-25 and 2026-08-26 outcomes reproduce the documented, verified results of the M365 POC run (handover §5/§7). The 2026-08-20/21/24 issues are invented synthetic story beats for chain demonstration — consistent with, but not sourced from, the original demo zips.

---

## 8. Build phases (execute in order; stop and report after each)

1. **Engine:** `engine/reconcile.py` + `engine/io_excel.py` + `tests/test_golden.py`. Gate: full pytest green.
2. **Register writer + upload path:** register `.xlsx` output, previous-register chaining from an uploaded workbook. Gate: a register produced from day N, fed as "previous" for day N+1, reproduces the chained statuses.
3. **Streamlit UI:** library mode complete per §5. Gate: manual walkthrough of all five days matches the golden table on screen.
4. **Upload mode + polish:** upload flow, theme, empty states, banner logic. Gate: pytest still green; `streamlit run app.py` clean from a fresh venv.

## 9. Demo script and positioning (for the presenter)

Open on 2026-08-20, walk the days forward, land on 2026-08-26: two Recurring, one New at Offer, and `101118` resolved — the "it notices fixes, not just breaks" moment. Close in the Config tab: flip the normaliser off, show the 58 false positives, flip it back.

Positioning line to say to the client, verbatim in spirit: **"This interface is how we demonstrate the engine. The delivered solution runs entirely inside your M365 tenancy — SharePoint, Power Automate, Office Scripts — with zero new infrastructure and zero AI components."** Never let the Streamlit skin imply a hosted-Python delivery architecture. Internal: complete the ECHR dedupe alignment before demoing widely.
