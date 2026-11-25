# qa-ticket-generator

A Google Apps Script that turns a test plan kept in Google Sheets into YouTrack tickets. QA writes reusable test
"blueprints" once, lists the jobs to run in a second sheet, and the script fills in each blueprint, checks the job is
ready, creates the tickets (with sub-tasks for long checklists) and writes the ticket IDs back to the sheet. A read-only
diagnostics pass checks the whole plan first and shows a preview, so a bad row is found before anything is created.

I built it in September and October 2025 for a video game QA project. This repository is a de-branded copy: the
studio, the game, the tracker address and every row of test data are gone. `sample-data/` holds small invented sheets
that show the pattern, and `tests/` runs the real script code against them with a fake YouTrack.

## The pattern

Three sheets drive everything:

| Sheet | What a row is | Key columns |
|---|---|---|
| `_Master_Test_Library` | A blueprint: the text of one kind of test, with `${placeholders}` | `Test_Type`, `Category`, `Sub_Category`, `Description`, `Steps`, `Expected_Result` |
| `_generator_config` | A job: "make tickets from this blueprint, with these values" | `Job_ID`, `Generate?`, the same three key columns, `Parameters`, `Depends_On`, `Campaign`, `Trigger_Type`, `Trigger_Value`, then the write-back columns |
| `_Global_Parameters` | A value every job can use, such as the build number | `Key`, `Value` |

A job finds its blueprint by the three-part key (`Test_Type`, `Category`, `Sub_Category`). Placeholders are filled from
the global values, then the job's own columns, then its `Parameters` cell (`Level_Name:Level 1; Language:French`), each
one overriding the one before. A placeholder with no value is not left blank: it becomes `[UNDEFINED_PARAMETER:name]`,
so the gap shows up in the preview.

There are three kinds of blueprint:

| `Test_Type` | Produces | Use it for |
|---|---|---|
| `Single Task` (also `Exploratory Suite`) | One ticket from one blueprint row | A smoke test, a playthrough, one feature check |
| `Checklist (Manual)` | One checklist ticket; every library row with the same key is one item | A short list of checks written by hand |
| `Checklist (from Sheet)` | One checklist item per row of another sheet (`Data_Source_Sheet`, `Data_Source_Column`) | Every UI string, every item in a list that changes |

A checklist of more than 25 items becomes a parent ticket with sub-tasks of up to 25 items each. Each sub-task is
linked to the parent through YouTrack's `/api/commands` endpoint (`add subtask of DEMO-103`, sent on the child).

## What a live run does

For each job, in sheet order, the run stops at the first gate the job fails:

1. `Generate?` is `TRUE` and the row has no ticket ID yet (so a second run never makes a duplicate).
2. Every job named in `Depends_On` has the status `Verified`, `Complete` or `Fixed` in the sheet.
3. A blueprint matches the job's three-part key.
4. Every placeholder in the blueprint has a value.

A job that passes all four gets its ticket(s). The script writes the short ID (`DEMO-101`), the permanent internal ID
(`0-101`), a link and the status `Open` back to the job's own row. The short ID is for people and for commands; the
permanent ID is what status sync uses, because it never changes.

Before creating anything, the run reads the project's custom fields and their allowed values from YouTrack. It keeps a
copy in the `_youtrack_schema_cache` sheet and rewrites that copy only when an MD5 hash of the schema changes. The
schema in memory is passed straight to issue creation, which uses it to fill a field the project's workflow requires
(here `Platforms`) with an allowed value.

## Diagnostics: the dry run

`QAGEN_runDiagnostics` reads the sheets and nothing else: it makes no HTTP request and writes no cell (a test checks
both). For every job marked `TRUE` or `TEST` it reports a missing blueprint or a missing placeholder value, then ends
with an error summary and a count. For a job marked `TEST` it also logs the summary and description of the ticket it
would create. Diagnostics, the live run's placeholder gate and the text filling all build the job's values with the same
function, `_buildParameterMap`, so a job that passes diagnostics passes the same gate in the live run.

On the sample data it finds the two jobs that are broken on purpose:

```
❌ (Row 5 - JobID 'LVL-002'): Missing required parameter(s). The blueprint requires: [Level_Name].
❌ (Row 9 - JobID 'BOSS-001'): A matching rule was not found for the combination: Type="Single Task", Category="Gameplay", Sub_Category="Boss Fight".
```

The intended loop is: run diagnostics, fix what the summary lists, repeat until it reports `Found 0 issue(s)`, and only
then run generation.

## Other entry points

All are on the spreadsheet menu (`QA Ticket Generator`), created by `onOpen`.

| Function | What it does |
|---|---|
| `QAGEN_runTestGeneration` | The live run above, for every enabled job |
| `QAGEN_runCampaign` / `QAGEN_executeCampaign` | A dialog lists the names in the `Campaign` column (a job can be in several, comma-separated); the run takes only that campaign's jobs that have no ticket yet |
| `QAGEN_evaluateTriggers` | For a nightly time trigger: runs enabled jobs whose `Trigger_Type` is due (`On Date`, `Weekly` on a day name, `Recurring` every N days) and stamps `Last_Run` |
| `QAGEN_syncYouTrackStatuses` | Reads each ticket's `State` by permanent ID and updates the sheet's status column when it changed |
| `QAGEN_applyDataValidationRules` | Turns the key columns into dropdowns sourced from a `_data_ref` sheet |
| `QAGEN_setYouTrackToken` | Prompts for the API token and stores it in Script Properties, never in code |

## Run the tests

Needs Node 20 or later. There are no dependencies to install.

```bash
npm test
```

The 12 tests load the files in `src/` into a Node `vm` context with in-memory stand-ins for `SpreadsheetApp`,
`PropertiesService`, `UrlFetchApp` and `Utilities`. The spreadsheet comes from `sample-data/*.tsv`; YouTrack is a fake
that answers the four calls the script makes and records every request.

| Test | What it shows |
|---|---|
| diagnostics is read-only | Exactly the two broken jobs are reported; no HTTP request is attempted, no cell or property is written |
| diagnostics previews a TEST job | The preview shows the filled-in summary and nothing is created |
| parameter precedence | Global, then column, then `Parameters` cell; a gap becomes `[UNDEFINED_PARAMETER:…]` |
| live run | Which jobs are created and which gate stops the rest; sub-task links; 25 + 5 checklist split; write-back cells |
| payload | Schema-driven `Platforms` default, Markdown flag, real newlines, token from Script Properties, schema cache and hash |
| no duplicates, dependency opens | A second run creates nothing; a blocked job runs once its dependency is `Verified` |
| campaign | Only the campaign's jobs run |
| scheduled triggers | On a fixed Monday the `On Date` and `Weekly` jobs fire; the `Recurring` limitation below is pinned |
| status sync | A changed `State` is copied into the sheet |
| data validation | Dropdowns are applied to the four columns `_data_ref` covers |
| hygiene (2 tests) | No token, key or long Google ID in any file; `.clasp.json` is ignored |

The tests were checked by breaking the code on purpose: raising the checklist size to 30 fails the live-run and
campaign tests, and adding a YouTrack call to diagnostics fails the read-only test.

## Deploy it (optional)

1. Make a Google Sheet with the tabs in `sample-data/` (import each TSV) plus an empty `_youtrack_schema_cache` tab.
2. Copy `.clasp.json.example` to `.clasp.json` and put in your script ID, then set the three values in `src/Config.js`.
3. `clasp push`, reload the sheet, and use **Configuration > Set YouTrack Token**.
4. Run diagnostics until it is clean, then run generation. Generation creates real tickets.

## Known limitations

These are in the code as it was used; the showcase keeps them rather than changing behavior.

- A `Recurring` job never fires until it has a `Last_Run` date, and a scheduled job that has fired once has a ticket ID,
  so the trigger run skips it from then on. In practice scheduled jobs fire once.
- `Depends_On` reads the status column in the sheet, so run status sync first. The run reads the sheet once at the
  start, so a dependency created in the same run does not count until the next run.
- A campaign run does not check `Generate?`, and does not stamp `Last_Run`.
- Diagnostics and the placeholder gate check the blueprint text only, and for a `Checklist (Manual)` only its first
  row. A placeholder in `Task_Title_Override` is filled but not checked, so a gap there shows up in the ticket title.
- Diagnostics does not check values against `_data_ref` or that a `Data_Source_Sheet` exists; the live run reports
  those.
- Write-back is one cell at a time, and there is no retry or back-off on API errors (each failure is logged with the
  server's response body).

## What changed from the original

- Every employer name, game term, URL, project ID, sheet ID and script ID is removed; `src/Config.js` holds placeholders.
- The function prefix is now `QAGEN_` and the menu is `QA Ticket Generator`.
- Diagnostics and generation now call one `_buildParameterMap`; the original had two copies of the same logic.
- `sample-data/`, `tests/` and this README are new. The original had no automated tests: every check was a live run
  that created tickets.

## How it was built

I designed the system and directed AI chat assistants to write most of the code, in set roles (one wrote code, one
kept the sheet data model). I ran every change against a live YouTrack project, read the logs, and decided what to fix.
[LESSONS.md](LESSONS.md) is the record of what went wrong and what fixed it. This copy, its tests and its docs were made
with Claude Code.
