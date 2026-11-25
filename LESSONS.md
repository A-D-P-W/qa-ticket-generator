# Lessons

What went wrong while building the generator, and what fixed it. Condensed from the project's original log of 20
lessons; related ones are merged. Each date is the day the lesson was written down. Lessons 1 to 3 were recorded on
2025-09-19 (the original log gave 2024-05-21, which is before the project began on 2025-09-17). Function names use the
showcase's `QAGEN_` prefix.

## YouTrack API

**1. Read data in the shape it was built (2025-09-19).** A required field was never added to the payload, and the
server answered `400 Field required`. The schema map stored `{ isRequired: true }` objects, but the code that read it
tested the map value as if it were the flag. When one module builds a structure and another reads it, check the
reader against the exact shape the builder makes. A related slip: sub-tasks were sent the raw expected-result text in
place of the formatted description, and the server's error (`entity with the specified name was not found`) pointed
nowhere near the cause.

**2. Pass data in memory, not through the sheet (2025-09-19).** The schema cache was written to a sheet and read back
in the same run, and the read sometimes came back empty. We did not prove why. The fix that held:
`QAGEN_updateYouTrackSchemaCache` returns the map it built, and issue creation takes that map as an argument, so
nothing re-reads the sheet.

**3. A field can be valid in three different ways (2025-09-19).** The admin API gives a field's name and allowed
values. The project's workflow can make a field required without the admin API saying so. And the issue API wants a
different `$type` (`MultiEnumIssueCustomField`) from the one the admin API reports (`EnumProjectCustomField`). The fix
is `REQUIRED_FIELD_DEFAULTS`: a short, human-kept list of workflow-required fields, filled with a value the schema
says is allowed, and sent with the issue API's type.

**4. The response body is the evidence; the docs are the spec (2025-09-19, 2025-09-22).** Guessing at `400` errors
went nowhere. Logging the full response body solved each one, for example by showing which ID field the API really
returns. For how to format a request, the API's own documentation is the authority: YouTrack's Markdown reference says
it follows CommonMark, which confirmed the newline fix in lesson 8.

**5. Link issues through `/api/commands`, with the exact keyword (2025-09-19).** Linking a sub-task with a `POST` to
`/api/issues/{id}/links` returned `405 Method Not Allowed` from both directions: a 405 is about the endpoint and verb,
not the data. `/api/commands` with a query string worked, but `add link Subtask DEMO-2` returned
`field name or field value expected`. The command parser accepts only its own keywords. What works is the query
`add subtask of <parent>`, sent with the child issue.

**6. Keep both IDs (2025-09-19).** Every issue has a short ID (`DEMO-101`) used in URLs and commands, and a permanent
internal ID (`0-101`) that never changes. The script asks for both when it creates an issue, writes both to the sheet,
and uses the permanent ID for status sync.

## Google Sheets and Apps Script

**7. Let a formula sheet do the filtering (2025-09-20).** Hand-maintained checklist rows for one menu did not scale. A
helper sheet with one `FILTER` formula pulls the matching rows from the raw text dump, and a `Checklist (from Sheet)`
blueprint points at it. The formula failed at first with `FILTER has mismatched range sizes`, because the condition
tested one cell while the data was a whole column. The condition range must be the same size as the data range.

**8. Sheet text is not API text (2025-09-21).** Descriptions arrived in YouTrack with `\n` shown as text. A sheet cell
holds a backslash and an `n`, not a newline, so the script converts them (`replaceAll('\\n', '\n')`) before sending.
Two more slips from the same day: `range.getValues()` always returns a 2D array, so a header row is `getValues()[0]`;
and a lowercase `testcase` where `testCase` was meant crashed the run. Small slips give the most confusing errors.

**9. A data source sheet needs its header row (2025-09-20).** A `Checklist (from Sheet)` job failed with
`Data_Source_Column "Text_Content" not found`, because the formula sheet had data but no header row. The error named
the sheet and the column, and the run carried on with the other jobs. Any sheet used as a data source must start with
the header the job names.

**10. One invalid cell can veto a whole write (2025-09-22).** A ticket was created, but its IDs were never written
back. A batch `setValues()` is all or nothing: one cell in the batch that broke a dropdown rule rejected the whole
write, so the sheet and the tracker disagreed. The current code writes the result cells one at a time, and the
dropdowns come from `QAGEN_applyDataValidationRules`, so the sheet rejects a bad value when it is typed.

**11. Get sheets by ID, not from the active context (2025-09-28).** Write-back failed now and then with
`Sheet ... not found`, depending on whether the run started from the editor or the menu. Core code no longer uses
`getActiveSpreadsheet()`: it opens the sheet with `openById`, or takes it from the range it was given
(`dataRange.getSheet()`).

## The test plan as data

**12. The blueprint library is the taxonomy (2025-09-27).** After the `Test_Type` names were changed in the job sheet
and the dropdown lists, diagnostics reported "no matching rule" for every job. The library still used the old names.
Diagnostics was right: a rename must be made in the library, the job sheet and the dropdown lists together.

**13. A visible gap beats a silent one (2025-10-08).** Tickets were created with `[UNDEFINED_PARAMETER:Level_Name]` in
their text, because the jobs used generic blueprints and left `Parameters` empty. The marker is deliberate: it shows a
person exactly which value is missing. A job that uses a generic blueprint must fill its `Parameters`.

**14. Check with the same logic you run (2025-10-08).** Diagnostics passed a job whose live ticket then showed a gap.
Diagnostics had learned to find values in the job's own columns; the generator still looked only in `Parameters`.
Checking and doing must share one source of truth. In this copy they share one function, `_buildParameterMap`.

**15. Run diagnostics until it is clean (2025-10-09).** After a large plan was entered, the live run listed many
missing parameters that diagnostics had not, because diagnostics had been run on an earlier, incomplete version of the
plan. The rule since then: enter the whole plan, run diagnostics, fix what its summary lists, and repeat until it
reports `Found 0 issue(s)`. Only then run generation.

## The way of working

**16. A person approves each hand-off (October 2025; undated in the original).** The plan was produced in three
steps: a test strategist turned design documents into test directives, a data architect turned the directives into
library and job rows, and the generator turned rows into tickets. Each role was an AI assistant with its own
instructions, and a separate coding assistant wrote the generator. A fully automatic chain
was brittle, so a person reviews and approves each step's output before the next step starts, and one role owns the
library and dropdown lists so they are complete before jobs are written.
