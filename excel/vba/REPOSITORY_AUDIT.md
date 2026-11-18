# RackScan source audit — 2026-11-18

## Scope

Reviewed every tracked file on the `main` branch of `finknc01/Random-Code` at
the time of the audit: **29 files**, all under `excel/` and all related to the
RackScan Excel/VBA application. The repository contains no `.xlsm` workbook,
so VBA has not been compiled or exercised in the target Excel application.

## Changes made

1. **Safer display/report text.** Equipment hostnames, serials and other
   identifiers are now written to generated display/report cells using Excel
   text formatting. Strings beginning with a formula character remain literal
   identifiers rather than being interpreted as formulas.
2. **Safer Edit Server rollback.** The edit routine now remembers the original
   values, number formats and any formulas for the six editable columns. If a
   write fails, it attempts to restore all of them and reenables worksheet
   events. This is best-effort recovery; Excel has no database transaction.
3. **Explicit audit success result.** The audit procedure returns a Boolean
   success value. The form no longer decides success by inspecting the text
   of a message, so changing the message wording will not break control flow.
4. **Eight-button documentation.** Updated the top-level overview, Add Server
   guide, ActiveX guide and test plan to consistently describe all eight
   buttons. Added the edge cases to the regression checklist.
5. **Repository navigation.** Added a root README linking source, installation,
   formula, testing and this audit.

## Review findings

| Severity | Finding | Current handling / next step |
| --- | --- | --- |
| High (deployment) | No committed workbook or Windows Excel/VBA compiler for end-to-end tests | Run **Debug > Compile VBAProject** and the full test plan in a backup `.xlsm` before production deployment |
| Medium | Edit/Decom use search-field combinations to find a record, not a permanent Server ID | Multiple matches are rejected safely. A stable Server ID column would simplify edits, moves and history, but requires a planned workbook migration |
| Medium | Manual edits directly to `tblInventory` can bypass Add/Edit collision validation | Rack elevation marks overlaps; consider input validation/protection or enforce edits through approved forms after workflow review |
| Medium | Table lookups and rack rendering scan inventory repeatedly | Current bounded scans cap inventory at 50,000 rows. Benchmark with realistic fictitious row counts in Excel before replacing the logic with cached arrays or an index |
| Medium | Existing decommission archives may use a different layout | `TryDecomServer` requires a compatible `tblDecom` and fails without deleting the active row if the schema differs. Map any existing archive explicitly before use |
| Low | Each new rack audit regenerates the same report worksheet | Intended for the latest audit. Export or copy a report if a historical audit trail is required |
| Environment | ActiveX may be disabled by managed Office security policy | Use only organization-approved controls; do not bypass administrative security settings |

## Source checks

- All tracked files were inventoried; the entire repository is one RackScan
  Excel/VBA codebase.
- The eight ActiveX Click procedures are present in the Rack Lookup sheet code.
- Cross-module calls were checked against the defined procedures/constants.
- All relative Markdown links were checked against the repository tree.
- A simple scan found no embedded passwords, API keys, private-key blocks or
  external HTTP URLs in the committed source text; this is not a full security
  assessment of Git history, workbook data or organizational infrastructure.
- The supplied Excel lookup formula and shared rack U convention agree that
  **U Position is the bottom occupied U** of multi-unit equipment.

## Recommended next iteration

Once the source has compiled and been tested on the working `.xlsm`, prioritize:

1. Add a stable per-server identifier after reviewing how to backfill legacy
   records without invalidating existing scans or archives.
2. Add tests with a realistic-sized **fictional** inventory to measure rack
   refresh time; optimize the renderer only if that benchmark identifies a
   meaningful bottleneck.
3. Define the desired retention and approval rules for physical audit logs,
   decommission history and direct table edits.

Do not upload real server inventory to a public repository for these tests.
