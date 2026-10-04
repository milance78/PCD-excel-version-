# DEBUG_STATE.md

> DO NOT START DEBUGGING FROM SCRATCH. READ THIS FILE FIRST.
>
> Persistent checkpoint for the PCD Excel updater debugging. Update after every meaningful debugging step, build, test, or verified result.

## Goal

The Excel button Proveri ažuriranje must detect a newer XLSM, download it, verify SHA-256, replace/update the SharePoint-hosted workbook, close the old Excel instance, and reopen the updated workbook.

The workbook can be opened directly from SharePoint, so ThisWorkbook.FullName can be a SharePoint URL rather than a local Windows path.

## Current state

- Repository: milance78/PCD-excel-version-
- Branch: main
- Current development target: DEV-45
- Updater source was deliberately reset to the exact DEV-37 updater implementation before bumping the build to DEV-45.
- DEV-45 build workflow was pending at the last checkpoint; do not tell the user to test DEV-45 until CI succeeds and VERSION.json is verified as DEV-45.
- Last relevant commits:
  - 6d1dd0d4aaf3b8010126a044aa2689c2059d8b63 — restore exact proven DEV-37 updater implementation
  - cb1896ecec48f946a9f5e275d2e5cafd9d516004 — bump DEV build to 45

## What is proven

### Version detection / manifest

The updater uses the GitHub API for VERSION.json and Base64-decodes the returned content. DEV-33 diagnostics proved local version/build, remote version/build, and comparison all worked.

Therefore basic manifest/version comparison is not the primary suspect.

### DEV-37 update path

DEV-37 successfully reached the updater path and automatically reopened the SharePoint workbook with the new version. The updater log showed START, NEW=temp XLSM, OLD=SharePoint URL, STARTING EXCEL COM, OPENING NEW XLSM, SAVING TO SHAREPOINT URL, SAVEAS OK, OPENING SHAREPOINT COPY, REOPEN OK, END.

This is the strongest evidence that the VBS + Excel COM SharePoint replacement architecture can work.

### VBA compile

DEV-35 and subsequent builds were manually opened and compiled successfully at the relevant stages. DEV-43 later exposed a compile error caused by a bad caller reference introduced during an architecture change; that was fixed in DEV-44.

## What failed

### Windows MOVE against SharePoint URL

Old updater attempted Windows MOVE using a SharePoint HTTPS URL. Logs repeatedly showed MOVE ERROR 52 Bad file name or number.

Do not return to the Windows MOVE architecture.

### VBS top-level Exit Sub

A generated VBS script once contained a top-level Exit Sub, causing Microsoft VBScript runtime error: Invalid 'exit' statement. It was replaced with WScript.Quit 0.

### ShellExecute / launcher experiments

ShellExecute was introduced to replace WScript.Shell.Run, with diagnostic logging around the launch. DEV-40 still failed with Ažuriranje nije izvršeno / The parameter is incorrect. The updater log remained stale at 22:23, suggesting the new launch path was not reaching the VBS script.

Do not assume ShellExecute solved the launch problem.

### Direct Excel COM from VBA

A direct COM replacement was attempted in DEV-42. It was conceptually wrong while the original Excel process/workbook was still active and ultimately produced The parameter is incorrect.

Do not revert to direct in-process COM replacement as the current architecture.

### DEV-43 bad caller

During restoration of the VBS architecture, the caller temporarily referenced ReplaceThroughExcelCom even though that routine had been removed. This caused Compile error: Sub or Function not defined. It was fixed to call ScheduleReplacement tempPath, ThisWorkbook.FullName, CurrentExcelProcessId. DEV-44 contained that fix.

## Important diagnostic fact

The user reported that the updater log timestamp stayed at 22:23 even when clicking update around 22:56.

This means the click was not reaching the VBS script that creates/recreates PCD-Excel-updater.log, or the relevant script was not being launched. The launcher boundary therefore became a major suspect.

## Proven source baseline

The current DEV-45 source was intentionally reset to the exact updater source from commit 0c0c339d3b0971e7faf125b7c2bec3fa4c426f46. This is the DEV-37 source that had successfully completed the SharePoint replacement/reopen path.

Its ScheduleReplacement uses the simple CreateObject WScript.Shell and shell.Run wscript.exe architecture.

Do not make another architectural change before verifying this baseline.

## Known-good / known-bad timeline

- DEV-31 — stale updater/cache investigation; old and new workbook confusion occurred.
- DEV-32 — build published; version state investigated.
- DEV-33 — diagnostics proved manifest/version comparison.
- DEV-34 — update prompt worked, but generated VBS had invalid top-level Exit Sub.
- DEV-35 — VBS exit fixed; workbook showed DEV-35 and could check latest.
- DEV-36 — further version/update testing.
- DEV-37 — important success: update path completed and reopened SharePoint workbook; later user saw a VBS The parameter is incorrect message in another update attempt.
- DEV-38 — shutdown error handling changed; did not resolve the failure.
- DEV-40 — ShellExecute launcher experiment; still failed with The parameter is incorrect.
- DEV-41 — diagnostic launcher logging; still failed.
- DEV-42 — direct Excel COM architecture; failed and was abandoned.
- DEV-43 — VBS architecture restored, but caller compile error introduced.
- DEV-44 — compile caller fixed; user still received The parameter is incorrect.
- DEV-45 — updater source reset exactly to DEV-37 baseline; pending CI verification at checkpoint.

## Things NOT to retry blindly

1. Windows MOVE/local filesystem replacement against a SharePoint URL.
2. Direct Excel COM replacement while the current workbook/Excel process is still active.
3. Random launcher architecture changes without first proving where execution stops.
4. Repeated DEV builds without checking CI result and published VERSION.json.
5. Asking the user to test a build before confirming that the build actually contains the intended VBA source.

## Required workflow for future steps

After every meaningful change:
1. Make one targeted change.
2. Commit it.
3. Verify GitHub Actions completed successfully.
4. Verify published VERSION.json contains the expected version and SHA.
5. Only then ask the user for a test.
6. Record the test result here immediately.
7. Record what the result proves or rules out.
8. Do not repeat an already disproven architecture.

## Immediate next step

1. Wait for DEV-45 GitHub Actions run to complete.
2. Verify VERSION.json is actually DEV-45 and contains the expected SHA.
3. Only then decide how to test the restored DEV-37 baseline.
4. If the baseline still fails, instrument one exact execution boundary at a time rather than changing the whole architecture again.

## User constraint

The user is exhausted by repeated trial-and-error. The objective is to finish the updater, not generate endless DEV versions. Prefer evidence from CI, source comparison, logs, and deterministic inspection over asking the user to repeatedly test speculative builds.

## Checkpoint 2026-10-04 — DEV-45 publication verified

- `VERSION.json` on `main` now reports `0.1.0-dev-45`.
- Published artifact SHA-256 in `VERSION.json`: `a0c168d4d706b32234c34ec8a11feaca7f6abfb040105f3a271ffd8ee854a037`.
- `build/build_xlsm.py` on `main` is `VERSION = "0.1.0-dev-45"`.
- Current `vba/PCD_Updater.bas` on `main` has blob SHA `d7c3facfd39f6b9ae8101c54ca440c3ad074db6e` and uses the GitHub API manifest plus the restored updater structure.
- Workflow definition currently validates `ScheduleReplacement` before building the XLSM.
- The GitHub commit-workflow-runs lookup for the DEV-45 version commit returned no run record, so the workflow completion itself has not been independently verified through that endpoint. However, the published `VERSION.json` has advanced to DEV-45, which confirms that the publish step occurred.

### Decision

DEV-45 is published and source state is consistent with the intended reset. Do not make another code change yet. The next useful action is a single controlled test of DEV-45, with the result recorded here immediately. If it fails, capture the exact new evidence before changing architecture.
