

## Checkpoint 2026-10-04 — DEV-45 test result

DEV-45 was tested and produced the same updater error: Azuriranje nije izvrseno / The parameter is incorrect.

This proves the DEV-37 source reset did not eliminate the error. The screenshot does not identify the failing statement. Next step is to inspect PCD-Excel-checkforupdate.log and PCD-Excel-updater.log from this exact test before changing architecture.


## Checkpoint 2026-10-04 — DEV-46 diagnostic instrumentation

DEV-45 failed again with the same error. The two logs established that CheckForUpdate started, but the old updater log was stale, so ScheduleReplacement/VBS was not proven to have been reached.

A single targeted change was made: DEV-46 adds checkpoints to CheckForUpdate before/after manifest HTTP, version comparison, download, SHA-256, and ScheduleReplacement. No updater architecture was changed.

Commits:
- 0c28e0d22cacdc978c6072e32be04255f80fd7a2 — add precise execution checkpoints
- 02a3efde53946877f5b79477f31984147c9ba168 — bump DEV build to 46

Next: wait for CI/publication, then use DEV-46 to identify the exact failing boundary. Do not make another architectural change before that evidence exists.


## Checkpoint 2026-10-04 — DEV-46 test result

User reports that DEV-46 still displays the same error dialog: Azuriranje nije izvrseno / The parameter is incorrect.

The screenshot does not identify the failing checkpoint. The next required evidence is the new PCD-Excel-checkforupdate.log from this DEV-46 run. PCD-Excel-updater.log is also useful to confirm whether ScheduleReplacement launched the detached script.

No architecture change should be made until those logs are inspected.


## Checkpoint 2026-10-04 — Root-cause candidate isolated at manifest HTTP

DEV-46 log stopped immediately after BEFORE MANIFEST HTTP; there was no AFTER MANIFEST HTTP. This localizes the error to HttpGetText / MSXML HTTP setup or request, before download, SHA-256, or VBS replacement.

A targeted fix was applied: cacheBust changed from CStr(Timer) to CStr(CLng(Timer * 1000)), removing a locale-dependent decimal separator from the URL query parameter on systems using a comma decimal separator.

Commits:
- 28ca3f908bfd0c81319ee802c9efa279ffaffd4f — fix locale-sensitive cache-buster URL
- e2578a1c0db2437fbc61e85c3272db5595221c80 — bump DEV build to 47

Next: verify CI/publication, then test DEV-47. If manifest HTTP passes, the diagnostic checkpoints will tell us the next boundary. Do not change any other part of the updater before that result.
