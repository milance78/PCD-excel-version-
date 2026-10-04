

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


## Checkpoint 2026-10-04 — DEV-47 still fails

DEV-47 produced the same `The parameter is incorrect` dialog. Therefore the cache-buster locale hypothesis was not sufficient to resolve the problem.

The next diagnostic target remains the manifest HTTP call. The existing checkpoint stops at BEFORE MANIFEST HTTP, so the failure is inside HttpGetText. Instrument the individual HTTP operations (CreateObject, Open, each request header, Send) and log the exact URL before changing behavior. Do not change the updater architecture.


## Checkpoint 2026-10-04 — debugger correction / DEV-49

The DEV-48 screenshot still showed the generic error. Previous checkpoint logging was insufficient because the global UpdateError handler did not record Err.Number, Err.Source, or the exact execution stage.

Correction: DEV-49 adds a stage variable to CheckForUpdate and logs the exact stage plus Err.Number, Err.Source, and Err.Description in the error handler. This is diagnostic only; no updater architecture or HTTP behavior was changed.

Commits:
- 1c2d556504cd8286c0f04d04a2f505d4a8d3b063 — exact error-stage logging
- 832171660e148506d6892c6ec9ae4d83408b0686 — bump DEV build to 49

Next: test DEV-49 once published and inspect the resulting checkfor log. The goal is to obtain the exact VBA error number/source/stage, not another generic screenshot.
