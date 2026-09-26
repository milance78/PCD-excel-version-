# PCD Excel Version

This repository is the Excel/VBA reconstruction of the reference project `milance78/interventions-pcd`.

## Current build

**0.1.0-corp-test**

The current GitHub Actions build produces a genuine `.xlsm` containing an embedded VBA project and the first corporate-test workbook layout/import mapping.

Important: this is an **early corporate test build**, not the final PCD replacement. The full VBA parser source is versioned in `vba/PCD_MagicImport.bas`; the first downloadable artifact uses a conservative build path while VBA project wiring is being finalized.

## Reference vs Excel project

- Reference: `milance78/interventions-pcd`
- Excel project: `milance78/PCD-excel-version-`
- Runtime target: local Excel/offline
- GitHub: development distribution/build source, not a runtime dependency

## Build

GitHub Actions workflow: **Build PCD Excel XLSM**.
