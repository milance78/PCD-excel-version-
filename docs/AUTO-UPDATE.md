# PCD Excel automatic update

The workbook checks VERSION.json on GitHub only when the user invokes the update command.

Flow:
1. Read VERSION.json.
2. Compare local and remote version.
3. Ask for confirmation.
4. Download the new XLSM over HTTPS.
5. Verify SHA-256.
6. Close the current workbook.
7. Replace it using a temporary Windows script.
8. Reopen the updated workbook.

Runtime remains local/offline when no update check is requested.

Corporate fallback:
- If WinHTTP/GitHub is blocked, the workbook must leave the existing version untouched and show the error.
- No unverified workbook is installed.
