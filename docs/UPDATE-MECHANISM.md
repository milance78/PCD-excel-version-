# Update mechanism

## Goal

A corporate PC should be able to obtain the latest approved Excel application without requiring the React application or any corporate server.

GitHub is used as the development/distribution source.

## Important distinction

The Excel application does not depend on GitHub while it is running.

GitHub is only used to distribute a newer version.

## Version flow

```
PCD Excel source
      |
      v
GitHub repository
      |
      v
approved version + checksum
      |
      v
corporate PC updater
      |
      v
local PCD Excel .xlsm
      |
      v
offline/local runtime
```

The exact transport will be implemented after the first working `.xlsm` exists and tested against the actual corporate restrictions.

## Security requirements

- Never execute downloaded VBA without verification.
- Verify the expected SHA-256 checksum before replacing the local application.
- Keep the currently working local copy until the new version has been verified.
- Do not silently replace a working application with an invalid or incomplete download.
- The final offline package must not depend on network access.
