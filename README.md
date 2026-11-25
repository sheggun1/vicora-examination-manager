# VICORA Exam Room Manager

Standalone examination room allocation website.

## Frontend
This repository hosts the GitHub Pages frontend in `index.html`.

## Backend setup
The Google Apps Script backend remains in the separate Apps Script project as `Untitled.gs`. It must use the VICORA database spreadsheet only. Do not add the GTMC spreadsheet ID or old GTMC import/sync code.

1. Paste the standalone backend into Apps Script.
2. Set a long private value for `VICORA_ADMIN_KEY`.
3. Deploy as a Web App, executing as you, with access set to Anyone.
4. Keep the deployment URL and admin key private.
5. Open the GitHub Pages website and enter the Web App URL and admin key in the connection panel.

## GitHub Pages
In repository Settings → Pages, select Deploy from a branch, choose `main` and `/(root)`, then save.

## Important
- The public website does not contain the admin key by default; enter it at runtime.
- Test with a draft exam session and sample rooms before using for a live examination.
- Verify every saved change in the VICORA spreadsheet.
- The app must not reconnect to or modify the GTMC source spreadsheet.
