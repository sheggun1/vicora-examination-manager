# VICORA Exam Room Manager

Standalone exam-room allocation application. GitHub Pages frontend + Google Apps Script backend + the existing VICORA Exam Room Manager spreadsheet.

## Important separation rule
The backend uses only this spreadsheet:
`12odSjNeGBTldU7rWoV27p7SHdKlkBhNdn1oWoQY_Fb0`

It does not connect to, import from, modify, or synchronise with GTMC. Do not change the database ID in `Untitled.gs`.

## Update the EXISTING Apps Script project
Do not create a new Apps Script project or a new spreadsheet.

1. Open the existing VICORA Exam Room Manager Apps Script project.
2. Make a private backup copy of the current script source.
3. Replace the old backend source with the complete contents of `Untitled.gs`. If the project has other .gs files, check that they do not define duplicate `doGet` or `doPost` functions. Do not delete unrelated files or triggers without checking them.
4. In `Untitled.gs`, replace `CHANGE-THIS-TO-A-LONG-PRIVATE-KEY` with a long, unique private administrator key. Do not publish or share that key.
5. Save the project.
6. In Deploy → Manage deployments, edit the existing Web App deployment, select **New version**, and deploy. Keep execution as **Me**. Access must permit the public student lookup endpoint and the web frontend to call the app; the admin key protects administrator actions.
7. Copy the deployed Web App URL ending in `/exec`.

The backend performs non-destructive schema checks. It creates missing tabs and appends missing required header columns; it does not clear existing tabs or records. Review any existing tab headers before the first live run. Keep a backup.

## GitHub Pages
In repository Settings → Pages, select Deploy from a branch, choose `main` and `/(root)`, then save. Wait for GitHub Pages to publish the site.

Open the published website, enter the Apps Script Web App URL and administrator key, then select Connect.

## Safe first-run test
1. Connect and confirm the database name/version shown by the backend.
2. Create a clearly named **DRAFT TEST** exam session.
3. Add test room settings with small capacities.
4. Generate a preview and inspect the complete candidate list and class distribution.
5. Do not confirm a real examination until candidate eligibility, class order, subject counts and room assignments have been checked against the source data.
6. Confirm a test allocation only when ready, then verify the Allocations and MutationLog tabs.
7. Test public lookup with a test student whose allocation is confirmed.

## Allocation behaviour
Candidates are ordered by their class's first appearance in the imported Students tab, then by their original student order. The room sequence carries forward as the next class begins. Room capacities are enforced. Confirmed sessions are locked; reallocation requires a reason and writes previous/new room details to AllocationHistory.

SS2B is excluded from the current candidate pool but retained in the database for future sessions. The two previously identified non-participants remain excluded from fallback candidate selection.

## Current limitations / deployment gate
This repository update is code-complete for the current feature scope, but it has not been executed against the live VICORA spreadsheet or deployed Apps Script URL in this session. Before a real examination, verify actual tab headers, candidate eligibility and allocation output using the safe test above. Do not treat a successful GitHub commit as proof of live deployment.
