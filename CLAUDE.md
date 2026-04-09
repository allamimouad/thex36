# SharePoint Explorer — Project Brief

Reusable Angular drag-and-drop explorer component, built to be transplanted into another app later.

## Core features

- Folder tree on the left, current folder content on the right
- Drag/drop move for files and folders
- Native desktop file drop for upload

## Architecture rules

- UI must not depend directly on real SharePoint API code
- SharePoint/backend integration stays behind small service boundaries
- Abstraction must remain easy to copy into another app (swap adapters, not rewrite UI)

## Key files

- `src/app/features/sharepoint-explorer/sharepoint-explorer.component.ts` — UI and drag/drop behavior
- `src/app/features/sharepoint-explorer/services/sharepoint-explorer.service.ts` — state, tree loading, refresh after moves/uploads
- `src/app/features/sharepoint-explorer/services/sharepoint-explorer.client.ts` — abstract client contract
- `src/app/features/sharepoint-explorer/services/sharepoint-explorer-upload.service.ts` — upload extraction/upload stub

## Current status

- App runs with the simulated client
- Real HTTP adapter exists only as a placeholder — backend integration will be finished later on another laptop

## What NOT to do

- Do not finish or expand the HTTP adapter here
- Do not write tests for this work
- Do not refactor the abstraction for cleanliness alone
- Do not change the service boundaries unless there is a real blocker
