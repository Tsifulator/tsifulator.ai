# tsifl Google Workspace Add-on — Setup & Deployment

## Prerequisites
- Google account
- Node.js with `clasp` CLI: `npm install -g @google/clasp`
- Enable Apps Script API: https://script.google.com/home/usersettings

## Initial Setup

### 1. Login to clasp
```bash
clasp login
```

### 2. Create a new Apps Script project
```bash
cd google-workspace-addon
clasp create --type standalone --title "tsifl"
```
This updates `.clasp.json` with your script ID.

### 3. Push code
```bash
clasp push
```

### 4. Open in browser
```bash
clasp open
```

## Deploy as Add-on

### Test as Editor Add-on
1. Open a Google Sheet, Doc, or Slides presentation
2. Go to **Extensions** > **Apps Script**
3. Copy the script ID from the URL
4. Update `.clasp.json` with the script ID
5. Run `clasp push`
6. In the script editor, run `onOpen` once to grant permissions
7. Reload the document — "tsifl" menu appears

### Deploy as Workspace Add-on
1. In the Apps Script editor, click **Deploy** > **Test deployments**
2. Add yourself as a test user
3. Go to the document type (Sheets/Docs/Slides) and the add-on appears in the sidebar

### Publish (for organization)
1. Deploy > New deployment > Add-on
2. Fill in details and submit for review

## Files
- **Code.gs** — Main server-side logic: auth, context capture, action execution
- **Sidebar.html** — Client-side chat UI (same tsifl design system)
- **appsscript.json** — Manifest with OAuth scopes
- **.clasp.json** — clasp deployment config

## Supported Apps
- **Google Sheets**: Read cells/formulas, write cells, format, charts, sort
- **Google Docs**: Read/insert text, paragraphs, tables, headers, find/replace
- **Google Slides**: Create slides, add text boxes, shapes, tables, images

## Troubleshooting
- **"Authorization required"**: Run any function from the script editor first to grant OAuth
- **Sidebar not appearing**: Reload the document after pushing new code
- **"Script function not found"**: Make sure `clasp push` succeeded
