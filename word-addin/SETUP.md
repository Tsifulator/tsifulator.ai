# tsifl Word Add-in — Setup & Sideloading

## Prerequisites
- Node.js 18+
- Microsoft Word (desktop, macOS or Windows)
- Office dev certs at `~/.office-addin-dev-certs/` (localhost.key, localhost.crt)

## Install & Build
```bash
cd word-addin
npm install
npm run build          # production build to dist/
npm run start          # dev server on https://localhost:3002
```

## Generate Dev Certs (if missing)
```bash
npx office-addin-dev-certs install
```
This creates `~/.office-addin-dev-certs/localhost.key` and `localhost.crt`.

## Sideload into Word

### macOS
1. Start the dev server: `npm run start`
2. Open Word
3. Go to **Insert** > **Add-ins** > **My Add-ins** > **Upload My Add-in**
4. Browse to `word-addin/manifest.xml` and click **Upload**
5. The tsifl button appears in the **Home** ribbon tab

### Windows
1. Start the dev server: `npm run start`
2. Open Word
3. Go to **Insert** > **Get Add-ins** > **Upload My Add-in**
4. Browse to `word-addin/manifest.xml` and click **Upload**
5. The tsifl button appears in the **Home** ribbon tab

### Alternative: Shared Folder (Windows)
1. Copy `dist/manifest.xml` to a network share
2. In Word: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs
3. Add the share URL and check "Show in Menu"
4. Restart Word, then Insert > My Add-ins > Shared Folder

## Manifest Details
- **Host**: Document (Word)
- **Dev Server Port**: 3002
- **Taskpane URL**: https://localhost:3002/taskpane.html
- **Permissions**: ReadWriteDocument

## Troubleshooting
- **Cert errors**: Re-run `npx office-addin-dev-certs install` and trust the CA
- **Add-in not loading**: Make sure dev server is running on port 3002
- **"We can't open this add-in"**: Clear Office cache: `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`
