# tsifl VS Code Extension — Setup

## Install from Source
```bash
cd vscode-extension
npm install
```

## Load in VS Code (Development)
1. Open the `vscode-extension/` folder in VS Code
2. Press **F5** to launch Extension Development Host
3. In the new VS Code window, click the tsifl icon in the activity bar (left sidebar)
4. Sign in with your tsifl account

## Package as .vsix
```bash
cd vscode-extension
npx @vscode/vsce package --no-dependencies
```
This creates `tsifl-1.0.0.vsix`. Install via:
```bash
code --install-extension tsifl-1.0.0.vsix
```

## Features
- **Activity bar icon**: Click the tsifl icon to open the chat sidebar
- **Context awareness**: Reads current file, selection, language, diagnostics, git status
- **Commands** (right-click menu or Command Palette):
  - `tsifl: Explain Selected Code`
  - `tsifl: Refactor Selected Code`
  - `tsifl: Generate Tests`
  - `tsifl: Fix Error`
  - `tsifl: Open Chat`
- **Actions**: insert code, replace selection, create/edit files, run terminal commands, show diffs
- **Image support**: Paste screenshots, attach image files
- **Auth**: Supabase (same account as all tsifl add-ins)

## Keyboard Shortcuts
Customize in VS Code Keyboard Shortcuts (`Cmd+K Cmd+S`):
- Search for "tsifl" to find all commands
