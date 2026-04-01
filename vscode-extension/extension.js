/**
 * tsifl VS Code Extension v2.3.0
 * Uses https module instead of fetch() for Node.js compatibility.
 */

const vscode = require("vscode");
const https = require("https");

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";

let globalPanel = null;

// ── HTTP helper (works in all Node.js versions) ─────────────────────────

function httpPost(url, body) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body);
    const parsed = new URL(url);
    const options = {
      hostname: parsed.hostname,
      port: 443,
      path: parsed.pathname + parsed.search,
      method: "POST",
      headers: { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(data) },
    };
    const req = https.request(options, (res) => {
      let chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        try {
          const text = Buffer.concat(chunks).toString();
          resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, status: res.statusCode, json: () => JSON.parse(text), text: () => text });
        } catch (e) { reject(e); }
      });
    });
    req.on("error", reject);
    req.setTimeout(90000, () => { req.destroy(); reject(new Error("Request timed out")); });
    req.write(data);
    req.end();
  });
}

function httpGet(url) {
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      let chunks = [];
      res.on("data", (chunk) => chunks.push(chunk));
      res.on("end", () => {
        try {
          const text = Buffer.concat(chunks).toString();
          resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, json: () => JSON.parse(text) });
        } catch (e) { reject(e); }
      });
    }).on("error", reject);
  });
}

// ── Activation ──────────────────────────────────────────────────────────

function activate(context) {
  const provider = new TsiflSidebarProvider(context.extensionUri);

  context.subscriptions.push(
    vscode.window.registerWebviewViewProvider("tsifl.chatView", provider, {
      webviewOptions: { retainContextWhenHidden: true },
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.openChat", () => openChatPanel(context, provider))
  );
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.explainCode", () => ensurePanelAndSend(context, provider, "Explain this code"))
  );
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.fixError", () => ensurePanelAndSend(context, provider, "Fix the error in this code"))
  );
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.refactor", () => ensurePanelAndSend(context, provider, "Refactor this code"))
  );
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.generateTests", () => ensurePanelAndSend(context, provider, "Generate tests for this code"))
  );

  // Git commit message generator (Improvement 47)
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.generateCommit", async () => {
      try {
        const terminal = vscode.window.activeTerminal || vscode.window.createTerminal("tsifl");
        // Get staged diff
        const { exec } = require("child_process");
        const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri?.fsPath || ".";
        exec("git diff --staged", { cwd: workspaceRoot, maxBuffer: 50000 }, (err, stdout) => {
          if (err || !stdout.trim()) {
            vscode.window.showWarningMessage("No staged changes. Stage files first with `git add`.");
            return;
          }
          ensurePanelAndSend(context, provider, `Generate a concise commit message for these staged changes:\n\n${stdout.slice(0, 3000)}`);
        });
      } catch (e) {
        vscode.window.showErrorMessage("Failed to get git diff: " + e.message);
      }
    })
  );

  // Error detection from terminal (Improvement 46)
  vscode.window.onDidWriteTerminalData && vscode.window.onDidWriteTerminalData(event => {
    const text = event.data;
    if (/Error:|TypeError:|SyntaxError:|ReferenceError:|ModuleNotFoundError/i.test(text)) {
      const errorSnippet = text.slice(0, 500);
      vscode.window.showInformationMessage("tsifl detected an error. Click to fix.", "Fix with tsifl").then(choice => {
        if (choice === "Fix with tsifl") {
          ensurePanelAndSend(context, provider, `Fix this error from the terminal:\n\n${errorSnippet}`);
        }
      });
    }
  });

  // Auto-open panel if sidebar doesn't appear after 3s
  setTimeout(() => {
    if (!provider._view && !globalPanel) {
      openChatPanel(context, provider);
    }
  }, 3000);

  // Show a status bar indicator so the user knows tsifl loaded
  const statusItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
  statusItem.text = "$(comment-discussion) tsifl";
  statusItem.tooltip = "tsifl — AI for Financial Analysts (click to open)";
  statusItem.command = "tsifl.openChat";
  statusItem.show();
  context.subscriptions.push(statusItem);
}

function ensurePanelAndSend(context, provider, prompt) {
  const wv = globalPanel?.webview || provider._view?.webview;
  if (wv) {
    wv.postMessage({ type: "sendPrompt", prompt });
  } else {
    openChatPanel(context, provider);
    setTimeout(() => {
      const w = globalPanel?.webview || provider._view?.webview;
      if (w) w.postMessage({ type: "sendPrompt", prompt });
    }, 1500);
  }
}

function openChatPanel(context, provider) {
  if (globalPanel) { globalPanel.reveal(vscode.ViewColumn.Beside); return; }
  globalPanel = vscode.window.createWebviewPanel("tsifl.chat", "tsifl", vscode.ViewColumn.Beside, {
    enableScripts: true,
    retainContextWhenHidden: true,
  });
  globalPanel.webview.html = provider._getHtml(globalPanel.webview);
  globalPanel.webview.onDidReceiveMessage(async (msg) => {
    if (msg.type === "chat") await provider._handleChat(msg, globalPanel.webview);
    else if (msg.type === "getContext") await provider._sendContext(globalPanel.webview);
    else if (msg.type === "executeAction") await provider._executeAction(msg.action, msg.userId, globalPanel.webview);
  });
  globalPanel.onDidDispose(() => { globalPanel = null; });
}

// ── Sidebar Provider ────────────────────────────────────────────────────

class TsiflSidebarProvider {
  constructor(extensionUri) { this._extensionUri = extensionUri; this._view = null; }

  resolveWebviewView(webviewView) {
    this._view = webviewView;
    webviewView.webview.options = { enableScripts: true, localResourceRoots: [this._extensionUri] };
    webviewView.webview.html = this._getHtml(webviewView.webview);
    webviewView.webview.onDidReceiveMessage(async (msg) => {
      if (msg.type === "chat") await this._handleChat(msg, webviewView.webview);
      else if (msg.type === "getContext") await this._sendContext(webviewView.webview);
      else if (msg.type === "executeAction") await this._executeAction(msg.action, msg.userId, webviewView.webview);
    });
  }

  _wv(override) { return override || globalPanel?.webview || this._view?.webview; }

  _getSessionId() {
    if (!this._sessionId) {
      const ws = vscode.workspace.workspaceFolders?.[0]?.uri?.fsPath || "default";
      const crypto = require("crypto");
      this._sessionId = "vscode-" + crypto.createHash("md5").update(ws).digest("hex").slice(0, 12);
    }
    return this._sessionId;
  }

  async _handleChat(message, webview) {
    const wv = this._wv(webview);
    try {
      const context = await this._getContext();
      const resp = await httpPost(`${BACKEND_URL}/chat/`, {
        user_id: message.userId,
        message: message.text,
        context,
        session_id: this._getSessionId(),
        images: message.images || [],
      });
      if (!resp.ok) {
        const err = resp.json();
        wv?.postMessage({ type: "chatResponse", error: err.detail || "Request failed" });
        return;
      }
      const result = resp.json();
      wv?.postMessage({ type: "chatResponse", result });
      const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
      for (const a of actions) await this._executeAction(a, message.userId, wv);
    } catch (e) {
      wv?.postMessage({ type: "chatResponse", error: e.message });
    }
  }

  async _getContext() {
    const editor = vscode.window.activeTextEditor;
    const ctx = { app: "vscode", workspace: vscode.workspace.name || "", open_files: vscode.window.visibleTextEditors.map(e => e.document.fileName) };
    if (editor) {
      const doc = editor.document;
      ctx.current_file = doc.fileName;
      ctx.language = doc.languageId;
      ctx.line_count = doc.lineCount;
      ctx.cursor_line = editor.selection.active.line + 1;
      if (!editor.selection.isEmpty) ctx.selection = doc.getText(editor.selection);
      if (editor.visibleRanges.length > 0) ctx.visible_text = doc.getText(editor.visibleRanges[0]).substring(0, 3000);
      ctx.file_content = doc.getText().substring(0, 5000);
    }
    const diags = [];
    vscode.languages.getDiagnostics().forEach(([uri, d]) => {
      d.forEach((x) => {
        if (x.severity <= vscode.DiagnosticSeverity.Warning)
          diags.push({ file: uri.fsPath, line: x.range.start.line + 1, severity: x.severity === 0 ? "error" : "warning", message: x.message });
      });
    });
    ctx.diagnostics = diags.slice(0, 20);
    try {
      const git = vscode.extensions.getExtension("vscode.git");
      if (git?.isActive) { const api = git.exports.getAPI(1); if (api.repositories.length) { ctx.git_branch = api.repositories[0].state.HEAD?.name || ""; ctx.git_changes = api.repositories[0].state.workingTreeChanges.length; } }
    } catch (_) {}
    // File tree context (Improvement 48)
    try {
      const fs = require("fs");
      const path = require("path");
      const root = vscode.workspace.workspaceFolders?.[0]?.uri?.fsPath;
      if (root) {
        const tree = [];
        const ignored = new Set(["node_modules", ".git", ".vscode", "__pycache__", "dist", "build", ".next", "venv", ".env"]);
        const walk = (dir, depth) => {
          if (depth > 2 || tree.length >= 50) return;
          try {
            const items = fs.readdirSync(dir);
            for (const item of items) {
              if (ignored.has(item) || item.startsWith(".")) continue;
              if (tree.length >= 50) break;
              const full = path.join(dir, item);
              const rel = path.relative(root, full);
              const stat = fs.statSync(full);
              tree.push({ path: rel, type: stat.isDirectory() ? "dir" : "file" });
              if (stat.isDirectory()) walk(full, depth + 1);
            }
          } catch (_) {}
        };
        walk(root, 0);
        ctx.file_tree = tree;
        // Detect framework (Improvement 55)
        try {
          const pkgPath = path.join(root, "package.json");
          if (fs.existsSync(pkgPath)) {
            const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf8"));
            const deps = { ...pkg.dependencies, ...pkg.devDependencies };
            if (deps.react) ctx.framework = "react";
            else if (deps.vue) ctx.framework = "vue";
            else if (deps.angular) ctx.framework = "angular";
            else if (deps.next) ctx.framework = "nextjs";
            else if (deps.express) ctx.framework = "express";
            else ctx.framework = "node";
          }
        } catch (_) {}
        try {
          if (fs.existsSync(path.join(root, "requirements.txt")) || fs.existsSync(path.join(root, "pyproject.toml"))) {
            ctx.framework = ctx.framework || "python";
          }
        } catch (_) {}
      }
    } catch (_) {}
    return ctx;
  }

  async _sendContext(webview) {
    const wv = this._wv(webview);
    const ctx = await this._getContext();
    wv?.postMessage({ type: "context", context: ctx });
  }

  async _executeAction(action, userId, webview) {
    if (!action?.type || !action?.payload) return;
    const { type, payload } = action;
    const wv = this._wv(webview);
    try {
      switch (type) {
        case "insert_code": {
          const editor = vscode.window.activeTextEditor;
          if (editor) { await editor.edit(b => b.insert(editor.selection.active, payload.code || payload.text || "")); }
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Code inserted" });
          return;
        }
        case "replace_selection": {
          const editor = vscode.window.activeTextEditor;
          if (editor) {
            const code = payload.code || payload.text || "";
            if (!editor.selection.isEmpty) { await editor.edit(b => b.replace(editor.selection, code)); }
            else { await editor.edit(b => b.insert(editor.selection.active, code)); }
          }
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Code applied" });
          return;
        }
        case "create_file": {
          let fullPath = payload.path;
          if (!fullPath.startsWith("/") && vscode.workspace.workspaceFolders?.length)
            fullPath = vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, fullPath).fsPath;
          const uri = vscode.Uri.file(fullPath);
          await vscode.workspace.fs.writeFile(uri, Buffer.from(payload.content || "", "utf8"));
          const doc = await vscode.workspace.openTextDocument(uri);
          await vscode.window.showTextDocument(doc);
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "File created: " + payload.path });
          return;
        }
        case "edit_file": {
          let fullPath = payload.path;
          if (!fullPath.startsWith("/") && vscode.workspace.workspaceFolders?.length)
            fullPath = vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, fullPath).fsPath;
          const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(fullPath));
          const editor = await vscode.window.showTextDocument(doc);
          if (payload.find && payload.replace !== undefined) {
            const text = doc.getText(); const idx = text.indexOf(payload.find);
            if (idx !== -1) { const start = doc.positionAt(idx); const end = doc.positionAt(idx + payload.find.length); await editor.edit(b => b.replace(new vscode.Range(start, end), payload.replace)); }
          }
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "File edited" });
          return;
        }
        case "run_terminal_command": case "run_shell_command": {
          const terminal = vscode.window.activeTerminal || vscode.window.createTerminal("tsifl");
          terminal.show(); terminal.sendText(payload.command);
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Running: " + payload.command });
          return;
        }
        case "open_file": {
          let fullPath = payload.path;
          if (!fullPath.startsWith("/") && vscode.workspace.workspaceFolders?.length)
            fullPath = vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, fullPath).fsPath;
          const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(fullPath));
          await vscode.window.showTextDocument(doc);
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Opened " + payload.path });
          return;
        }
        case "show_diff": {
          const ch = vscode.window.createOutputChannel("tsifl Diff");
          ch.clear(); ch.appendLine("=== BEFORE ==="); ch.appendLine(payload.before || ""); ch.appendLine("\n=== AFTER ==="); ch.appendLine(payload.after || ""); ch.show();
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Diff shown" });
          return;
        }
        case "launch_app": {
          try {
            const resp = await httpPost(`${BACKEND_URL}/launch-app`, { app_name: payload.app_name });
            const r = resp.json();
            wv?.postMessage({ type: "actionComplete", action: type, success: true, message: r.message || "Launched" });
          } catch (e) { wv?.postMessage({ type: "actionComplete", action: type, success: false, message: e.message }); }
          return;
        }
        case "open_notes": {
          vscode.env.openExternal(vscode.Uri.parse(`${BACKEND_URL}/notes-app`));
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Opened Notes" });
          return;
        }
        case "open_url": {
          if (payload.url) vscode.env.openExternal(vscode.Uri.parse(payload.url));
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Opened " + (payload.url || "") });
          return;
        }
        case "create_note": {
          try {
            const resp = await httpPost(`${BACKEND_URL}/notes/`, { user_id: userId || "vscode-user", title: payload.title || "Untitled", content: payload.content || "", folder: "General" });
            const note = resp.json();
            wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Note created: " + (note.title || "Untitled") });
          } catch (e) { wv?.postMessage({ type: "actionComplete", action: type, success: false, message: e.message }); }
          return;
        }
        case "explain_code": case "fix_error": case "refactor": case "generate_tests": return;
      }
      wv?.postMessage({ type: "actionComplete", action: type, success: true });
    } catch (e) {
      wv?.postMessage({ type: "actionComplete", action: type, success: false, error: e.message });
    }
  }

  _getHtml(webview) {
    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <style>
    :root {
      --blue: #0D5EAF;
      --bh: #0A4896;
      --bl: var(--vscode-textLink-activeForeground, #EBF3FB);
      --bg: var(--vscode-editor-background, #1e1e1e);
      --sf: var(--vscode-sideBar-background, #252526);
      --bd: var(--vscode-panel-border, #3c3c3c);
      --tx: var(--vscode-foreground, #cccccc);
      --mu: var(--vscode-descriptionForeground, #8b8b8b);
      --gn: #4ade80;
      --rd: #f87171;
      --input-bg: var(--vscode-input-background, #3c3c3c);
      --input-bd: var(--vscode-input-border, #3c3c3c);
      --input-fg: var(--vscode-input-foreground, #cccccc);
      --badge-bg: var(--vscode-badge-background, #0D5EAF);
      --badge-fg: var(--vscode-badge-foreground, #ffffff);
      --code-bg: var(--vscode-textCodeBlock-background, #1e1e1e);
      --code-tx: var(--vscode-editor-foreground, #d4d4d4);
      --hover-bg: var(--vscode-list-hoverBackground, #2a2d2e);
      --focus-bd: var(--vscode-focusBorder, #007fd4);
      --btn-bg: var(--vscode-button-background, #0D5EAF);
      --btn-fg: var(--vscode-button-foreground, #ffffff);
      --btn-hover: var(--vscode-button-hoverBackground, #0A4896);
      --btn-sec-bg: var(--vscode-button-secondaryBackground, #3a3d41);
      --btn-sec-fg: var(--vscode-button-secondaryForeground, #cccccc);
      --error-fg: var(--vscode-errorForeground, #f87171);
      --font-family: var(--vscode-font-family, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif);
      --mono-font: var(--vscode-editor-font-family, 'Menlo', 'Consolas', 'Courier New', monospace);
      --font-size: var(--vscode-font-size, 13px);
      --editor-font-size: var(--vscode-editor-font-size, 12px);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: var(--font-family);
      font-size: var(--font-size);
      background: var(--bg);
      color: var(--tx);
      height: 100vh;
      -webkit-font-smoothing: antialiased;
      overflow: hidden;
    }
    #app { display: flex; flex-direction: column; height: 100vh; }

    /* Login */
    #login-screen { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; padding: 24px; gap: 8px; }
    #login-screen h2 { color: var(--btn-bg); font-size: 20px; font-weight: 700; margin-bottom: 2px; letter-spacing: -0.3px; }
    #login-tagline { font-size: 12px; color: var(--mu); margin-bottom: 20px; }
    #login-form { display: flex; flex-direction: column; gap: 8px; width: 100%; max-width: 280px; }
    #login-form input {
      width: 100%; background: var(--input-bg); color: var(--input-fg);
      border: 1px solid var(--input-bd); border-radius: 4px; padding: 7px 10px;
      font-size: 13px; outline: none; font-family: var(--font-family);
    }
    #login-form input:focus { border-color: var(--focus-bd); outline: 1px solid var(--focus-bd); outline-offset: -1px; }
    #login-btn {
      background: var(--btn-bg); color: var(--btn-fg); border: none; border-radius: 4px;
      padding: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.1s;
    }
    #login-btn:hover { background: var(--btn-hover); }
    #signup-btn {
      background: var(--btn-sec-bg); color: var(--btn-sec-fg); border: none; border-radius: 4px;
      padding: 7px; font-size: 13px; cursor: pointer; transition: background 0.1s;
    }
    #signup-btn:hover { opacity: 0.9; }
    #auth-error { font-size: 12px; color: var(--error-fg); text-align: center; min-height: 16px; }

    /* Chat screen */
    #chat-screen { display: none; flex-direction: column; height: 100vh; }
    #header-bar {
      display: flex; align-items: center; justify-content: space-between;
      padding: 8px 12px; border-bottom: 1px solid var(--bd); background: var(--sf);
      min-height: 36px;
    }
    #header-bar .title { color: var(--btn-bg); font-size: 13px; font-weight: 700; letter-spacing: -0.2px; }
    #header-bar .hdr-btns { display: flex; gap: 6px; align-items: center; }
    .hdr-btn {
      background: var(--btn-sec-bg); color: var(--btn-sec-fg); border: none; border-radius: 4px;
      padding: 3px 10px; font-size: 11px; font-weight: 500; cursor: pointer; transition: background 0.1s;
    }
    .hdr-btn:hover { background: var(--hover-bg); }
    #user-bar { font-size: 11px; color: var(--mu); padding: 4px 12px; background: var(--sf); border-bottom: 1px solid var(--bd); cursor: default; }

    /* Quick actions */
    #quick-actions { display: flex; gap: 4px; padding: 8px 10px; flex-wrap: wrap; border-bottom: 1px solid var(--bd); background: var(--bg); }
    .quick-btn {
      background: var(--btn-sec-bg); color: var(--btn-sec-fg); border: none; border-radius: 4px;
      padding: 3px 10px; font-size: 11px; cursor: pointer; transition: all 0.1s; white-space: nowrap;
    }
    .quick-btn:hover { background: var(--btn-bg); color: var(--btn-fg); }

    /* Chat history */
    #chat-history {
      flex: 1; overflow-y: auto; padding: 12px; display: flex; flex-direction: column; gap: 8px;
      scroll-behavior: smooth;
    }
    #chat-history::-webkit-scrollbar { width: 6px; }
    #chat-history::-webkit-scrollbar-track { background: transparent; }
    #chat-history::-webkit-scrollbar-thumb { background: var(--bd); border-radius: 3px; }
    #chat-history::-webkit-scrollbar-thumb:hover { background: var(--mu); }

    /* Messages */
    .msg {
      padding: 10px 14px; border-radius: 6px; line-height: 1.65; word-wrap: break-word;
      font-size: 13px; animation: msgSlideIn 0.25s ease-out;
    }
    .msg.user {
      background: color-mix(in srgb, var(--btn-bg) 12%, transparent);
      border-left: 3px solid var(--btn-bg);
      margin-right: 16px;
    }
    .msg.assistant {
      background: var(--sf);
      border: 1px solid var(--bd);
      margin-left: 0;
    }
    .msg.error {
      background: color-mix(in srgb, var(--rd) 10%, transparent);
      border: 1px solid color-mix(in srgb, var(--rd) 30%, transparent);
      color: var(--error-fg);
    }
    .msg.assistant p { margin: 0 0 8px 0; }
    .msg.assistant p:last-child { margin-bottom: 0; }
    .msg.assistant strong { font-weight: 600; color: var(--tx); }
    .msg.assistant em { font-style: italic; opacity: 0.9; }
    .msg.assistant ul, .msg.assistant ol { margin: 6px 0 6px 20px; }
    .msg.assistant li { margin: 3px 0; list-style: disc; }
    .msg.assistant ol li { list-style: decimal; }
    .msg.assistant hr { border: none; border-top: 1px solid var(--bd); margin: 10px 0; }

    /* Inline code */
    .msg.assistant code:not(.code-block-code) {
      background: var(--code-bg); padding: 2px 6px; border-radius: 3px;
      font-size: var(--editor-font-size); font-family: var(--mono-font);
      border: 1px solid var(--bd);
    }

    /* Code blocks */
    .code-block-wrap {
      position: relative; margin: 10px 0; border-radius: 6px; overflow: hidden;
      border: 1px solid var(--bd); background: var(--code-bg);
    }
    .code-block-header {
      display: flex; align-items: center; justify-content: space-between;
      padding: 4px 12px; background: color-mix(in srgb, var(--bd) 50%, var(--code-bg));
      border-bottom: 1px solid var(--bd); font-size: 11px; min-height: 28px;
    }
    .code-block-lang {
      color: var(--mu); font-family: var(--mono-font); font-size: 11px;
      text-transform: lowercase; letter-spacing: 0.3px;
    }
    .code-copy-btn {
      background: transparent; color: var(--mu); border: 1px solid transparent;
      border-radius: 4px; padding: 2px 8px; font-size: 11px; cursor: pointer;
      font-family: var(--font-family); transition: all 0.15s; display: flex; align-items: center; gap: 4px;
    }
    .code-copy-btn:hover { background: var(--hover-bg); color: var(--tx); border-color: var(--bd); }
    .code-copy-btn.copied { color: var(--gn); }
    .code-block-pre {
      margin: 0; padding: 12px 16px; overflow-x: auto; background: var(--code-bg);
      font-family: var(--mono-font); font-size: var(--editor-font-size);
      line-height: 1.55; color: var(--code-tx); tab-size: 4;
    }
    .code-block-pre::-webkit-scrollbar { height: 6px; }
    .code-block-pre::-webkit-scrollbar-thumb { background: var(--bd); border-radius: 3px; }
    .code-block-code { font-family: inherit; font-size: inherit; background: transparent; border: none; padding: 0; }

    /* Diff highlights inside code blocks */
    .diff-add { color: #4ade80; background: rgba(74,222,128,0.08); display: inline-block; width: 100%; padding: 0 4px; margin: 0 -4px; }
    .diff-del { color: #f87171; background: rgba(248,113,113,0.08); display: inline-block; width: 100%; padding: 0 4px; margin: 0 -4px; text-decoration: line-through; }

    /* Action messages */
    .msg.action {
      background: color-mix(in srgb, var(--gn) 8%, transparent);
      border-left: 3px solid var(--gn);
      font-size: 12px; color: var(--gn); font-family: var(--mono-font);
      padding: 6px 12px; display: flex; align-items: center; gap: 6px;
    }
    .msg.action.action-fail {
      border-left-color: var(--rd); color: var(--rd);
      background: color-mix(in srgb, var(--rd) 8%, transparent);
    }
    .action-icon { font-size: 14px; flex-shrink: 0; }

    /* Typing indicator */
    .typing-ind {
      display: flex; gap: 5px; padding: 12px 14px; align-items: center;
      background: var(--sf); border: 1px solid var(--bd); border-radius: 6px;
      animation: msgSlideIn 0.25s ease-out;
    }
    .typing-ind span {
      width: 7px; height: 7px; background: var(--mu); border-radius: 50%;
      animation: typingBounce 1.4s infinite ease-in-out;
    }
    .typing-ind span:nth-child(2) { animation-delay: 0.16s; }
    .typing-ind span:nth-child(3) { animation-delay: 0.32s; }

    /* Input area */
    #input-area { padding: 10px 12px; border-top: 1px solid var(--bd); display: flex; flex-direction: column; gap: 6px; background: var(--sf); }
    #image-preview-bar { display: none; gap: 4px; flex-wrap: wrap; }
    .img-preview { position: relative; display: inline-block; }
    .img-preview img { width: 40px; height: 40px; object-fit: cover; border-radius: 4px; border: 1px solid var(--bd); }
    .img-remove {
      position: absolute; top: -4px; right: -4px; width: 16px; height: 16px;
      background: var(--rd); color: white; border: none; border-radius: 50%;
      font-size: 10px; line-height: 16px; text-align: center; cursor: pointer; padding: 0;
    }
    #user-input {
      width: 100%; background: var(--input-bg); color: var(--input-fg);
      border: 1px solid var(--input-bd); border-radius: 6px; padding: 8px 10px;
      font-size: 13px; font-family: var(--font-family); resize: none; outline: none;
      line-height: 1.45; max-height: 120px; overflow-y: auto;
    }
    #user-input:focus { border-color: var(--focus-bd); outline: 1px solid var(--focus-bd); outline-offset: -1px; }
    #user-input::placeholder { color: var(--mu); }
    #input-actions { display: flex; gap: 4px; }
    #attach-btn {
      background: var(--btn-sec-bg); color: var(--btn-sec-fg); border: none; border-radius: 6px;
      padding: 6px 12px; font-size: 16px; cursor: pointer; flex-shrink: 0; transition: all 0.1s;
      display: flex; align-items: center; justify-content: center;
    }
    #attach-btn:hover { background: var(--hover-bg); }
    #submit-btn {
      background: var(--btn-bg); color: var(--btn-fg); border: none; border-radius: 6px;
      padding: 6px; font-size: 12px; font-weight: 600; cursor: pointer; flex: 1;
      transition: background 0.1s;
    }
    #submit-btn:hover { background: var(--btn-hover); }
    #submit-btn:disabled { background: var(--btn-sec-bg); color: var(--mu); cursor: not-allowed; opacity: 0.6; }
    #status-bar { padding: 4px 12px; font-size: 11px; color: var(--mu); border-top: 1px solid var(--bd); background: var(--sf); }
    .image-badge { display: inline-block; background: var(--badge-bg); color: var(--badge-fg); font-size: 10px; padding: 1px 8px; border-radius: 10px; margin-top: 4px; }

    /* Syntax highlighting (basic) */
    .tok-kw { color: #c586c0; }
    .tok-str { color: #ce9178; }
    .tok-num { color: #b5cea8; }
    .tok-cm { color: #6a9955; font-style: italic; }
    .tok-fn { color: #dcdcaa; }
    .tok-type { color: #4ec9b0; }
    .tok-op { color: #d4d4d4; }
    .tok-const { color: #4fc1ff; }

    @keyframes msgSlideIn {
      from { opacity: 0; transform: translateY(8px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @keyframes typingBounce {
      0%, 80%, 100% { transform: scale(0.6); opacity: 0.3; }
      40% { transform: scale(1); opacity: 1; }
    }
  </style>
</head>
<body>
  <div id="app">
    <div id="login-screen">
      <h2>tsifl</h2>
      <div id="login-tagline">AI for Financial Analysts</div>
      <div id="login-form">
        <input type="email" id="auth-email" placeholder="Email"/>
        <input type="password" id="auth-password" placeholder="Password"/>
        <button id="login-btn">Sign In</button>
        <button id="signup-btn">Create Account</button>
        <div id="auth-error"></div>
      </div>
    </div>
    <div id="chat-screen">
      <div id="header-bar">
        <span class="title">tsifl</span>
        <div class="hdr-btns">
          <button class="hdr-btn" id="notes-btn">Notes</button>
          <span id="tasks-remaining" style="font-size:10px;color:var(--mu);"></span>
        </div>
      </div>
      <div id="user-bar"></div>
      <div id="quick-actions">
        <button class="quick-btn" data-prompt="Explain this code">Explain</button>
        <button class="quick-btn" data-prompt="Fix the errors in this code">Fix</button>
        <button class="quick-btn" data-prompt="Refactor this code">Refactor</button>
        <button class="quick-btn" data-prompt="Write tests for this code using the project's testing framework">Test</button>
        <button class="quick-btn" data-prompt="Generate JSDoc/docstring documentation for all functions in this file">Docs</button>
        <button class="quick-btn" data-prompt="Sort and organize the import statements in this file. Remove unused imports. Group by built-in, external, local.">Imports</button>
      </div>
      <div id="chat-history"></div>
      <div id="input-area">
        <div id="image-preview-bar"></div>
        <textarea id="user-input" placeholder="Ask about code, fix errors, generate tests..." rows="2"></textarea>
        <div id="input-actions">
          <input type="file" id="image-input" accept="image/*,.pdf,.csv,.txt,.json,.xml,.r,.R,.py,.js,.ts,.sql,.md,.html,.yaml,.yml,.docx,.xlsx,.sas,.do,.log" multiple style="display:none;"/>
          <button id="attach-btn" title="Attach file">+</button>
          <button id="submit-btn">Send</button>
        </div>
      </div>
      <div id="status-bar">Ready</div>
    </div>
  </div>
  <script>
    const vscode = acquireVsCodeApi();
    const SB_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
    const SB_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";
    const API = "https://focused-solace-production-6839.up.railway.app";
    let currentUser = null, pendingImages = [];
    const state = vscode.getState() || {};

    async function sbAuth(ep, body) {
      const r = await fetch(SB_URL + "/auth/v1/" + ep, { method: "POST", headers: { "Content-Type": "application/json", "apikey": SB_KEY }, body: JSON.stringify(body) });
      return r.json();
    }
    async function syncSession(s) {
      if (!s || !s.access_token) return;
      try { await fetch(API + "/auth/set-session", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ access_token: s.access_token, refresh_token: s.refresh_token, user_id: s.user?.id || "", email: s.user?.email || "" }) }); } catch(e) {}
    }
    async function initSession() {
      try {
        const resp = await fetch(API + "/auth/get-session");
        const data = await resp.json();
        if (data.session && data.session.refresh_token) {
          const r = await sbAuth("token?grant_type=refresh_token", { refresh_token: data.session.refresh_token });
          if (r.access_token) { vscode.setState({ session: r, email: r.user?.email || data.session.email }); syncSession(r); showChat({ id: r.user?.id || data.session.user_id, email: r.user?.email || data.session.email }); return; }
        }
      } catch(e) {}
      if (state.session && state.session.refresh_token) {
        try {
          const r = await sbAuth("token?grant_type=refresh_token", { refresh_token: state.session.refresh_token });
          if (r.access_token) { vscode.setState({ session: r, email: r.user?.email }); syncSession(r); showChat({ id: r.user?.id, email: r.user?.email }); return; }
        } catch(e) {}
      }
      if (state.session && state.session.access_token) {
        try {
          const payload = JSON.parse(atob(state.session.access_token.split(".")[1]));
          if (Date.now() < payload.exp * 1000) {
            const email = state.session.user?.email || state.email || "";
            if (email) { showChat({ id: state.session.user?.id || "", email }); return; }
          }
        } catch(e) {}
      }
      const savedEmail = state.email || "";
      if (savedEmail) document.getElementById("auth-email").value = savedEmail;
      showLogin();
    }
    initSession();

    function showLogin() { document.getElementById("login-screen").style.display = "flex"; document.getElementById("chat-screen").style.display = "none"; }
    function showChat(u) { currentUser = u; document.getElementById("login-screen").style.display = "none"; document.getElementById("chat-screen").style.display = "flex"; document.getElementById("user-bar").textContent = (u.email || "") + " \\u00b7 VS Code"; }

    document.getElementById("login-btn").onclick = async () => {
      const email = document.getElementById("auth-email").value.trim(), pw = document.getElementById("auth-password").value, err = document.getElementById("auth-error");
      err.textContent = ""; err.style.color = "var(--error-fg)";
      if (!email || !pw) { err.textContent = "Enter email and password."; return; }
      try {
        const r = await sbAuth("token?grant_type=password", { email, password: pw });
        if (r.error || r.error_description) { err.textContent = r.error_description || r.error || "Failed"; return; }
        vscode.setState({ session: r, email: r.user.email }); syncSession(r); showChat({ id: r.user.id, email: r.user.email });
      } catch(e) { err.textContent = "Network error"; }
    };
    document.getElementById("signup-btn").onclick = async () => {
      const email = document.getElementById("auth-email").value.trim(), pw = document.getElementById("auth-password").value, err = document.getElementById("auth-error");
      err.textContent = ""; if (!email || !pw) { err.textContent = "Enter email and password."; return; }
      if (pw.length < 6) { err.textContent = "6+ characters"; return; }
      try { const r = await sbAuth("signup", { email, password: pw }); if (r.error) { err.textContent = r.error_description || "Failed"; return; } err.style.color = "var(--gn)"; err.textContent = "Check email to confirm."; } catch(e) { err.textContent = "Network error"; }
    };
    document.getElementById("auth-password").addEventListener("keydown", e => { if (e.key === "Enter") document.getElementById("login-btn").click(); });
    document.getElementById("user-bar").addEventListener("dblclick", () => { vscode.setState({ email: currentUser?.email }); fetch(API + "/auth/clear-session", { method: "POST" }).catch(() => {}); currentUser = null; document.getElementById("chat-history").innerHTML = ""; showLogin(); });

    document.getElementById("notes-btn").onclick = () => { vscode.postMessage({ type: "executeAction", action: { type: "open_notes", payload: {} }, userId: currentUser?.id }); };

    document.querySelectorAll(".quick-btn").forEach(btn => { btn.onclick = () => { document.getElementById("user-input").value = btn.dataset.prompt; handleSubmit(); }; });

    // File attachments
    document.getElementById("attach-btn").onclick = () => document.getElementById("image-input").click();
    document.getElementById("image-input").onchange = e => { for (const f of e.target.files) { const r = new FileReader(); r.onload = () => { pendingImages.push({ media_type: f.type || (f.name && f.name.match(/\\.(png|jpg|jpeg|gif|webp)$/i) ? "image/png" : "application/octet-stream"), data: r.result.split(",")[1], preview: f.type.startsWith("image/") ? r.result : null, file_name: f.name || "" }); updatePrev(); }; r.readAsDataURL(f); } e.target.value = ""; };
    document.getElementById("user-input").addEventListener("paste", e => { for (const item of (e.clipboardData || {}).items || []) { if (item.type.startsWith("image/") || item.kind === "file") { const f = item.getAsFile(); if (f) { const r = new FileReader(); r.onload = () => { pendingImages.push({ media_type: f.type || "application/octet-stream", data: r.result.split(",")[1], preview: f.type.startsWith("image/") ? r.result : null, file_name: f.name || "" }); updatePrev(); }; r.readAsDataURL(f); } } } });
    function updatePrev() {
      const bar = document.getElementById("image-preview-bar");
      if (!pendingImages.length) { bar.style.display = "none"; bar.innerHTML = ""; return; }
      bar.style.display = "flex";
      bar.innerHTML = pendingImages.map((img, i) => {
        if (img.preview && img.media_type.startsWith("image/")) {
          return '<div class="img-preview"><img src="' + img.preview + '"/><button class="img-remove" data-i="' + i + '">x</button></div>';
        } else {
          const ext = img.file_name ? img.file_name.split(".").pop().toUpperCase() : "FILE";
          return '<div class="img-preview"><div style="width:40px;height:40px;display:flex;align-items:center;justify-content:center;background:var(--btn-sec-bg);border-radius:4px;border:1px solid var(--bd);font-size:9px;font-weight:700;color:var(--btn-bg);">' + ext + '</div><button class="img-remove" data-i="' + i + '">x</button></div>';
        }
      }).join("");
      bar.querySelectorAll(".img-remove").forEach(b => b.onclick = () => { pendingImages.splice(+b.dataset.i, 1); updatePrev(); });
    }

    // Chat submission
    document.getElementById("submit-btn").onclick = handleSubmit;
    document.getElementById("user-input").addEventListener("keydown", e => {
      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
      if (e.key === "Escape") { e.target.value = ""; }
    });
    document.getElementById("user-input").addEventListener("input", function() { this.style.height = "auto"; this.style.height = Math.min(this.scrollHeight, 120) + "px"; });

    async function handleSubmit() {
      const input = document.getElementById("user-input"), msg = input.value.trim();
      if (!msg && !pendingImages.length) return;
      input.value = ""; input.style.height = "auto";
      document.getElementById("submit-btn").disabled = true;
      document.getElementById("status-bar").textContent = "Thinking...";
      appendMsg("user", msg, pendingImages.length);
      showTyping();
      const imgs = pendingImages.map(i => ({ media_type: i.media_type, data: i.data })); pendingImages = []; updatePrev();
      vscode.postMessage({ type: "chat", text: msg, userId: currentUser?.id || "", images: imgs });
    }

    function showTyping() {
      hideTyping();
      const h = document.getElementById("chat-history");
      const d = document.createElement("div"); d.id = "typing-ind"; d.className = "typing-ind";
      d.innerHTML = "<span></span><span></span><span></span>";
      h.appendChild(d); h.scrollTop = h.scrollHeight;
    }
    function hideTyping() { const el = document.getElementById("typing-ind"); if (el) el.remove(); }

    // ── Markdown Renderer ──────────────────────────────────────────────
    function escapeHtml(str) {
      return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }

    function highlightSyntax(code, lang) {
      // Basic syntax highlighting for common languages
      const l = (lang || "").toLowerCase();
      let h = escapeHtml(code);

      // Diff highlighting takes priority
      if (l === "diff") {
        return h.split("\\n").map(line => {
          if (line.startsWith("+") && !line.startsWith("+++")) return '<span class="diff-add">' + line + '</span>';
          if (line.startsWith("-") && !line.startsWith("---")) return '<span class="diff-del">' + line + '</span>';
          return line;
        }).join("\\n");
      }

      // Apply diff-style highlighting for lines starting with +/- in any language
      h = h.split("\\n").map(line => {
        if (/^\\+[^+]/.test(line)) return '<span class="diff-add">' + line + '</span>';
        if (/^-[^-]/.test(line)) return '<span class="diff-del">' + line + '</span>';
        return line;
      }).join("\\n");

      // Comments (single-line)
      h = h.replace(/((?:^|\\n)\\s*)(#[^\\n]*)/g, '$1<span class="tok-cm">$2</span>');
      h = h.replace(/((?:^|\\n)\\s*)(\\/\\/[^\\n]*)/g, '$1<span class="tok-cm">$2</span>');

      // Strings (double and single quoted, simple non-greedy)
      h = h.replace(/(&quot;(?:[^&]|&(?!quot;))*?&quot;)/g, '<span class="tok-str">$1</span>');
      h = h.replace(/(&#x27;(?:[^&]|&(?!#x27;))*?&#x27;)/g, '<span class="tok-str">$1</span>');
      h = h.replace(/('(?:[^'\\\\]|\\\\.)*?')/g, '<span class="tok-str">$1</span>');
      h = h.replace(/("(?:[^"\\\\]|\\\\.)*?")/g, '<span class="tok-str">$1</span>');

      // Numbers
      h = h.replace(/\\b(\\d+\\.?\\d*(?:e[+-]?\\d+)?)\\b/gi, '<span class="tok-num">$1</span>');

      // Keywords by language family
      const pyKw = "def|class|import|from|return|if|elif|else|for|while|try|except|finally|with|as|raise|pass|break|continue|yield|async|await|lambda|in|not|and|or|is|None|True|False|self";
      const jsKw = "function|const|let|var|return|if|else|for|while|do|switch|case|break|continue|try|catch|finally|throw|new|class|extends|import|export|from|default|async|await|yield|typeof|instanceof|this|super|null|undefined|true|false|of|in";
      const sqlKw = "SELECT|FROM|WHERE|JOIN|LEFT|RIGHT|INNER|OUTER|ON|AND|OR|NOT|IN|IS|NULL|GROUP|BY|ORDER|ASC|DESC|HAVING|LIMIT|OFFSET|INSERT|INTO|VALUES|UPDATE|SET|DELETE|CREATE|TABLE|ALTER|DROP|INDEX|PRIMARY|KEY|FOREIGN|REFERENCES|DISTINCT|AS|UNION|ALL|EXISTS|BETWEEN|LIKE|CASE|WHEN|THEN|ELSE|END|COUNT|SUM|AVG|MIN|MAX|CAST";
      const rKw = "function|if|else|for|while|repeat|return|next|break|in|TRUE|FALSE|NULL|NA|Inf|NaN|library|require|source";
      const goKw = "func|package|import|return|if|else|for|range|switch|case|default|break|continue|go|defer|select|chan|map|struct|interface|type|var|const|nil|true|false|make|new|append|len|cap|error";

      let kwPattern;
      if (["python", "py"].includes(l)) kwPattern = pyKw;
      else if (["javascript", "js", "typescript", "ts", "jsx", "tsx"].includes(l)) kwPattern = jsKw;
      else if (["sql", "mysql", "postgresql", "sqlite"].includes(l)) kwPattern = sqlKw;
      else if (["r"].includes(l)) kwPattern = rKw;
      else if (["go", "golang"].includes(l)) kwPattern = goKw;
      else kwPattern = jsKw + "|" + pyKw; // fallback: combine common keywords

      try {
        const kwRe = new RegExp("\\\\b(" + kwPattern + ")\\\\b", "g");
        h = h.replace(kwRe, '<span class="tok-kw">$1</span>');
      } catch(e) {}

      // Function calls: word followed by (
      h = h.replace(/\\b([a-zA-Z_]\\w*)\\s*(?=\\()/g, '<span class="tok-fn">$1</span>');

      return h;
    }

    let codeBlockCounter = 0;

    function renderMd(text) {
      if (!text) return "";

      // Extract code blocks first, replace with placeholders
      const codeBlocks = [];
      let processed = text.replace(/\`\`\`(\\w*)\\n?([\\s\\S]*?)\`\`\`/g, function(_, lang, code) {
        const idx = codeBlocks.length;
        codeBlocks.push({ lang: lang || "", code: code });
        return "%%CODEBLOCK_" + idx + "%%";
      });

      // Escape HTML in the remaining text
      processed = escapeHtml(processed);

      // Bold: **text**
      processed = processed.replace(/\\*\\*(.+?)\\*\\*/g, "<strong>$1</strong>");

      // Italic: *text* (but not inside already-processed bold)
      processed = processed.replace(/(?<!\\*)\\*(?!\\*)(.+?)(?<!\\*)\\*(?!\\*)/g, "<em>$1</em>");

      // Inline code: \`code\`
      processed = processed.replace(/\`([^\`\\n]+?)\`/g, '<code>$1</code>');

      // Horizontal rules
      processed = processed.replace(/^---$/gm, "<hr>");

      // Ordered lists
      processed = processed.replace(/^(\\d+)\\. (.+)$/gm, "<li style='list-style:decimal;margin-left:20px;margin-top:2px;'>$2</li>");

      // Unordered lists
      processed = processed.replace(/^[\\-\\*] (.+)$/gm, "<li style='list-style:disc;margin-left:20px;margin-top:2px;'>$1</li>");

      // Headers (within messages)
      processed = processed.replace(/^### (.+)$/gm, '<div style="font-weight:700;font-size:13px;margin:10px 0 4px 0;">$1</div>');
      processed = processed.replace(/^## (.+)$/gm, '<div style="font-weight:700;font-size:14px;margin:12px 0 4px 0;">$1</div>');
      processed = processed.replace(/^# (.+)$/gm, '<div style="font-weight:700;font-size:15px;margin:14px 0 6px 0;">$1</div>');

      // Line breaks (double newline = paragraph break, single = <br>)
      processed = processed.replace(/\\n\\n/g, '</p><p style="margin:8px 0;">');
      processed = processed.replace(/\\n/g, "<br>");

      // Restore code blocks with full rendering
      for (let i = 0; i < codeBlocks.length; i++) {
        const { lang, code } = codeBlocks[i];
        const blockId = "codeblock_" + (++codeBlockCounter);
        const highlighted = highlightSyntax(code.trim(), lang);
        const langLabel = lang ? lang.toLowerCase() : "code";
        const block = '<div class="code-block-wrap">' +
          '<div class="code-block-header">' +
            '<span class="code-block-lang">' + escapeHtml(langLabel) + '</span>' +
            '<button class="code-copy-btn" data-block="' + blockId + '" onclick="copyCodeBlock(this, \\'' + blockId + '\\')">Copy</button>' +
          '</div>' +
          '<pre class="code-block-pre"><code id="' + blockId + '" class="code-block-code">' + highlighted + '</code></pre>' +
        '</div>';
        processed = processed.replace("%%CODEBLOCK_" + i + "%%", block);
      }

      return '<p style="margin:0;">' + processed + '</p>';
    }

    // Global copy function accessible from onclick
    function copyCodeBlock(btn, blockId) {
      const el = document.getElementById(blockId);
      if (!el) return;
      const text = el.textContent || el.innerText;
      navigator.clipboard.writeText(text).then(() => {
        btn.textContent = "Copied!";
        btn.classList.add("copied");
        setTimeout(() => { btn.textContent = "Copy"; btn.classList.remove("copied"); }, 2000);
      }).catch(() => {
        // Fallback for older webview contexts
        const ta = document.createElement("textarea");
        ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta); ta.select();
        try { document.execCommand("copy"); btn.textContent = "Copied!"; btn.classList.add("copied"); setTimeout(() => { btn.textContent = "Copy"; btn.classList.remove("copied"); }, 2000); } catch(e) {}
        document.body.removeChild(ta);
      });
    }
    window.copyCodeBlock = copyCodeBlock;

    // ── Message Listener ──────────────────────────────────────────────
    window.addEventListener("message", event => {
      const m = event.data;
      if (m.type === "chatResponse") {
        hideTyping();
        if (m.error) {
          appendMsg("error", "Error: " + m.error);
        } else {
          if (m.result.tasks_remaining >= 0) document.getElementById("tasks-remaining").textContent = m.result.tasks_remaining + " tasks";
          appendMsg("assistant", m.result.reply);
          const acts = m.result.actions?.length ? m.result.actions : (m.result.action?.type ? [m.result.action] : []);
          if (acts.length) {
            appendMsg("action", "Executing " + acts.length + " action(s): " + acts.map(a => a.type).join(", "));
          }
        }
        document.getElementById("submit-btn").disabled = false;
        document.getElementById("status-bar").textContent = "Ready";
      } else if (m.type === "actionComplete") {
        const success = m.success !== false;
        const cls = success ? "action" : "action action-fail";
        const icon = success ? "\\u2713" : "\\u2717";
        const msg = m.action + ": " + (success ? (m.message || "done") : "failed - " + (m.error || ""));
        const h = document.getElementById("chat-history"), d = document.createElement("div");
        d.className = "msg " + cls;
        d.innerHTML = '<span class="action-icon">' + icon + '</span> ' + escapeHtml(msg);
        h.appendChild(d); h.scrollTop = h.scrollHeight;
      } else if (m.type === "sendPrompt") {
        document.getElementById("user-input").value = m.prompt; handleSubmit();
      }
    });

    function appendMsg(role, text, imgCount) {
      const h = document.getElementById("chat-history"), d = document.createElement("div");
      d.className = "msg " + role;
      if ((role === "assistant" || role === "error") && text) {
        d.innerHTML = renderMd(text);
      } else if (role === "action") {
        d.innerHTML = '<span class="action-icon">\\u25B6</span> ' + escapeHtml(text || "");
      } else {
        d.textContent = text || "";
      }
      if (imgCount > 0) {
        const b = document.createElement("div"); b.className = "image-badge";
        b.textContent = imgCount + " file" + (imgCount > 1 ? "s" : "") + " attached";
        d.appendChild(b);
      }
      h.appendChild(d); h.scrollTop = h.scrollHeight;
    }
  </script>
</body>
</html>`;
  }
}

function deactivate() {}
module.exports = { activate, deactivate };
