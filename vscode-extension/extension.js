/**
 * tsifl VS Code Extension
 * Provides an AI chat sidebar with full VS Code context awareness.
 * Supports: code explanation, error fixing, refactoring, test generation,
 * file operations, terminal commands, and cross-app launch.
 */

const vscode = require("vscode");

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";

let globalPanel = null;
let providerRef = null;

function activate(context) {
  const provider = new TsiflSidebarProvider(context.extensionUri);
  providerRef = provider;

  // Register sidebar webview provider
  context.subscriptions.push(
    vscode.window.registerWebviewViewProvider("tsifl.chatView", provider, {
      webviewOptions: { retainContextWhenHidden: true },
    })
  );

  // Open Chat — opens as a panel tab (more reliable than sidebar webview)
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.openChat", () => {
      openChatPanel(context, provider);
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.explainCode", () => {
      ensurePanelAndSend(context, provider, "Explain this code");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.fixError", () => {
      ensurePanelAndSend(context, provider, "Fix the error in this code");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.refactor", () => {
      ensurePanelAndSend(context, provider, "Refactor this code");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.generateTests", () => {
      ensurePanelAndSend(context, provider, "Generate tests for this code");
    })
  );

  // Auto-open the panel on first activation if sidebar webview fails
  setTimeout(() => {
    if (!provider._view && !globalPanel) {
      openChatPanel(context, provider);
    }
  }, 2500);
}

function ensurePanelAndSend(context, provider, prompt) {
  const webview = globalPanel?.webview || provider._view?.webview;
  if (webview) {
    webview.postMessage({ type: "sendPrompt", prompt });
  } else {
    // Panel not open yet — open it, then send prompt after it loads
    openChatPanel(context, provider);
    setTimeout(() => {
      const wv = globalPanel?.webview || provider._view?.webview;
      if (wv) wv.postMessage({ type: "sendPrompt", prompt });
    }, 1000);
  }
}

function openChatPanel(context, provider) {
  if (globalPanel) {
    globalPanel.reveal(vscode.ViewColumn.Beside);
    return;
  }
  globalPanel = vscode.window.createWebviewPanel(
    "tsifl.chat",
    "tsifl",
    vscode.ViewColumn.Beside,
    { enableScripts: true, retainContextWhenHidden: true }
  );
  globalPanel.webview.html = provider._getHtml(globalPanel.webview);

  // Wire up message handling
  globalPanel.webview.onDidReceiveMessage(async (message) => {
    switch (message.type) {
      case "chat":
        await provider._handleChat(message, globalPanel.webview);
        break;
      case "getContext":
        await provider._sendContext(globalPanel.webview);
        break;
      case "executeAction":
        await provider._executeAction(message.action, globalPanel.webview);
        break;
    }
  });

  globalPanel.onDidDispose(() => { globalPanel = null; });
}

class TsiflSidebarProvider {
  constructor(extensionUri) {
    this._extensionUri = extensionUri;
    this._view = null;
  }

  resolveWebviewView(webviewView) {
    this._view = webviewView;

    webviewView.webview.options = {
      enableScripts: true,
      localResourceRoots: [this._extensionUri],
    };

    webviewView.webview.html = this._getHtml(webviewView.webview);

    webviewView.webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case "chat":
          await this._handleChat(message, webviewView.webview);
          break;
        case "getContext":
          await this._sendContext(webviewView.webview);
          break;
        case "executeAction":
          await this._executeAction(message.action, webviewView.webview);
          break;
      }
    });
  }

  _getWebview(override) {
    return override || globalPanel?.webview || this._view?.webview;
  }

  async _handleChat(message, webview) {
    const wv = this._getWebview(webview);
    try {
      const context = await this._getVSCodeContext();

      const chatBody = JSON.stringify({
        user_id: message.userId,
        message: message.text,
        context,
        session_id: `vscode-${Date.now()}`,
        images: message.images || [],
      });

      // Fetch with timeout (90s)
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 90000);
      const resp = await fetch(`${BACKEND_URL}/chat/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: chatBody,
        signal: controller.signal,
      });
      clearTimeout(timeout);

      if (!resp.ok) {
        const err = await resp.json().catch(() => ({ detail: resp.statusText }));
        wv?.postMessage({ type: "chatResponse", error: err.detail || "Request failed" });
        return;
      }

      const result = await resp.json();
      wv?.postMessage({ type: "chatResponse", result });

      // Execute actions
      const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
      for (const action of actions) {
        await this._executeAction(action, wv);
      }
    } catch (e) {
      wv?.postMessage({ type: "chatResponse", error: e.message });
    }
  }

  async _getVSCodeContext() {
    const editor = vscode.window.activeTextEditor;
    const context = {
      app: "vscode",
      workspace: vscode.workspace.name || "",
      open_files: vscode.window.visibleTextEditors.map(e => e.document.fileName),
    };

    if (editor) {
      const doc = editor.document;
      context.current_file = doc.fileName;
      context.language = doc.languageId;
      context.line_count = doc.lineCount;

      const selection = editor.selection;
      if (!selection.isEmpty) {
        context.selection = doc.getText(selection);
      }

      const visibleRanges = editor.visibleRanges;
      if (visibleRanges.length > 0) {
        context.visible_text = doc.getText(visibleRanges[0]).substring(0, 3000);
      }

      context.file_content = doc.getText().substring(0, 5000);
    }

    // Diagnostics
    const diagnostics = [];
    vscode.languages.getDiagnostics().forEach(([uri, diags]) => {
      diags.forEach((d) => {
        if (d.severity <= vscode.DiagnosticSeverity.Warning) {
          diagnostics.push({
            file: uri.fsPath,
            line: d.range.start.line + 1,
            severity: d.severity === 0 ? "error" : "warning",
            message: d.message,
          });
        }
      });
    });
    context.diagnostics = diagnostics.slice(0, 20);

    // Git
    try {
      const gitExt = vscode.extensions.getExtension("vscode.git");
      if (gitExt?.isActive) {
        const api = gitExt.exports.getAPI(1);
        if (api.repositories.length > 0) {
          const repo = api.repositories[0];
          context.git_branch = repo.state.HEAD?.name || "";
          context.git_changes = repo.state.workingTreeChanges.length;
        }
      }
    } catch (_) {}

    return context;
  }

  async _sendContext(webview) {
    const wv = this._getWebview(webview);
    const context = await this._getVSCodeContext();
    wv?.postMessage({ type: "context", context });
  }

  async _executeAction(action, webview) {
    const { type, payload } = action;
    if (!type || !payload) return;

    const wv = this._getWebview(webview);

    try {
      switch (type) {
        case "insert_code": {
          const editor = vscode.window.activeTextEditor;
          if (editor) {
            const code = payload.code || payload.text || "";
            await editor.edit((editBuilder) => {
              editBuilder.insert(editor.selection.active, code);
            });
            wv?.postMessage({ type: "actionComplete", action: type, success: true, message: `Inserted ${code.split('\n').length} lines` });
          }
          return;
        }

        case "replace_selection": {
          const editor = vscode.window.activeTextEditor;
          if (editor) {
            const code = payload.code || payload.text || "";
            if (!editor.selection.isEmpty) {
              await editor.edit((editBuilder) => {
                editBuilder.replace(editor.selection, code);
              });
              wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Selection replaced" });
            } else {
              // If no selection, insert at cursor
              await editor.edit((editBuilder) => {
                editBuilder.insert(editor.selection.active, code);
              });
              wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Code inserted at cursor" });
            }
          }
          return;
        }

        case "create_file": {
          const filePath = payload.path;
          // Resolve relative paths against workspace
          let fullPath = filePath;
          if (!filePath.startsWith("/") && vscode.workspace.workspaceFolders?.length) {
            fullPath = vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, filePath).fsPath;
          }
          const uri = vscode.Uri.file(fullPath);
          const content = Buffer.from(payload.content || "", "utf8");
          await vscode.workspace.fs.writeFile(uri, content);
          const doc = await vscode.workspace.openTextDocument(uri);
          await vscode.window.showTextDocument(doc);
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: `Created ${filePath}` });
          return;
        }

        case "edit_file": {
          const filePath = payload.path;
          let fullPath = filePath;
          if (!filePath.startsWith("/") && vscode.workspace.workspaceFolders?.length) {
            fullPath = vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, filePath).fsPath;
          }
          const uri = vscode.Uri.file(fullPath);
          const doc = await vscode.workspace.openTextDocument(uri);
          const editor = await vscode.window.showTextDocument(doc);
          if (payload.find && payload.replace !== undefined) {
            const text = doc.getText();
            const idx = text.indexOf(payload.find);
            if (idx !== -1) {
              const start = doc.positionAt(idx);
              const end = doc.positionAt(idx + payload.find.length);
              await editor.edit((editBuilder) => {
                editBuilder.replace(new vscode.Range(start, end), payload.replace);
              });
              wv?.postMessage({ type: "actionComplete", action: type, success: true, message: `Edited ${filePath}` });
            } else {
              wv?.postMessage({ type: "actionComplete", action: type, success: false, message: "Text not found in file" });
            }
          }
          return;
        }

        case "run_terminal_command":
        case "run_shell_command": {
          const terminal = vscode.window.activeTerminal || vscode.window.createTerminal("tsifl");
          terminal.show();
          terminal.sendText(payload.command);
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: `Running: ${payload.command}` });
          return;
        }

        case "open_file": {
          const filePath = payload.path;
          let fullPath = filePath;
          if (!filePath.startsWith("/") && vscode.workspace.workspaceFolders?.length) {
            fullPath = vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, filePath).fsPath;
          }
          const uri = vscode.Uri.file(fullPath);
          const doc = await vscode.workspace.openTextDocument(uri);
          await vscode.window.showTextDocument(doc);
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: `Opened ${filePath}` });
          return;
        }

        case "show_diff": {
          // Use VS Code's built-in diff editor
          const beforeUri = vscode.Uri.parse("untitled:Before");
          const afterUri = vscode.Uri.parse("untitled:After");
          try {
            // Create temp documents via output channel as fallback
            const channel = vscode.window.createOutputChannel("tsifl Diff");
            channel.clear();
            channel.appendLine("=== BEFORE ===");
            channel.appendLine(payload.before || "");
            channel.appendLine("\n=== AFTER ===");
            channel.appendLine(payload.after || "");
            channel.show();
          } catch (_) {}
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Diff shown" });
          return;
        }

        case "launch_app": {
          // Request backend to open a local app
          try {
            const resp = await fetch(`${BACKEND_URL}/launch-app`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ app_name: payload.app_name }),
            });
            const result = await resp.json();
            wv?.postMessage({ type: "actionComplete", action: type, success: true, message: result.message || "Launched" });
          } catch (e) {
            wv?.postMessage({ type: "actionComplete", action: type, success: false, message: e.message });
          }
          return;
        }

        case "open_notes": {
          const notesUrl = `${BACKEND_URL}/notes-app`;
          vscode.env.openExternal(vscode.Uri.parse(notesUrl));
          wv?.postMessage({ type: "actionComplete", action: type, success: true, message: "Opened Notes" });
          return;
        }

        case "open_url": {
          if (payload.url) {
            vscode.env.openExternal(vscode.Uri.parse(payload.url));
            wv?.postMessage({ type: "actionComplete", action: type, success: true, message: `Opened ${payload.url}` });
          }
          return;
        }

        case "explain_code":
        case "fix_error":
        case "refactor":
        case "generate_tests":
          // These are text responses — handled by chat reply
          return;

        default:
          console.log("tsifl: Unknown action type:", type);
      }

      wv?.postMessage({ type: "actionComplete", action: type, success: true });
    } catch (e) {
      wv?.postMessage({
        type: "actionComplete",
        action: type,
        success: false,
        error: e.message,
      });
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
      --blue-hover: #0A4896;
      --blue-light: #EBF3FB;
      --bg: var(--vscode-editor-background, #FFFFFF);
      --surface: var(--vscode-sideBar-background, #F8FAFC);
      --border: var(--vscode-panel-border, #E2E8F0);
      --text: var(--vscode-foreground, #1E293B);
      --text-muted: var(--vscode-descriptionForeground, #64748B);
      --green: #16A34A;
      --red: #DC2626;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: var(--vscode-font-family, -apple-system, sans-serif); font-size: 13px; background: var(--bg); color: var(--text); height: 100vh; }
    #app { display: flex; flex-direction: column; height: 100vh; }

    /* Login */
    #login-screen { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; padding: 20px; gap: 8px; }
    #login-screen h2 { color: var(--blue); font-size: 18px; margin-bottom: 4px; }
    #login-tagline { font-size: 11px; color: var(--text-muted); margin-bottom: 20px; }
    #login-form { display: flex; flex-direction: column; gap: 8px; width: 100%; }
    #login-form input { width: 100%; background: var(--surface); color: var(--text); border: 1px solid var(--border); border-radius: 4px; padding: 8px 10px; font-size: 13px; outline: none; }
    #login-form input:focus { border-color: var(--blue); }
    #login-btn { background: var(--blue); color: white; border: none; border-radius: 4px; padding: 8px; font-size: 13px; font-weight: 600; cursor: pointer; }
    #login-btn:hover { background: var(--blue-hover); }
    #signup-btn { background: transparent; color: var(--blue); border: 1px solid var(--border); border-radius: 4px; padding: 7px; font-size: 13px; cursor: pointer; }
    #auth-error { font-size: 11px; color: var(--red); text-align: center; min-height: 14px; }

    /* Chat */
    #chat-screen { display: none; flex-direction: column; height: 100vh; }
    #user-bar { font-size: 10px; color: var(--text-muted); padding: 4px 10px; background: var(--surface); border-bottom: 1px solid var(--border); cursor: default; }
    #chat-history { flex: 1; overflow-y: auto; padding: 8px; display: flex; flex-direction: column; gap: 6px; }
    #chat-history::-webkit-scrollbar { width: 4px; }
    #chat-history::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
    .msg { padding: 7px 10px; border-radius: 4px; line-height: 1.5; word-wrap: break-word; font-size: 12px; white-space: pre-wrap; }
    .msg.user { background: var(--blue-light); border-left: 2px solid var(--blue); }
    .msg.assistant { background: var(--bg); border-left: 2px solid #86EFAC; }
    .msg.assistant strong { font-weight: 600; }
    .msg.assistant code { background: var(--surface); padding: 1px 4px; border-radius: 3px; font-size: 11px; font-family: var(--vscode-editor-font-family, monospace); }
    .msg.action { background: rgba(22,163,74,0.08); border-left: 2px solid var(--green); font-size: 11px; color: var(--green); font-family: monospace; padding: 5px 10px; }

    /* Input */
    #input-area { padding: 8px; border-top: 1px solid var(--border); display: flex; flex-direction: column; gap: 4px; }
    #image-preview-bar { display: none; gap: 4px; flex-wrap: wrap; }
    .img-preview { position: relative; display: inline-block; }
    .img-preview img { width: 40px; height: 40px; object-fit: cover; border-radius: 4px; border: 1px solid var(--border); }
    .img-remove { position: absolute; top: -3px; right: -3px; width: 14px; height: 14px; background: var(--red); color: white; border: none; border-radius: 50%; font-size: 9px; line-height: 14px; text-align: center; cursor: pointer; padding: 0; }
    #user-input { width: 100%; background: var(--surface); color: var(--text); border: 1px solid var(--border); border-radius: 4px; padding: 6px 8px; font-size: 12px; font-family: inherit; resize: none; outline: none; line-height: 1.4; }
    #user-input:focus { border-color: var(--blue); }
    #input-actions { display: flex; gap: 4px; }
    #attach-btn { background: var(--surface); color: var(--text-muted); border: 1px solid var(--border); border-radius: 4px; padding: 6px 10px; font-size: 14px; cursor: pointer; flex-shrink: 0; }
    #submit-btn { background: var(--blue); color: white; border: none; border-radius: 4px; padding: 6px; font-size: 12px; font-weight: 600; cursor: pointer; flex: 1; }
    #submit-btn:disabled { background: var(--surface); color: var(--text-muted); cursor: not-allowed; }
    #status-bar { padding: 3px 10px; font-size: 10px; color: var(--text-muted); border-top: 1px solid var(--border); }
    .image-badge { display: inline-block; background: var(--blue-light); color: var(--blue); font-size: 10px; padding: 1px 6px; border-radius: 8px; margin-top: 4px; }
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
      <div id="user-bar"></div>
      <div id="chat-history"></div>
      <div id="input-area">
        <div id="image-preview-bar"></div>
        <textarea id="user-input" placeholder="Ask about code, fix errors, generate tests..." rows="3"></textarea>
        <div id="input-actions">
          <input type="file" id="image-input" accept="image/*" multiple style="display:none;"/>
          <button id="attach-btn" title="Attach image">+</button>
          <button id="submit-btn">Send</button>
        </div>
      </div>
      <div id="status-bar">Ready</div>
    </div>
  </div>

  <script>
    const vscode = acquireVsCodeApi();
    const SUPABASE_URL = "https://dvynmzeyttwlmvunicqz.supabase.co";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR2eW5temV5dHR3bG12dW5pY3F6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ2NTIwMTIsImV4cCI6MjA5MDIyODAxMn0.9j_f-2f1VswxWfqiuXy4bPnUi1qLk9nAeTDlodUBUZw";

    let currentUser = null;
    let pendingImages = [];
    const state = vscode.getState() || {};

    const BACKEND_AUTH_URL = "https://focused-solace-production-6839.up.railway.app";

    async function syncSessionToBackend(session) {
      if (!session || !session.access_token) return;
      try {
        await fetch(BACKEND_AUTH_URL + "/auth/set-session", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            access_token: session.access_token,
            refresh_token: session.refresh_token,
            user_id: session.user?.id || "",
            email: session.user?.email || "",
          }),
        });
      } catch (e) {}
    }

    async function restoreFromBackend() {
      try {
        const resp = await fetch(BACKEND_AUTH_URL + "/auth/get-session");
        const data = await resp.json();
        if (!data.session || !data.session.access_token) return false;
        const result = await supabaseAuth("token?grant_type=refresh_token", { refresh_token: data.session.refresh_token });
        if (result.access_token) {
          vscode.setState({ session: result });
          syncSessionToBackend(result);
          showChat({ id: result.user?.id || data.session.user_id, email: result.user?.email || data.session.email });
          return true;
        }
      } catch (e) {}
      return false;
    }

    async function initSession() {
      if (state.session) {
        const ok = await refreshSession(state.session);
        if (ok) return;
      }
      const restored = await restoreFromBackend();
      if (!restored) showLogin();
    }
    initSession();

    async function supabaseAuth(endpoint, body) {
      const resp = await fetch(SUPABASE_URL + "/auth/v1/" + endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", "apikey": SUPABASE_ANON_KEY },
        body: JSON.stringify(body),
      });
      return resp.json();
    }

    async function refreshSession(session) {
      try {
        const result = await supabaseAuth("token?grant_type=refresh_token", { refresh_token: session.refresh_token });
        if (result.access_token) {
          vscode.setState({ session: result });
          syncSessionToBackend(result);
          showChat({ id: result.user?.id || session.user?.id, email: result.user?.email || session.user?.email });
          return true;
        }
      } catch (e) {}
      showLogin();
      return false;
    }

    document.getElementById("login-btn").onclick = async () => {
      const email = document.getElementById("auth-email").value.trim();
      const password = document.getElementById("auth-password").value;
      const errEl = document.getElementById("auth-error");
      errEl.textContent = "";
      errEl.style.color = "var(--red)";
      if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
      const result = await supabaseAuth("token?grant_type=password", { email, password });
      if (result.error || result.error_description) { errEl.textContent = result.error_description || "Sign in failed"; return; }
      vscode.setState({ session: result });
      syncSessionToBackend(result);
      showChat({ id: result.user.id, email: result.user.email });
    };

    document.getElementById("signup-btn").onclick = async () => {
      const email = document.getElementById("auth-email").value.trim();
      const password = document.getElementById("auth-password").value;
      const errEl = document.getElementById("auth-error");
      errEl.textContent = "";
      errEl.style.color = "var(--red)";
      if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
      if (password.length < 6) { errEl.textContent = "Password must be 6+ characters."; return; }
      const result = await supabaseAuth("signup", { email, password });
      if (result.error) { errEl.textContent = result.error_description || "Sign up failed"; return; }
      errEl.style.color = "#16A34A";
      errEl.textContent = "Check email to confirm, then sign in.";
    };

    // Password enter to sign in
    document.getElementById("auth-password").addEventListener("keydown", (e) => {
      if (e.key === "Enter") document.getElementById("login-btn").click();
    });

    function showLogin() {
      document.getElementById("login-screen").style.display = "flex";
      document.getElementById("chat-screen").style.display = "none";
    }

    function showChat(user) {
      currentUser = user;
      document.getElementById("login-screen").style.display = "none";
      document.getElementById("chat-screen").style.display = "flex";
      document.getElementById("user-bar").textContent = user.email + " \\u00b7 VS Code";
    }

    // Sign out on double-click user bar
    document.getElementById("user-bar").addEventListener("dblclick", () => {
      vscode.setState({});
      fetch(BACKEND_AUTH_URL + "/auth/clear-session", { method: "POST" }).catch(() => {});
      currentUser = null;
      document.getElementById("chat-history").innerHTML = "";
      showLogin();
    });

    // Image handling
    document.getElementById("attach-btn").onclick = () => document.getElementById("image-input").click();
    document.getElementById("image-input").onchange = (e) => {
      for (const file of e.target.files) {
        const reader = new FileReader();
        reader.onload = () => {
          pendingImages.push({ media_type: file.type || "image/png", data: reader.result.split(",")[1], preview: reader.result });
          updatePreview();
        };
        reader.readAsDataURL(file);
      }
      e.target.value = "";
    };

    document.getElementById("user-input").addEventListener("paste", (e) => {
      for (const item of (e.clipboardData || {}).items || []) {
        if (item.type.startsWith("image/")) {
          const file = item.getAsFile();
          if (file) {
            const reader = new FileReader();
            reader.onload = () => {
              pendingImages.push({ media_type: file.type, data: reader.result.split(",")[1], preview: reader.result });
              updatePreview();
            };
            reader.readAsDataURL(file);
          }
        }
      }
    });

    function updatePreview() {
      const bar = document.getElementById("image-preview-bar");
      if (pendingImages.length === 0) { bar.style.display = "none"; bar.innerHTML = ""; return; }
      bar.style.display = "flex";
      bar.innerHTML = pendingImages.map((img, i) =>
        '<div class="img-preview"><img src="' + img.preview + '"/><button class="img-remove" data-i="' + i + '">x</button></div>'
      ).join("");
      bar.querySelectorAll(".img-remove").forEach(b => b.onclick = () => { pendingImages.splice(+b.dataset.i, 1); updatePreview(); });
    }

    // Submit
    document.getElementById("submit-btn").onclick = handleSubmit;
    document.getElementById("user-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSubmit(); }
    });

    async function handleSubmit() {
      const input = document.getElementById("user-input");
      const msg = input.value.trim();
      if (!msg && pendingImages.length === 0) return;
      input.value = "";
      document.getElementById("submit-btn").disabled = true;
      document.getElementById("status-bar").textContent = "Thinking...";

      const imgCount = pendingImages.length;
      appendMsg("user", msg, imgCount);

      const images = pendingImages.map(i => ({ media_type: i.media_type, data: i.data }));
      pendingImages = [];
      updatePreview();

      vscode.postMessage({ type: "chat", text: msg, userId: currentUser.id, images });
    }

    // Render basic markdown
    function renderMarkdown(text) {
      return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\\*\\*(.+?)\\*\\*/g, "<strong>$1</strong>")
        .replace(/\\*(.+?)\\*/g, "<em>$1</em>")
        .replace(/\`([^\`]+)\`/g, '<code>$1</code>')
        .replace(/\\n/g, "<br>");
    }

    // Listen for messages from extension
    window.addEventListener("message", (event) => {
      const msg = event.data;
      switch (msg.type) {
        case "chatResponse":
          if (msg.error) { appendMsg("assistant", "Error: " + msg.error); }
          else {
            if (msg.result.tasks_remaining >= 0) {
              document.getElementById("user-bar").textContent = currentUser.email + " \\u00b7 " + msg.result.tasks_remaining + " tasks left";
            }
            appendMsg("assistant", msg.result.reply);
            const actions = msg.result.actions?.length ? msg.result.actions : (msg.result.action?.type ? [msg.result.action] : []);
            if (actions.length > 0) {
              appendMsg("action", "Executing " + actions.length + " action(s): " + actions.map(a => a.type).join(", "));
            }
          }
          document.getElementById("submit-btn").disabled = false;
          document.getElementById("status-bar").textContent = "Ready";
          break;
        case "actionComplete":
          appendMsg("action", msg.action + ": " + (msg.success ? (msg.message || "done") : "failed \\u2014 " + (msg.error || "")));
          break;
        case "sendPrompt":
          document.getElementById("user-input").value = msg.prompt;
          handleSubmit();
          break;
      }
    });

    function appendMsg(role, text, imgCount) {
      const h = document.getElementById("chat-history");
      const d = document.createElement("div");
      d.className = "msg " + role;
      if (role === "assistant" && text) {
        d.innerHTML = renderMarkdown(text);
      } else {
        d.textContent = text || "";
      }
      if (imgCount > 0) {
        const b = document.createElement("div");
        b.className = "image-badge";
        b.textContent = imgCount + " image" + (imgCount > 1 ? "s" : "") + " attached";
        d.appendChild(b);
      }
      h.appendChild(d);
      h.scrollTop = h.scrollHeight;
    }
  </script>
</body>
</html>`;
  }
}

function deactivate() {}

module.exports = { activate, deactivate };
