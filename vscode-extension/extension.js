/**
 * tsifl VS Code Extension
 * Provides an AI chat sidebar with full VS Code context awareness.
 */

const vscode = require("vscode");

const BACKEND_URL = "https://focused-solace-production-6839.up.railway.app";

let globalPanel = null;

function activate(context) {
  // Register the sidebar webview provider
  const provider = new TsiflSidebarProvider(context.extensionUri);
  context.subscriptions.push(
    vscode.window.registerWebviewViewProvider("tsifl.chatView", provider)
  );

  // Register commands
  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.openChat", () => {
      vscode.commands.executeCommand("tsifl.chatView.focus");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.explainCode", () => {
      sendContextCommand(provider, "Explain this code");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.fixError", () => {
      sendContextCommand(provider, "Fix the error in this code");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.refactor", () => {
      sendContextCommand(provider, "Refactor this code");
    })
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("tsifl.generateTests", () => {
      sendContextCommand(provider, "Generate tests for this code");
    })
  );
}

function sendContextCommand(provider, prompt) {
  if (provider._view) {
    provider._view.webview.postMessage({ type: "sendPrompt", prompt });
  }
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

    // Handle messages from the webview
    webviewView.webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case "chat":
          await this._handleChat(message);
          break;
        case "getContext":
          await this._sendContext();
          break;
        case "executeAction":
          await this._executeAction(message.action);
          break;
      }
    });
  }

  async _handleChat(message) {
    try {
      const context = await this._getVSCodeContext();

      const resp = await fetch(`${BACKEND_URL}/chat/`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          user_id: message.userId,
          message: message.text,
          context,
          session_id: `vscode-${Date.now()}`,
          images: message.images || [],
        }),
      });

      if (!resp.ok) {
        const err = await resp.json().catch(() => ({ detail: resp.statusText }));
        this._view.webview.postMessage({ type: "chatResponse", error: err.detail || "Request failed" });
        return;
      }

      const result = await resp.json();
      this._view.webview.postMessage({ type: "chatResponse", result });

      // Execute actions
      const actions = result.actions?.length ? result.actions : (result.action?.type ? [result.action] : []);
      for (const action of actions) {
        await this._executeAction(action);
      }
    } catch (e) {
      this._view.webview.postMessage({ type: "chatResponse", error: e.message });
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

      // Get selected text or surrounding context
      const selection = editor.selection;
      if (!selection.isEmpty) {
        context.selection = doc.getText(selection);
      }

      // Get visible range
      const visibleRanges = editor.visibleRanges;
      if (visibleRanges.length > 0) {
        context.visible_text = doc.getText(visibleRanges[0]).substring(0, 3000);
      }

      // Get full file (truncated)
      context.file_content = doc.getText().substring(0, 5000);
    }

    // Get diagnostics (errors/warnings)
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

    // Get git status if available
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

  async _sendContext() {
    const context = await this._getVSCodeContext();
    this._view.webview.postMessage({ type: "context", context });
  }

  async _executeAction(action) {
    const { type, payload } = action;
    if (!type || !payload) return;

    try {
      switch (type) {
        case "insert_code": {
          const editor = vscode.window.activeTextEditor;
          if (editor) {
            await editor.edit((editBuilder) => {
              editBuilder.insert(editor.selection.active, payload.code || payload.text || "");
            });
          }
          break;
        }

        case "replace_selection": {
          const editor = vscode.window.activeTextEditor;
          if (editor && !editor.selection.isEmpty) {
            await editor.edit((editBuilder) => {
              editBuilder.replace(editor.selection, payload.code || payload.text || "");
            });
          }
          break;
        }

        case "create_file": {
          const uri = vscode.Uri.file(payload.path);
          const content = Buffer.from(payload.content || "", "utf8");
          await vscode.workspace.fs.writeFile(uri, content);
          const doc = await vscode.workspace.openTextDocument(uri);
          await vscode.window.showTextDocument(doc);
          break;
        }

        case "edit_file": {
          const uri = vscode.Uri.file(payload.path);
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
            }
          }
          break;
        }

        case "run_terminal_command": {
          const terminal = vscode.window.activeTerminal || vscode.window.createTerminal("tsifl");
          terminal.show();
          terminal.sendText(payload.command);
          break;
        }

        case "open_file": {
          const uri = vscode.Uri.file(payload.path);
          const doc = await vscode.workspace.openTextDocument(uri);
          await vscode.window.showTextDocument(doc);
          break;
        }

        case "show_diff": {
          // Show before/after in output channel
          const channel = vscode.window.createOutputChannel("tsifl Diff");
          channel.clear();
          channel.appendLine("=== BEFORE ===");
          channel.appendLine(payload.before || "");
          channel.appendLine("\n=== AFTER ===");
          channel.appendLine(payload.after || "");
          channel.show();
          break;
        }

        case "explain_code":
        case "fix_error":
        case "refactor":
        case "generate_tests":
          // These are text responses — handled by chat reply
          break;

        case "run_shell_command": {
          const terminal = vscode.window.activeTerminal || vscode.window.createTerminal("tsifl");
          terminal.show();
          terminal.sendText(payload.command);
          break;
        }

        default:
          console.log("tsifl: Unknown action type:", type);
      }

      this._view.webview.postMessage({
        type: "actionComplete",
        action: type,
        success: true,
      });
    } catch (e) {
      this._view.webview.postMessage({
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
    #user-bar { font-size: 10px; color: var(--text-muted); padding: 4px 10px; background: var(--surface); border-bottom: 1px solid var(--border); }
    #chat-history { flex: 1; overflow-y: auto; padding: 8px; display: flex; flex-direction: column; gap: 6px; }
    #chat-history::-webkit-scrollbar { width: 4px; }
    #chat-history::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
    .msg { padding: 7px 10px; border-radius: 4px; line-height: 1.5; word-wrap: break-word; font-size: 12px; white-space: pre-wrap; }
    .msg.user { background: var(--blue-light); border-left: 2px solid var(--blue); }
    .msg.assistant { background: var(--bg); border-left: 2px solid #86EFAC; }
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

    // Restore session
    if (state.session) {
      refreshSession(state.session);
    }

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
          showChat({ id: result.user?.id || session.user?.id, email: result.user?.email || session.user?.email });
          return;
        }
      } catch (e) {}
      showLogin();
    }

    document.getElementById("login-btn").onclick = async () => {
      const email = document.getElementById("auth-email").value.trim();
      const password = document.getElementById("auth-password").value;
      const errEl = document.getElementById("auth-error");
      errEl.textContent = "";
      if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
      const result = await supabaseAuth("token?grant_type=password", { email, password });
      if (result.error || result.error_description) { errEl.textContent = result.error_description || "Sign in failed"; return; }
      vscode.setState({ session: result });
      showChat({ id: result.user.id, email: result.user.email });
    };

    document.getElementById("signup-btn").onclick = async () => {
      const email = document.getElementById("auth-email").value.trim();
      const password = document.getElementById("auth-password").value;
      const errEl = document.getElementById("auth-error");
      errEl.textContent = "";
      if (!email || !password) { errEl.textContent = "Enter email and password."; return; }
      if (password.length < 6) { errEl.textContent = "Password must be 6+ characters."; return; }
      const result = await supabaseAuth("signup", { email, password });
      if (result.error) { errEl.textContent = result.error_description || "Sign up failed"; return; }
      errEl.style.color = "#16A34A";
      errEl.textContent = "Check email to confirm, then sign in.";
    };

    function showLogin() {
      document.getElementById("login-screen").style.display = "flex";
      document.getElementById("chat-screen").style.display = "none";
    }

    function showChat(user) {
      currentUser = user;
      document.getElementById("login-screen").style.display = "none";
      document.getElementById("chat-screen").style.display = "flex";
      document.getElementById("user-bar").textContent = user.email;
    }

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

    // Listen for messages from extension
    window.addEventListener("message", (event) => {
      const msg = event.data;
      switch (msg.type) {
        case "chatResponse":
          if (msg.error) { appendMsg("assistant", "Error: " + msg.error); }
          else {
            if (msg.result.tasks_remaining >= 0) {
              document.getElementById("user-bar").textContent = currentUser.email + " · " + msg.result.tasks_remaining + " tasks left";
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
          appendMsg("action", msg.action + ": " + (msg.success ? "done" : "failed — " + (msg.error || "")));
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
      d.textContent = text || "";
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
