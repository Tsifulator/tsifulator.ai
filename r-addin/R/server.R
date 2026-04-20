#' tsifl Background Server
#'
#' Runs the Shiny UI as a background job so the main R session stays free.
#' When the main session is free, sendToConsole() routes plots to the Plots pane.
#'
#' @keywords internal
run_tsifl_server <- function(port = 7444) {

  # ── IPC Warmup ─────────────────────────────────────────────────────────────
  # When this background job starts, its rstudioapi IPC pipe to the main
  # RStudio IDE isn't registered until the first sendToConsole call. That
  # first call often silently times out ("RStudio did not respond"). Fire a
  # harmless warmup with retries so the pipe is live by the time a user chats.
  for (.warmup_attempt in 1:5) {
    ok <- tryCatch({
      rstudioapi::sendToConsole(
        "invisible(TRUE)",
        execute = TRUE, echo = FALSE, focus = FALSE
      )
      TRUE
    }, error = function(e) FALSE)
    if (ok) break
    Sys.sleep(0.5 * .warmup_attempt)  # 0.5, 1, 1.5, 2, 2.5s
  }

  BACKEND_URL <- if (nchar(Sys.getenv("TSIFULATOR_BACKEND_URL")) > 0) {
    Sys.getenv("TSIFULATOR_BACKEND_URL")
  } else {
    "https://focused-solace-production-6839.up.railway.app"
  }

  config_path <- path.expand("~/.tsifulator_user")
  USER_ID <- if (file.exists(config_path)) {
    trimws(readLines(config_path, n = 1, warn = FALSE))
  } else if (nchar(Sys.getenv("TSIFULATOR_USER_ID")) > 0) {
    Sys.getenv("TSIFULATOR_USER_ID")
  } else {
    "dev-user-001"
  }

  # ── CSS ────────────────────────────────────────────────────────────────────
  CSS <- "
    * { box-sizing: border-box; margin: 0; padding: 0; }

    body, .container-fluid {
      background: #FFFFFF;
      color: #1E293B;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', sans-serif;
      font-size: 14px;
      height: 100vh;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      -webkit-font-smoothing: antialiased;
      margin: 0;
      padding: 0 !important;
    }

    #header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 12px 16px;
      background: #FFFFFF;
      border-bottom: 1px solid #F0F0F0;
      flex-shrink: 0;
    }

    #logo {
      font-weight: 600;
      font-size: 15px;
      color: #0D5EAF;
      letter-spacing: -0.3px;
    }

    #tasks_label {
      font-size: 11px;
      color: #8E8E93;
      background: #F2F2F7;
      padding: 3px 10px;
      border-radius: 12px;
      font-weight: 500;
      border: none;
    }

    /* ── Tab bar — understated, Bloomberg/Linear-ish ───────────── */
    #tsifl_tabs {
      display: flex;
      align-items: stretch;
      gap: 0;
      padding: 0 14px;
      background: #FFFFFF;
      border-bottom: 1px solid #E8ECF1;
      flex-shrink: 0;
      height: 36px;
    }
    .tsifl-tab {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 0 14px;
      font-size: 12px;
      font-weight: 500;
      letter-spacing: 0.3px;
      text-transform: uppercase;
      color: #94A3B8;
      background: transparent;
      border: none;
      border-bottom: 1.5px solid transparent;
      cursor: pointer;
      transition: color 0.12s ease, border-bottom-color 0.12s ease;
      font-family: inherit;
      margin-bottom: -1px;
    }
    .tsifl-tab:hover { color: #334155; }
    .tsifl-tab.active {
      color: #0F172A;
      border-bottom-color: #0F172A;
    }
    .tsifl-tab .tab-count {
      font-size: 10px;
      font-weight: 600;
      color: #475569;
      background: #F1F5F9;
      padding: 1px 6px;
      border-radius: 10px;
      letter-spacing: 0;
      text-transform: none;
      min-width: 16px;
      text-align: center;
    }
    .tsifl-tab.active .tab-count {
      background: #0F172A;
      color: #FFFFFF;
    }
    #plot_tab {
      display: none;
      flex: 1;
      flex-direction: column;
      overflow: hidden;
      background: #FFFFFF;
    }
    #plot_tab.active { display: flex; }
    #chat_tab { display: flex; flex-direction: column; flex: 1; overflow: hidden; }
    #chat_tab.hidden { display: none !important; }
    #plot_toolbar {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 10px 14px;
      border-bottom: 1px solid #F1F5F9;
      font-size: 12px;
      background: #FCFCFD;
    }
    #plot_toolbar select {
      flex: 1;
      padding: 6px 10px;
      font-size: 12px;
      font-weight: 500;
      border: 1px solid #E2E8F0;
      border-radius: 4px;
      background: #FFFFFF;
      color: #0F172A;
      font-family: inherit;
      letter-spacing: 0.1px;
      outline: none;
      cursor: pointer;
    }
    #plot_toolbar select:hover { border-color: #CBD5E1; }
    #plot_toolbar select:focus { border-color: #0F172A; }
    #plot_toolbar button {
      padding: 6px 12px;
      font-size: 11px;
      font-weight: 500;
      letter-spacing: 0.3px;
      text-transform: uppercase;
      color: #475569;
      background: #FFFFFF;
      border: 1px solid #E2E8F0;
      border-radius: 4px;
      cursor: pointer;
      font-family: inherit;
      white-space: nowrap;
      transition: all 0.12s ease;
    }
    #plot_toolbar button:hover {
      color: #0F172A;
      border-color: #CBD5E1;
      background: #F8FAFC;
    }
    #plot_toolbar button:active { background: #F1F5F9; }
    #plot_iframe_wrap {
      flex: 1;
      padding: 10px 14px 14px 14px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }
    #plot_iframe {
      flex: 1;
      width: 100%;
      min-height: 500px;
      border: 1px solid #E8ECF1;
      border-radius: 4px;
      background: #FFFFFF;
    }
    #plot_empty_state {
      flex: 1;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 40px;
    }

    #chat_history {
      flex: 1;
      overflow-y: auto;
      padding: 16px 14px;
      padding-bottom: 140px;
      display: flex;
      flex-direction: column;
      gap: 6px;
      background: #FFFFFF;
    }

    #chat_history::-webkit-scrollbar { width: 0; }

    .msg-user {
      background: #0D5EAF;
      color: #FFFFFF;
      padding: 10px 14px;
      border-radius: 18px 18px 4px 18px;
      line-height: 1.5;
      font-size: 14px;
      align-self: flex-end;
      max-width: 82%;
      word-wrap: break-word;
      animation: msgIn 0.2s ease;
    }

    .msg-assistant {
      padding: 4px 2px;
      line-height: 1.65;
      font-size: 14px;
      color: #1D1D1F;
      max-width: 100%;
      word-wrap: break-word;
      animation: msgIn 0.2s ease;
    }

    .msg-action {
      padding: 1px 2px;
      font-size: 12px;
      color: #8E8E93;
      line-height: 1.4;
    }

    #input_area {
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      background: #FFFFFF;
      padding: 8px 12px 10px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    #user_input {
      width: 100%;
      box-sizing: border-box;
      background: #F2F2F7;
      color: #1D1D1F;
      border: none;
      border-radius: 20px;
      padding: 10px 16px;
      font-size: 14px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      resize: none;
      outline: none;
      transition: box-shadow 0.15s;
      line-height: 1.4;
    }

    #user_input:focus {
      box-shadow: 0 0 0 2px rgba(13, 94, 175, 0.2);
    }

    #user_input::placeholder { color: #8E8E93; }

    #input_actions {
      display: flex;
      gap: 6px;
    }

    #attach_btn {
      background: #F2F2F7;
      color: #8E8E93;
      border: none;
      border-radius: 20px;
      padding: 10px 14px;
      font-size: 16px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.15s;
      line-height: 1;
      flex-shrink: 0;
    }

    #attach_btn:hover {
      background: #E5E5EA;
      color: #1D1D1F;
    }

    #send_btn {
      flex: 1;
      background: #0D5EAF;
      color: white;
      border: none;
      border-radius: 20px;
      padding: 10px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s;
      letter-spacing: 0.1px;
    }

    #send_btn:hover { background: #0A4E94; }
    #send_btn:active { background: #083D7A; transform: scale(0.98); }

    #image_preview_bar {
      display: none;
      flex-wrap: wrap;
      gap: 5px;
      padding: 4px 0;
    }

    .image-preview-item {
      position: relative;
      display: inline-block;
    }

    .image-preview-item canvas {
      border-radius: 8px;
      border: 1px solid #E5E5EA;
    }

    .image-preview-item .remove-img {
      position: absolute;
      top: -4px;
      right: -4px;
      width: 18px;
      height: 18px;
      background: #FF3B30;
      color: white;
      border: none;
      border-radius: 50%;
      font-size: 11px;
      line-height: 18px;
      text-align: center;
      cursor: pointer;
      padding: 0;
    }

    .chat-canvas {
      border-radius: 10px;
      border: 1px solid #E5E5EA;
      display: block;
      max-width: 100%;
      margin-top: 6px;
    }

    .image-badge {
      display: inline-block;
      background: #F2F2F7;
      color: #8E8E93;
      font-size: 11px;
      font-weight: 500;
      padding: 3px 10px;
      border-radius: 12px;
      margin-top: 4px;
    }

    #status_bar { display: none; }

    @keyframes msgIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }
    @keyframes shimmer {
      0% { background-position: -200% 0; }
      100% { background-position: 200% 0; }
    }
    @keyframes fadeInUp {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }

    /* Claude-style thinking indicator */
    #tsifl-thinking-bubble {
      padding: 12px 4px;
      animation: fadeInUp 0.3s ease;
      display: flex;
      align-items: flex-start;
      gap: 10px;
    }
    #tsifl-thinking-bubble .thinking-phase { display: none; }
    #tsifl-thinking-bubble .thinking-orb {
      width: 20px;
      height: 20px;
      border-radius: 50%;
      background: linear-gradient(90deg, #C7C7CC 25%, #E8E8ED 50%, #C7C7CC 75%);
      background-size: 200% 100%;
      animation: shimmer 1.8s ease-in-out infinite;
      flex-shrink: 0;
      margin-top: 1px;
    }
    #tsifl-thinking-bubble .thinking-content {
      display: flex;
      flex-direction: column;
      gap: 2px;
    }
    #tsifl-thinking-bubble .thinking-text {
      font-size: 14px;
      color: #8E8E93;
      transition: opacity 0.3s ease, transform 0.3s ease;
      line-height: 1.4;
    }

    .msg-assistant pre { background: #F2F2F7; color: #1D1D1F; border-radius: 10px; padding: 10px 12px; margin: 6px 0; font-family: 'SF Mono', Menlo, Consolas, monospace; font-size: 12px; line-height: 1.5; overflow-x: auto; white-space: pre-wrap; }
    .msg-assistant code { background: #F2F2F7; padding: 2px 6px; border-radius: 5px; font-size: 12px; font-family: 'SF Mono', Menlo, Consolas, monospace; color: #1D1D1F; }
  "

  # ── UI ─────────────────────────────────────────────────────────────────────
  ui <- shiny::fluidPage(
    shiny::tags$head(
      shiny::tags$style(shiny::HTML(CSS)),
      shiny::tags$script(shiny::HTML("
        // Keep-alive ping every 10 seconds to prevent WebSocket idle disconnect
        setInterval(function() {
          if (typeof Shiny !== 'undefined' && Shiny.setInputValue) {
            Shiny.setInputValue('_keepalive', Date.now(), {priority: 'event'});
          }
        }, 10000);

        // Also use visibility change to immediately ping when tab becomes active
        document.addEventListener('visibilitychange', function() {
          if (!document.hidden && typeof Shiny !== 'undefined' && Shiny.setInputValue) {
            Shiny.setInputValue('_keepalive', Date.now(), {priority: 'event'});
          }
        });

        // Aggressive reconnect if WebSocket drops
        var _tsiflReconnecting = false;
        $(document).on('shiny:disconnected', function(event) {
          if (_tsiflReconnecting) return;
          _tsiflReconnecting = true;
          var attempts = 0;
          function tryReconnect() {
            attempts++;
            fetch('http://127.0.0.1:7444', {mode: 'no-cors'}).then(function() {
              location.reload();
            }).catch(function() {
              if (attempts < 360) { // retry for 30 min
                setTimeout(tryReconnect, 5000);
              }
            });
          }
          var chatEl = document.getElementById('chat_history');
          if (chatEl) {
            var old = document.getElementById('reconnect-notice');
            if (!old) {
              var notice = document.createElement('div');
              notice.id = 'reconnect-notice';
              notice.style.cssText = 'text-align:center;padding:12px;color:#64748B;font-size:13px;';
              notice.textContent = 'Reconnecting...';
              chatEl.appendChild(notice);
            }
          }
          setTimeout(tryReconnect, 1000);
        });
      "))
    ),

    shiny::div(id = "header",
      shiny::span(id = "logo", "\u26a1 tsifl"),
      shiny::div(style = "display:flex;align-items:center;gap:6px;",
        shiny::uiOutput("tasks_label")
      )
    ),

    # ── Tab bar: Chat | Plot ─────────────────────────────────────────
    shiny::div(id = "tsifl_tabs",
      shiny::tags$button(id = "tab_chat_btn", class = "tsifl-tab active",
                        onclick = "tsiflShowTab('chat');",
                        "Chat"),
      shiny::tags$button(id = "tab_plot_btn", class = "tsifl-tab",
                        onclick = "tsiflShowTab('plot');",
                        "Plots",
                        shiny::span(id = "plot_count", class = "tab-count",
                                    style = "display:none;", "0"))
    ),

    # ── Chat tab content (default) ───────────────────────────────────
    shiny::div(id = "chat_tab",
      shiny::div(id = "chat_history",
        shiny::uiOutput("chat_messages"),
        shiny::div(id = "thinking_container")
      ),

      shiny::div(id = "input_area",
        shiny::div(id = "image_preview_bar"),
        shiny::tags$textarea(id = "user_input",
          placeholder = "What can I help you with?",
          rows = "2"
        ),
        shiny::tags$input(type = "file", id = "image_input",
          accept = "image/*,.pdf,.csv,.txt,.json,.xml,.r,.R,.py,.js,.ts,.sql,.md,.html,.yaml,.yml,.docx,.xlsx,.sas,.do,.log",
          multiple = "multiple",
          style = "display:none;"
        ),
        shiny::div(id = "input_actions",
          shiny::tags$button(id = "attach_btn", title = "Attach file", "+"),
          shiny::tags$button(id = "send_btn", style = "width:100%;",
            onclick = "if(window._tsiflSend){window._tsiflSend();}",
            "Send")
        ),
        shiny::div(id = "status_bar", class = "idle", shiny::HTML('<span class="status-dot"></span><span id="status_text"></span>'))
      )
    ),

    # ── Plot tab content (hidden until user clicks Plot) ─────────────
    shiny::div(id = "plot_tab",
      shiny::div(id = "plot_toolbar",
        shiny::uiOutput("plot_dropdown_ui", inline = TRUE),
        shiny::tags$button(id = "plot_open_browser_btn",
                          onclick = "Shiny.setInputValue('plot_open_browser', Date.now(), {priority: 'event'});",
                          "Open in browser"),
        shiny::tags$button(id = "plot_save_btn",
                          onclick = "Shiny.setInputValue('plot_save_downloads', Date.now(), {priority: 'event'});",
                          "Save")
      ),
      shiny::div(id = "plot_iframe_wrap",
        shiny::uiOutput("plot_iframe_ui")
      )
    ),

    shiny::tags$script(shiny::HTML("
      // ── tsifl image handling for R add-in ──
      var pendingImages = [];

      // Attach button → open file picker
      document.getElementById('attach_btn').addEventListener('click', function() {
        document.getElementById('image_input').click();
      });

      // File picker change — accepts all file types
      document.getElementById('image_input').addEventListener('change', function(e) {
        var files = Array.from(e.target.files);
        files.forEach(function(f) {
          readFileAsBase64(f);
        });
        e.target.value = '';
      });

      // Drag & drop on input area
      var inputArea = document.getElementById('input_area');
      inputArea.addEventListener('dragover', function(e) {
        e.preventDefault(); e.stopPropagation();
        inputArea.style.outline = '2px dashed #0D5EAF';
        inputArea.style.outlineOffset = '-2px';
        inputArea.style.background = '#EBF3FB';
      });
      inputArea.addEventListener('dragleave', function(e) {
        e.preventDefault(); e.stopPropagation();
        inputArea.style.outline = '';
        inputArea.style.outlineOffset = '';
        inputArea.style.background = '';
      });
      inputArea.addEventListener('drop', function(e) {
        e.preventDefault(); e.stopPropagation();
        inputArea.style.outline = '';
        inputArea.style.outlineOffset = '';
        inputArea.style.background = '';
        var files = Array.from(e.dataTransfer.files || []);
        files.forEach(function(f) {
          readFileAsBase64(f);
        });
      });

      // Paste from clipboard — bind at WINDOW level with capture so it runs
      // before Shiny/textarea handlers can swallow it. Handles both items[] and files[].
      function handlePasteEvent(e) {
        var cd = e.clipboardData || window.clipboardData;
        if (!cd) return;
        var handled = false;
        // Try files[] first (most reliable for screenshots)
        var files = cd.files ? Array.from(cd.files) : [];
        files.forEach(function(f) {
          if (f) { readFileAsBase64(f); handled = true; }
        });
        // Fallback to items[]
        if (!handled && cd.items) {
          Array.from(cd.items).forEach(function(item) {
            if (item.kind === 'file' || (item.type && item.type.startsWith('image/'))) {
              var f = item.getAsFile();
              if (f) { readFileAsBase64(f); handled = true; }
            }
          });
        }
        if (handled) {
          e.preventDefault();
          e.stopPropagation();
          console.log('[tsifl] Image pasted, total pending:', pendingImages.length);
          var ta = document.getElementById('user_input');
          if (ta) { try { ta.focus(); } catch(_) {} }
        }
      }
      window.addEventListener('paste', handlePasteEvent, true);
      document.addEventListener('paste', handlePasteEvent, true);

      // Read any file → base64
      function readFileAsBase64(file) {
        var reader = new FileReader();
        reader.onload = function() {
          var base64 = reader.result;
          var isImage = file.type && file.type.startsWith('image/');
          var mediaType = file.type || (isImage ? 'image/png' : 'application/octet-stream');
          var data = base64.split(',')[1];
          pendingImages.push({ media_type: mediaType, data: data, file_name: file.name || '' });
          updatePreview();
        };
        reader.readAsDataURL(file);
      }

      // Backward compat alias
      function readImageFile(file) { readFileAsBase64(file); }

      // Render base64 → canvas (bypasses CSP restrictions)
      function renderToCanvas(base64Data, mediaType, maxW, maxH) {
        return new Promise(function(resolve) {
          try {
            var byteChars = atob(base64Data);
            var byteArray = new Uint8Array(byteChars.length);
            for (var i = 0; i < byteChars.length; i++) byteArray[i] = byteChars.charCodeAt(i);
            var blob = new Blob([byteArray], { type: mediaType || 'image/png' });
            createImageBitmap(blob).then(function(bitmap) {
              var w = bitmap.width, h = bitmap.height;
              var scale = Math.min(maxW / w, maxH / h, 1);
              w = Math.round(w * scale);
              h = Math.round(h * scale);
              var canvas = document.createElement('canvas');
              canvas.width = w;
              canvas.height = h;
              canvas.getContext('2d').drawImage(bitmap, 0, 0, w, h);
              bitmap.close();
              resolve(canvas);
            }).catch(function() { resolve(null); });
          } catch(e) { resolve(null); }
        });
      }

      // Update preview bar
      function updatePreview() {
        var bar = document.getElementById('image_preview_bar');
        var btn = document.getElementById('attach_btn');
        bar.innerHTML = '';
        if (pendingImages.length === 0) {
          bar.style.display = 'none';
          btn.textContent = '+';
          btn.title = 'Attach file';
          return;
        }
        btn.textContent = pendingImages.length;
        btn.title = pendingImages.length + ' file(s) attached';
        bar.style.display = 'flex';
        pendingImages.forEach(function(img, i) {
          var wrapper = document.createElement('div');
          wrapper.className = 'image-preview-item';
          var isImage = img.media_type && img.media_type.startsWith('image/');
          if (isImage) {
            renderToCanvas(img.data, img.media_type, 48, 48).then(function(canvas) {
              if (canvas) wrapper.insertBefore(canvas, wrapper.firstChild);
            });
          } else {
            var docIcon = document.createElement('div');
            var ext = img.file_name ? img.file_name.split('.').pop().toUpperCase() : 'FILE';
            docIcon.style.cssText = 'width:48px;height:48px;display:flex;align-items:center;justify-content:center;background:#F1F5F9;border-radius:4px;border:1px solid #E2E8F0;font-size:9px;font-weight:700;color:#0D5EAF;text-align:center;';
            docIcon.textContent = ext;
            wrapper.insertBefore(docIcon, wrapper.firstChild);
          }
          var rm = document.createElement('button');
          rm.className = 'remove-img';
          rm.textContent = 'x';
          rm.addEventListener('click', function() {
            pendingImages.splice(i, 1);
            updatePreview();
          });
          wrapper.appendChild(rm);
          bar.appendChild(wrapper);
        });
      }

      // Make pendingImages accessible globally
      window._tsiflImages = pendingImages;

      // Unbind Shiny's automatic textarea binding (prevents value restoration)
      var ta = document.getElementById('user_input');
      if (ta) {
        try { Shiny.unbindAll(ta.parentElement); } catch(e) {}
      }

      // Send function — captures message, clears textarea
      window._tsiflSend = function() {
        var ta = document.getElementById('user_input');
        if (!ta) return;
        var msg = ta.value.trim();
        if (!msg) return;

        // Button text indicates in-flight state; reset when response arrives.
        var btn = document.getElementById('send_btn');
        if (btn) btn.textContent = 'Sending...';

        // Safety net: if set_status('done') or set_status('error') never
        // reaches us (network drop, server-side exception swallowed, etc.),
        // force the button back to 'Send' after a generous timeout so the
        // UI doesn't look permanently stuck.
        if (window._tsiflSafetyTimer) clearTimeout(window._tsiflSafetyTimer);
        window._tsiflSafetyTimer = setTimeout(function() {
          var b = document.getElementById('send_btn');
          if (b && b.textContent === 'Sending...') b.textContent = 'Send';
        }, 120000);  // 2 minutes — longer than any normal R job

        // Pack images directly into send_message payload to avoid race
        var imgPayload = JSON.stringify(window._tsiflImages || []);
        if (window._tsiflImages) window._tsiflImages.length = 0;
        try { updatePreview(); } catch(e) {}
        console.log('[tsifl] Sending with images:', imgPayload.length, 'chars');

        // Hide old textarea and create new one
        ta.style.display = 'none';
        var parent = ta.parentNode;
        var newTa = document.createElement('textarea');
        newTa.id = 'user_input_new';
        newTa.placeholder = 'What can I help you with?';
        newTa.rows = 2;
        parent.insertBefore(newTa, ta);
        ta.remove();
        newTa.id = 'user_input';

        // Paste handler is document-level now — no need to rebind.

        // Re-attach Enter key handler
        newTa.addEventListener('keydown', function(e) {
          if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            document.getElementById('send_btn').click();
          }
        });

        // Send message to Shiny with images embedded in payload
        Shiny.setInputValue('send_message', JSON.stringify({msg: msg, images: imgPayload, ts: Date.now()}), {priority: 'event'});
      };

      // Enter key sends (Shift+Enter for newline)
      document.getElementById('user_input').addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          document.getElementById('send_btn').click();
        }
      });

      // Listen for image display in chat
      Shiny.addCustomMessageHandler('show_chat_image', function(msg) {
        // Wait for DOM update then render canvas into the target element
        setTimeout(function() {
          var target = document.getElementById(msg.target_id);
          if (!target) return;
          renderToCanvas(msg.data, msg.media_type, 260, 160).then(function(canvas) {
            if (canvas) {
              canvas.className = 'chat-canvas';
              target.appendChild(canvas);
            } else {
              target.innerHTML = '<span class=\"image-badge\">Image attached</span>';
            }
          });
        }, 100);
      });

      // ── Animated status rotation system ──
      var _statusInterval = null;
      var _statusMessages = {
        thinking: [
          'Reading your question...',
          'Analyzing the screenshots...',
          'Understanding what you need...',
          'Consulting the statistical gods...',
          'Thinking really hard right now...',
          'If I had a chin, I\\'d be stroking it...',
          'Formulating the perfect approach...',
          'My neurons are firing... all 175B of them...'
        ],
        generating: [
          'Writing R code...',
          'Building your analysis...',
          'Crafting the perfect model...',
          'Assembling statistical firepower...',
          'Making R do the heavy lifting...',
          'This is the fun part...',
          'Putting the pieces together...',
          'Almost there, promise...'
        ],
        running: [
          'Running code in your console...',
          'R is crunching numbers...',
          'Executing analysis...',
          'Plots are cooking...',
          'Waiting for R to finish...',
          'R is doing its thing...',
          'Numbers are being crunched as we speak...',
          'Your CPU is earning its keep right now...'
        ],
        interpreting: [
          'Reading the R output...',
          'Interpreting your results...',
          'Extracting the key values...',
          'Translating R-speak to human...',
          'Making sense of the numbers...',
          'Almost done, packaging your answers...',
          'Double-checking the values...',
          'No hallucinations on my watch...'
        ]
      };

      var _phaseLabels = {
        thinking: 'Thinking',
        generating: 'Generating code',
        running: 'Running in R',
        interpreting: 'Reading results'
      };

      function getOrCreateBubble() {
        var bubble = document.getElementById('tsifl-thinking-bubble');
        if (!bubble) {
          bubble = document.createElement('div');
          bubble.id = 'tsifl-thinking-bubble';
          bubble.innerHTML = '<div class=\"thinking-phase\"></div>' +
            '<div class=\"thinking-orb\"></div>' +
            '<div class=\"thinking-content\">' +
              '<span class=\"thinking-text\"></span>' +
            '</div>';
          var container = document.getElementById('thinking_container');
          if (container) {
            container.appendChild(bubble);
          }
          var chat = document.getElementById('chat_history');
          if (chat) chat.scrollTop = chat.scrollHeight;
        }
        return bubble;
      }

      function removeBubble() {
        var container = document.getElementById('thinking_container');
        if (container) container.innerHTML = '';
      }

      var _activePhase = null;

      function startStatusRotation(phase) {
        stopStatusRotation();
        _activePhase = phase;
        var msgs = _statusMessages[phase] || _statusMessages.thinking;
        var idx = 0;

        // Show and update status bar
        var bar = document.getElementById('status_bar');
        if (bar) bar.className = '';
        var dot = document.querySelector('#status_bar .status-dot');
        var statusText = document.getElementById('status_text');
        if (dot) dot.className = 'status-dot ' + (phase === 'running' || phase === 'interpreting' ? 'running' : 'thinking');

        function ensureBubble() {
          // Re-create bubble if it was destroyed by Shiny re-render
          var bubble = getOrCreateBubble();
          var phaseEl = bubble.querySelector('.thinking-phase');
          if (phaseEl) phaseEl.textContent = _phaseLabels[_activePhase] || _activePhase;
          return bubble;
        }

        function show() {
          if (!_activePhase) return;
          var bubble = ensureBubble();
          var textEl = bubble.querySelector('.thinking-text');
          var msg = msgs[idx % msgs.length];
          if (textEl) {
            textEl.style.opacity = '0';
            textEl.style.transform = 'translateY(4px)';
            setTimeout(function() {
              textEl.textContent = msg;
              textEl.style.opacity = '1';
              textEl.style.transform = 'translateY(0)';
            }, 200);
          }
          var chat = document.getElementById('chat_history');
          if (chat) chat.scrollTop = chat.scrollHeight;
          idx++;
        }

        // Delay initial show to let Shiny re-render complete first
        setTimeout(function() { if (_activePhase) show(); }, 200);
        _statusInterval = setInterval(show, 2500);
      }

      function stopStatusRotation() {
        if (_statusInterval) { clearInterval(_statusInterval); _statusInterval = null; }
      }

      function resetSendButton() {
        // Send button's onclick sets textContent to 'Sending...' but doesn't
        // reset it on completion — do it here so the user knows they can
        // send another message.
        var btn = document.getElementById('send_btn');
        if (btn) btn.textContent = 'Send';
        // Cancel the safety-net timer if present
        if (window._tsiflSafetyTimer) {
          clearTimeout(window._tsiflSafetyTimer);
          window._tsiflSafetyTimer = null;
        }
      }

      function setStatusDone(msg) {
        _activePhase = null;
        stopStatusRotation();
        removeBubble();
        var bar = document.getElementById('status_bar');
        if (bar) bar.className = 'idle';
        resetSendButton();
      }

      function setStatusError(msg) {
        _activePhase = null;
        stopStatusRotation();
        removeBubble();
        var bar = document.getElementById('status_bar');
        if (bar) bar.className = '';
        var dot = document.querySelector('#status_bar .status-dot');
        var text = document.getElementById('status_text');
        if (dot) dot.className = 'status-dot error';
        if (text) { text.style.opacity = '1'; text.textContent = msg || 'Disconnected'; }
        resetSendButton();
      }

      // Listen for status phase changes from Shiny server
      Shiny.addCustomMessageHandler('tsifl_status', function(msg) {
        if (msg.phase === 'done') { setStatusDone(msg.text || 'Done'); }
        else if (msg.phase === 'error') { setStatusError(msg.text || 'Disconnected'); }
        else { startStatusRotation(msg.phase); }
      });

      // Quick action — send message directly without touching textarea
      Shiny.addCustomMessageHandler('trigger_send', function(msg) {
        Shiny.setInputValue('send_message', msg.msg, {priority: 'event'});
      });

      // ── Tab switching: Chat | Plot ─────────────────────────────────
      window.tsiflShowTab = function(name) {
        var chatTab    = document.getElementById('chat_tab');
        var plotTab    = document.getElementById('plot_tab');
        var chatBtn    = document.getElementById('tab_chat_btn');
        var plotBtn    = document.getElementById('tab_plot_btn');
        if (!chatTab || !plotTab) return;
        if (name === 'plot') {
          chatTab.classList.add('hidden');
          plotTab.classList.add('active');
          chatBtn.classList.remove('active');
          plotBtn.classList.add('active');
          // Tell the server we want the plot list refreshed
          if (typeof Shiny !== 'undefined' && Shiny.setInputValue) {
            Shiny.setInputValue('plot_tab_opened', Date.now(), {priority: 'event'});
          }
          // Force the plot iframe to reload NOW that it's visible.
          // Without this, plotly inside the iframe measured itself at 0x0
          // while the tab was hidden and rendered at that size — you saw
          // only the title with an empty canvas below.
          setTimeout(function() {
            var f = document.getElementById('plot_iframe');
            if (f && f.src) {
              var url = f.src;
              f.src = 'about:blank';
              setTimeout(function() { f.src = url; }, 30);
            }
          }, 50);
        } else {
          chatTab.classList.remove('hidden');
          plotTab.classList.remove('active');
          chatBtn.classList.add('active');
          plotBtn.classList.remove('active');
        }
      };

      // When the selected plot changes (user picked a different timestamp
      // from the dropdown), also trigger a clean reload so plotly remeasures.
      window.addEventListener('message', function(e) {
        // no-op placeholder for potential cross-frame comms
      });

      // When a new plot is announced, flash the Plot tab badge so the
      // user notices something new is available without forcing them away
      // from the chat.
      Shiny.addCustomMessageHandler('plot_count_update', function(msg) {
        var badge = document.getElementById('plot_count');
        if (!badge) return;
        if (msg.count && msg.count > 0) {
          badge.textContent = msg.count;
          badge.style.display = 'inline-block';
        } else {
          badge.style.display = 'none';
        }
      });
    "))

  )

  # ── Server ─────────────────────────────────────────────────────────────────
  server <- function(input, output, session) {

    messages    <- shiny::reactiveVal(list())
    tasks_left  <- shiny::reactiveVal(NA)

    # Server-side heartbeat — keeps the session alive even when viewer tab
    # is inactive and JS keepalive stops running
    shiny::observe({
      shiny::invalidateLater(10000)  # every 10 seconds
      invisible(NULL)  # no-op, just prevent session GC
    })

    # Status helper — sends phase to JS for animated rotation
    set_status <- function(phase, text = NULL) {
      session$sendCustomMessage("tsifl_status", list(phase = phase, text = text))
    }

    output$chat_messages <- shiny::renderUI({
      msgs <- messages()
      if (length(msgs) == 0) {
        return(shiny::p("What can I help you with?",
          style = "color:#C7C7CC; font-size:14px; padding:20px 0; text-align:center;"))
      }
      ui_list <- lapply(msgs, function(m) {
        # Convert markdown-style formatting to HTML for assistant messages
        display_text <- m$text
        if (m$role == "assistant") {
          # Strip bold markers — keep text clean without heavy formatting
          display_text <- gsub("\\*\\*(.+?)\\*\\*", "\\1", display_text)
          # Inline code: `text` → <code>text</code>
          display_text <- gsub("`([^`]+)`", "<code>\\1</code>", display_text)
          # Line breaks
          display_text <- gsub("\n\n", "<br><br>", display_text)
          display_text <- gsub("\n", "<br>", display_text)
          # Bullet points
          display_text <- gsub("<br>- ", "<br>\u2022 ", display_text)
          display_text <- gsub("^- ", "\u2022 ", display_text)
        }
        children <- if (m$role == "assistant") {
          list(shiny::HTML(paste0('<span>', display_text, '</span>')))
        } else {
          list(shiny::span(m$text))
        }
        # Show compact badge for attached images/files — infer type from
        # filename extension so CSVs etc. don't get mislabeled as "images".
        if (!is.null(m$images) && length(m$images) > 0) {
          n_imgs <- length(m$images)
          # Inspect filenames to decide the label
          exts <- tolower(vapply(m$images, function(x) {
            nm <- if (is.list(x)) x$file_name else ""
            if (is.null(nm) || !nzchar(nm)) "" else tools::file_ext(nm)
          }, character(1)))
          is_image_ext <- exts %in% c("png", "jpg", "jpeg", "gif", "webp", "bmp", "svg", "")
          label <- if (all(is_image_ext)) {
            if (n_imgs == 1) "image" else "images"
          } else if (!any(is_image_ext)) {
            if (n_imgs == 1) "file" else "files"
          } else {
            if (n_imgs == 1) "attachment" else "attachments"
          }
          badge_text <- paste(n_imgs, label, "attached")
          children <- c(children, list(
            shiny::span(
              badge_text,
              style = "display:inline-block;margin-top:4px;font-size:10px;color:#64748B;background:#F1F5F9;padding:2px 8px;border-radius:8px;border:1px solid #E2E8F0;"
            )
          ))
        }
        # Render inline plot thumbnail + "Open interactive" button if the
        # assistant attached one. Clicking the thumbnail (or button) opens
        # the HTML version in the user's default browser. If there's no
        # HTML (e.g. it's a static ggplot), only show the thumbnail.
        if (!is.null(m$plot) && is.list(m$plot) && !is.null(m$plot$png_data_uri)) {
          plot_children <- list(
            shiny::tags$img(
              src = m$plot$png_data_uri,
              style = paste(
                "display:block; width:100%; max-width:560px; height:auto;",
                "border-radius:8px; border:1px solid #E2E8F0;",
                "margin-top:8px; margin-bottom:4px;",
                if (isTRUE(m$plot$interactive)) "cursor:pointer;" else ""
              ),
              onclick = if (isTRUE(m$plot$interactive) &&
                            !is.null(m$plot$html_path)) {
                # Use Shiny input to trigger server-side browseURL()
                sprintf("Shiny.setInputValue('open_plot_html', '%s', {priority: 'event'});",
                        gsub("'", "\\\\'", m$plot$html_path, fixed = TRUE))
              } else NULL
            )
          )
          if (isTRUE(m$plot$interactive) && !is.null(m$plot$html_path)) {
            plot_children <- c(plot_children, list(
              shiny::tags$button(
                "Open interactive",
                onclick = sprintf(
                  "Shiny.setInputValue('open_plot_html', '%s', {priority: 'event'});",
                  gsub("'", "\\\\'", m$plot$html_path, fixed = TRUE)
                ),
                style = paste(
                  "display:inline-block; margin-top:6px; font-size:10px;",
                  "font-weight:500; letter-spacing:0.3px; text-transform:uppercase;",
                  "color:#475569; background:#FFFFFF; border:1px solid #E2E8F0;",
                  "padding:4px 10px; border-radius:4px; cursor:pointer;",
                  "font-family:-apple-system,BlinkMacSystemFont,sans-serif;"
                )
              )
            ))
          }
          children <- c(children, plot_children)
        }
        shiny::div(class = paste0("msg-", m$role), children)
      })
      ui_list
    })

    output$tasks_label <- shiny::renderUI({
      t <- tasks_left()
      if (is.na(t)) return(shiny::span(id = "tasks_label", ""))
      shiny::span(id = "tasks_label", paste(t, "tasks left"))
    })

    # (status is now JS-driven via set_status())

    # ── Plot tab: serve files from /tmp/.tsifl_plots via /plots/ URL ───
    # Shiny doesn't serve arbitrary local files by default. We register
    # the plots dir as a static resource path so the iframe can load
    # plot HTML directly (file:// URLs don't work cross-origin in Shiny's
    # iframe sandbox, but http://127.0.0.1:7444/plots/x.html does).
    tryCatch(
      shiny::addResourcePath("plots", "/tmp/.tsifl_plots"),
      error = function(e) {}
    )

    # Reactive list of available plots (refreshed when user opens Plot
    # tab OR every 5s while it's open). Returns named vector of "label" =
    # "filename", sorted newest first.
    plot_files <- shiny::reactivePoll(
      intervalMillis = 5000,
      session = session,
      checkFunc = function() {
        files <- list.files("/tmp/.tsifl_plots", pattern = "\\.html$",
                            full.names = TRUE)
        if (length(files) == 0) return(0)
        # sum() doesn't work on POSIXct directly — convert to numeric
        # seconds first. Any change in file set or mtime shifts this
        # checksum, which is all reactivePoll needs.
        sum(as.numeric(file.info(files)$mtime))
      },
      valueFunc = function() {
        files <- list.files("/tmp/.tsifl_plots", pattern = "\\.html$",
                            full.names = TRUE)
        if (length(files) == 0) return(character(0))
        info <- file.info(files)
        ord  <- order(info$mtime, decreasing = TRUE)
        files <- files[ord]
        # Use basename as both label (humanized) and value (the URL
        # fragment we'll iframe).
        labels <- basename(files)
        # Format like "Plot — 5:42 PM" using the file timestamp
        labels <- vapply(seq_along(files), function(i) {
          mt <- info$mtime[ord[i]]
          paste0("Plot ", format(mt, "%I:%M %p"))
        }, character(1))
        names(files) <- labels
        files
      }
    )

    selected_plot <- shiny::reactiveVal(NULL)

    # Push plot count to JS so the badge on the Plot tab updates as new
    # plots come in (without forcing the user to switch away from chat).
    shiny::observe({
      pf <- plot_files()
      session$sendCustomMessage("plot_count_update", list(count = length(pf)))
      # Auto-select most recent if nothing chosen yet
      if (is.null(shiny::isolate(selected_plot())) && length(pf) > 0) {
        selected_plot(pf[1])
      }
    })

    # When the user opens the Plot tab, snap selection to most recent
    shiny::observeEvent(input$plot_tab_opened, {
      pf <- plot_files()
      if (length(pf) > 0) selected_plot(pf[1])
    })

    # Render the dropdown of plots
    output$plot_dropdown_ui <- shiny::renderUI({
      pf <- plot_files()
      if (length(pf) == 0) {
        return(shiny::span("No plots yet — generate one in chat",
                           style = "color:#94A3B8;font-size:12px;"))
      }
      sel <- shiny::isolate(selected_plot())
      if (is.null(sel) || !(sel %in% pf)) sel <- pf[1]
      shiny::tags$select(
        id = "plot_picker",
        onchange = "Shiny.setInputValue('plot_picker_change', this.value, {priority: 'event'});",
        lapply(seq_along(pf), function(i) {
          shiny::tags$option(value = pf[i], selected = if (pf[i] == sel) "selected" else NULL,
                             names(pf)[i])
        })
      )
    })

    shiny::observeEvent(input$plot_picker_change, {
      selected_plot(input$plot_picker_change)
    })

    # Render the iframe pointing at /plots/<basename>
    output$plot_iframe_ui <- shiny::renderUI({
      sel <- selected_plot()
      pf <- plot_files()
      if (is.null(sel) || length(pf) == 0 || !file.exists(sel)) {
        return(shiny::div(id = "plot_empty_state",
          shiny::HTML(paste(
            "<div style='font-size:13px;font-weight:500;color:#475569;",
            "letter-spacing:0.2px;'>No plots yet</div>",
            "<div style='font-size:11px;margin-top:6px;color:#94A3B8;",
            "letter-spacing:0.1px;'>Generate a visualization from the Chat tab.</div>"
          ))
        ))
      }
      url <- paste0("plots/", basename(sel))
      shiny::tags$iframe(id = "plot_iframe", src = url, frameborder = "0")
    })

    # "Open in browser" button
    shiny::observeEvent(input$plot_open_browser, {
      sel <- selected_plot()
      if (is.null(sel) || !file.exists(sel)) return()
      tryCatch(utils::browseURL(sel), error = function(e) {})
    })

    # "Save to Downloads" button — copies the HTML (and matching PNG if
    # present) to ~/Downloads/ with a friendly timestamped name.
    shiny::observeEvent(input$plot_save_downloads, {
      sel <- selected_plot()
      if (is.null(sel) || !file.exists(sel)) return()
      tryCatch({
        downloads_dir <- path.expand("~/Downloads")
        if (!dir.exists(downloads_dir)) dir.create(downloads_dir, recursive = TRUE)
        ts_tag <- format(Sys.time(), "%Y%m%d_%H%M%S")
        html_target <- file.path(downloads_dir, paste0("tsifl_plot_", ts_tag, ".html"))
        file.copy(sel, html_target, overwrite = TRUE)
        # If a sibling PNG exists, copy that too — bosses paste PNGs into decks
        png_src <- sub("\\.html$", ".png", sel)
        if (file.exists(png_src)) {
          png_target <- file.path(downloads_dir, paste0("tsifl_plot_", ts_tag, ".png"))
          file.copy(png_src, png_target, overwrite = TRUE)
        }
        # Tell the user it worked via a toast in the chat
        add_message("action",
          paste0("Saved plot to Downloads as tsifl_plot_", ts_tag,
                 if (file.exists(png_src)) ".html + .png" else ".html"))
      }, error = function(e) {
        add_message("action",
          paste0("Could not save plot: ", conditionMessage(e)))
      })
    })

    img_counter <- shiny::reactiveVal(0)

    add_message <- function(role, text, images = NULL, plot = NULL) {
      current <- shiny::isolate(messages())
      entry <- list(role = role, text = text)
      if (!is.null(images) && length(images) > 0) {
        img_ids <- lapply(seq_along(images), function(i) {
          n <- shiny::isolate(img_counter()) + i
          paste0("chat_img_", n)
        })
        img_counter(shiny::isolate(img_counter()) + length(images))
        entry$images <- images
        entry$img_ids <- img_ids
      }
      # plot attachment: list(png = "/path.png", html = "/path.html" | NULL,
      #                      interactive = TRUE/FALSE)
      # Rendered as an inline thumbnail with a "Open interactive" button
      # when html is provided.
      if (!is.null(plot) && is.list(plot) && !is.null(plot$png)) {
        entry$plot <- plot
      }
      messages(c(current, list(entry)))
    }

    # ── Environment snapshot (background job → main session IPC) ───────────
    # The tsifl Shiny server runs in a background job with its OWN .GlobalEnv.
    # To see the user's actual data we use two mechanisms:
    #   A) At startup: install a recurring `later` callback in the MAIN session
    #      that auto-captures the environment every 3 seconds.
    #   B) On each chat: also fire a one-shot capture as a safety net.
    # Both write to the same shared temp file.

    ENV_SNAPSHOT_FILE <- "/tmp/.tsifl_env_snapshot.rds"

    # The R code that captures the main session's environment.
    # Written as a standalone script file so `source()` is reliable.
    ENV_CAPTURE_SCRIPT <- "/tmp/.tsifl_capture_env.R"
    writeLines(c(
      'tryCatch({',
      '  nms <- setdiff(ls(.GlobalEnv), c(".tsifl_watcher", ".tsifl_capture"))',
      '  info <- lapply(nms, function(nm) {',
      '    obj <- tryCatch(get(nm, envir = .GlobalEnv), error = function(e) NULL)',
      '    if (is.null(obj)) return(list(name = nm, class = "unknown"))',
      '    r <- list(name = nm, class = paste(class(obj), collapse = ", "))',
      '    if (!is.null(dim(obj))) r$dim <- paste(dim(obj), collapse = "x")',
      '    if (!is.null(names(obj))) r$col_names <- paste(head(names(obj), 10), collapse = ", ")',
      '    tryCatch({',
      '      r$preview <- paste(utils::capture.output(utils::str(obj, max.level = 0, give.attr = FALSE))[1], collapse = "")',
      '    }, error = function(e) {})',
      '    r',
      '  })',
      '  pkgs <- gsub("^package:", "", grep("^package:", search(), value = TRUE))',
      '  saveRDS(list(env = info, pkgs = pkgs, ts = Sys.time()), "/tmp/.tsifl_env_snapshot.rds")',
      '}, error = function(e) {',
      '  saveRDS(list(env = list(), pkgs = character(0), ts = Sys.time(), err = conditionMessage(e)), "/tmp/.tsifl_env_snapshot.rds")',
      '})'
    ), ENV_CAPTURE_SCRIPT)

    # (A) Install a recurring watcher in the MAIN R session via sendToConsole.
    #     Uses later::later (always available — it's a shiny dependency).
    #     The watcher captures the env every 3 seconds automatically.
    watcher_cmd <- paste0(
      'local({ ',
      'if (!exists(".tsifl_watcher", envir = .GlobalEnv)) { ',
      '  assign(".tsifl_watcher", TRUE, envir = .GlobalEnv); ',
      '  .tsifl_capture <- function() { ',
      '    tryCatch(source("/tmp/.tsifl_capture_env.R", local = TRUE, echo = FALSE), error = function(e) {}); ',
      '    later::later(.tsifl_capture, delay = 3) ',
      '  }; ',
      '  assign(".tsifl_capture", .tsifl_capture, envir = .GlobalEnv); ',
      '  .tsifl_capture() ',
      '} })'
    )
    tryCatch(
      rstudioapi::sendToConsole(watcher_cmd, execute = TRUE, echo = FALSE, focus = FALSE),
      error = function(e) {
        # If sendToConsole fails at startup, try again after a short delay
        later::later(function() {
          tryCatch(
            rstudioapi::sendToConsole(watcher_cmd, execute = TRUE, echo = FALSE, focus = FALSE),
            error = function(e2) {}
          )
        }, delay = 2)
      }
    )


    get_r_context <- function() {
      # (B) Fire a one-shot capture as safety net (in case the watcher isn't running)
      tryCatch(
        rstudioapi::sendToConsole(
          'tryCatch(source("/tmp/.tsifl_capture_env.R", local = TRUE, echo = FALSE), error = function(e) {})',
          execute = TRUE, echo = FALSE, focus = FALSE
        ),
        error = function(e) {}
      )

      # Wait for a fresh snapshot (must be < 10 seconds old)
      snap <- NULL
      for (attempt in 1:6) {
        Sys.sleep(0.5)
        snap <- tryCatch(readRDS(ENV_SNAPSHOT_FILE), error = function(e) NULL)
        if (!is.null(snap) && !is.null(snap$ts)) {
          age <- as.numeric(difftime(Sys.time(), snap$ts, units = "secs"))
          if (age < 10) break  # Fresh enough
          snap <- NULL  # Too stale, wait for fresh one
        }
      }

      env_objs <- list()
      pkgs <- character(0)
      if (!is.null(snap)) {
        env_objs <- if (!is.null(snap$env)) snap$env else list()
        pkgs <- if (!is.null(snap$pkgs)) snap$pkgs else character(0)
      }

      # Fallback: at least get packages from this background session
      if (length(pkgs) == 0) {
        pkgs <- tryCatch({
          gsub("^package:", "", grep("^package:", search(), value = TRUE))
        }, error = function(e) character(0))
      }

      # Get active editor tab from RStudio IDE
      open_tabs <- tryCatch({
        ctx <- rstudioapi::getActiveDocumentContext()
        doc_info <- list()

        # Try multiple methods to get the file path
        doc_path <- ctx$path
        if (!nzchar(doc_path)) {
          # Fallback: try documentPath for the active document
          doc_path <- tryCatch(rstudioapi::documentPath(ctx$id), error = function(e) "")
          if (is.null(doc_path)) doc_path <- ""
        }

        if (nzchar(doc_path)) {
          doc_info$active_file <- basename(doc_path)
          doc_info$active_file_path <- doc_path
        }
        if (length(ctx$contents) > 0) {
          doc_info$active_preview <- paste(ctx$contents, collapse = "\n")
          # If we still don't have a filename, detect Rmd from content
          if (is.null(doc_info$active_file)) {
            content_str <- tolower(doc_info$active_preview)
            # If it has YAML header + exercise chunks, it's an Rmd
            if (grepl("^---", ctx$contents[1]) && (grepl("exercise", content_str) || grepl("```\\{r", content_str))) {
              doc_info$active_file <- "document.Rmd"
            }
          }
        }
        doc_info
      }, error = function(e) list())

      list(
        app         = "rstudio",
        r_version   = R.version$version.string,
        working_dir = getwd(),
        loaded_pkgs = paste(pkgs, collapse = ", "),
        env_objects = env_objs,
        open_editor = open_tabs
      )
    }

    # Open a plot HTML file in the user's default browser when they click
    # the thumbnail or "Open interactive version" button in a chat message.
    shiny::observeEvent(input$open_plot_html, {
      path <- input$open_plot_html
      if (is.null(path) || !nzchar(path)) return()
      # Only allow files inside our plots dir — prevents any funny business
      # from a tampered input value.
      if (!startsWith(path, "/tmp/.tsifl_plots/")) return()
      if (!file.exists(path)) return()
      tryCatch(utils::browseURL(path), error = function(e) {})
    })

    # Quick action handler (Improvements 60, 62, 63, 68)
    shiny::observeEvent(input$quick_action, {
      prompts <- list(
        t_test = "Run a t-test on my data. Ask me which variable and hypothesis if unclear.",
        linear_reg = "Run a linear regression on my data. Ask which variables to use if unclear.",
        anova = "Run an ANOVA test on my data. Ask which variables to use if unclear.",
        ggplot_scatter = "Create a ggplot scatter plot from my data. Ask which variables to use if unclear.",
        summary_stats = "Generate comprehensive summary statistics for all numeric variables in my data.",
        correlation = "Create a correlation matrix for all numeric variables in my data with a visualization.",
        which_test = "Based on my data, which statistical test should I use? Help me choose the right test for my research question.",
        profile_data = "Profile my data: show nrow, ncol, column types, percent missing per column, unique values per column, min/max/mean for numerics, and top 5 values for categoricals.",
        compare_models = "Compare all model objects in my R environment. Show R-squared, Adjusted R-squared, AIC, BIC, and p-values in a comparison table."
      )
      prompt <- prompts[[input$quick_action]]
      if (!is.null(prompt)) {
        # Directly trigger send_message (no need to populate textarea)
        session$sendCustomMessage("trigger_send", list(msg = prompt))
      }
    }, ignoreInit = TRUE)

    # Poll the backend for cross-app R jobs sent from Excel (every 5s)
    seen_r_jobs <- new.env(parent = emptyenv())
    shiny::observe({
      shiny::invalidateLater(5000)
      tryCatch({
        resp <- httr2::request(BACKEND_URL) |>
          httr2::req_url_path_append("transfer", "pending", "rstudio") |>
          httr2::req_options(ssl_verifypeer = 0) |>
          httr2::req_timeout(5) |>
          httr2::req_perform()
        pending <- httr2::resp_body_json(resp, simplifyVector = FALSE)$pending
        if (is.list(pending) && length(pending) > 0) {
          for (p in pending) {
            if (identical(p$data_type, "r_job") && !is.null(p$transfer_id) &&
                is.null(seen_r_jobs[[p$transfer_id]])) {
              seen_r_jobs[[p$transfer_id]] <- TRUE
              # Fetch the full transfer to get the prompt
              tryCatch({
                full_resp <- httr2::request(BACKEND_URL) |>
                  httr2::req_url_path_append("transfer", p$transfer_id) |>
                  httr2::req_options(ssl_verifypeer = 0) |>
                  httr2::req_timeout(5) |>
                  httr2::req_perform()
                full <- httr2::resp_body_json(full_resp, simplifyVector = FALSE)
                job_prompt <- full$data
                if (is.character(job_prompt) && nchar(job_prompt) > 0) {
                  session$sendCustomMessage("trigger_send", list(msg = job_prompt))
                }
              }, error = function(e) {})
            }
          }
        }
      }, error = function(e) {})
    })


    shiny::observeEvent(input$send_message, {
      raw <- input$send_message
      if (is.null(raw) || nchar(raw) == 0) return()

      # New format: JSON {msg, images, ts}. Fallback: plain string.
      msg <- ""
      images <- list()
      parsed <- tryCatch(jsonlite::fromJSON(raw, simplifyVector = FALSE), error = function(e) NULL)
      if (is.list(parsed) && !is.null(parsed$msg)) {
        msg <- trimws(parsed$msg)
        if (!is.null(parsed$images) && nchar(parsed$images) > 2) {
          images <- jsonlite::fromJSON(parsed$images, simplifyVector = FALSE)
        }
      } else {
        msg <- trimws(raw)
      }
      if (nchar(msg) == 0) return()
      cat("[tsifl] send_message received:", length(images), "images\n")

      # Input already cleared by JS — just show user message + start thinking
      add_message("user", msg, images = if (length(images) > 0) images else NULL)
      set_status("thinking")

      # Capture context NOW (inside reactive context), before deferring
      r_context <- get_r_context()

      # Build request body now (inside reactive context)
      body <- list(
        user_id = USER_ID,
        message = msg,
        context = r_context
      )
      if (length(images) > 0) {
        body$images <- images
      }

      # Capture locals so the deferred callback can access them
      local_msg <- msg
      local_wants_excel <- grepl("excel|export", msg, ignore.case = TRUE)

      # Wait for Shiny to flush UI updates, THEN start the blocking HTTP call
      session$onFlushed(function() {
        later::later(function() {

        tryCatch({
          resp <- httr2::request(BACKEND_URL) |>
            httr2::req_url_path_append("chat", "") |>
            httr2::req_headers("Content-Type" = "application/json") |>
            httr2::req_body_json(body) |>
            httr2::req_options(ssl_verifypeer = 0) |>
            # 180s — allows room for the server-side retry loop (up to 2
            # extra API calls at ~20-30s each) without dropping the client.
            httr2::req_timeout(180) |>
            httr2::req_perform()

          set_status("generating")
          data <- httr2::resp_body_json(resp, simplifyVector = FALSE)
          has_reply   <- !is.null(data$reply) && nchar(trimws(data$reply)) > 0
          # Peek at actions presence without full parsing (detailed extraction below)
          has_actions <- (!is.null(data$actions) && length(data$actions) > 0) ||
                         (!is.null(data$action) && is.list(data$action) &&
                          !is.null(data$action$type) &&
                          !identical(data$action$type, "none"))
          if (has_reply) {
            add_message("assistant", data$reply)
          } else if (!has_actions) {
            # Backend guard already injects a fallback, but belt-and-suspenders:
            # if somehow we get a genuinely empty response, the user should
            # see SOMETHING so they don't think tsifl is broken.
            add_message("assistant",
              paste0("I didn't get a response to that — could you rephrase? ",
                     "(If this keeps happening, it might be a backend hiccup; ",
                     "try again in a moment.)"))
          }

          if (!is.null(data$tasks_remaining) && data$tasks_remaining >= 0)
            tasks_left(data$tasks_remaining)

          all_actions <- list()
          if (!is.null(data$actions) && is.list(data$actions) && length(data$actions) > 0) {
            for (a in data$actions) {
              if (is.list(a) && !is.null(a$type)) {
                all_actions <- c(all_actions, list(a))
              }
            }
          }
          # Only use singular action if actions list was empty (avoid duplicates)
          if (length(all_actions) == 0 &&
              !is.null(data$action) && is.list(data$action) && !is.null(data$action$type) &&
              !identical(data$action$type, "none")) {
            all_actions <- c(all_actions, list(data$action))
          }

          # Merge multiple run_r_code actions into ONE
          r_code_parts <- c()
          non_r_actions <- list()
          for (a in all_actions) {
            if (identical(a$type, "run_r_code") && !is.null(a$payload$code)) {
              r_code_parts <- c(r_code_parts, a$payload$code)
            } else {
              non_r_actions <- c(non_r_actions, list(a))
            }
          }
          if (length(r_code_parts) > 0) {
            merged_code <- paste(r_code_parts, collapse = "\n\n")
            merged_action <- list(type = "run_r_code", payload = list(code = merged_code))
            all_actions <- c(list(merged_action), non_r_actions)
          } else {
            all_actions <- non_r_actions
          }

          if (length(all_actions) > 10) {
            add_message("action", paste0(length(all_actions), " actions received, executing first 10"))
            all_actions <- all_actions[1:10]
          }

          if (length(all_actions) > 0) set_status("running")
          r_code_executed <- FALSE
          for (action in all_actions) {
            tryCatch({
              execute_r_action(action, add_message, r_context, user_wants_excel = local_wants_excel)
              if (identical(action$type, "run_r_code")) r_code_executed <- TRUE
            }, error = function(e) {
              add_message("action", paste0("Error: ", e$message))
            })
          }

          # Phase 2: interpret R output
          # Threshold lowered from 50 → 5 chars: even tiny outputs like a single
          # coefficient ("0.7611") deserve a chat-side summary so the user isn't
          # forced to read the R console to find the answer.
          if (r_code_executed) {
            tryCatch({
              Sys.sleep(5)
              r_output <- ""
              for (.retry in 1:3) {
                if (file.exists("/tmp/.tsifl_last_output.txt")) {
                  r_output <- tryCatch(
                    paste(readLines("/tmp/.tsifl_last_output.txt", warn = FALSE), collapse = "\n"),
                    error = function(e) ""
                  )
                  if (nchar(trimws(r_output)) > 5) break
                }
                Sys.sleep(2)
              }

              if (nchar(trimws(r_output)) > 5) {
                r_codes <- sapply(all_actions, function(a) {
                  if (identical(a$type, "run_r_code")) a$payload$code else NULL
                })
                r_codes <- paste(Filter(Negate(is.null), r_codes), collapse = "\n")

                phase1_reply <- if (!is.null(data$reply)) substr(data$reply, 1, 3000) else ""
                followup_msg <- paste0(
                  "[R OUTPUT INTERPRETATION]\n",
                  "The user asked: \"", local_msg, "\"\n\n",
                  "Your earlier analysis identified these questions/tasks:\n", phase1_reply, "\n\n",
                  "R code executed:\n```r\n", substr(r_codes, 1, 2000), "\n```\n\n",
                  "R output:\n```\n", substr(r_output, 1, 8000), "\n```\n\n",
                  "Answer EACH question/part using the actual R output values.\n",
                  "FORMAT RULES (strict):\n",
                  "- Put each answer on its OWN LINE with a blank line between parts\n",
                  "- Start each with **a.**, **b.**, etc. on a new line\n",
                  "- Bold key values\n",
                  "- Keep each answer to 1-2 sentences max\n",
                  "- No introductions, no dataset descriptions\n",
                  "- If a question asks for a rounded number, round it\n\n",
                  "Example format:\n",
                  "**a.** Categorical variables: sex, smoker\n\n",
                  "**b.** F-statistic p-value is **0.00** (< 0.05), model is significant\n\n",
                  "**c.** R-squared = **0.7509**, meaning 75.09% of variation is explained\n\n",
                  "End your reply with a single short line asking if the user wants a plot or any additional analysis. Do NOT say 'Done' or 'let me know if you'd like changes'."
                )

                followup_body <- list(
                  user_id = USER_ID,
                  message = followup_msg,
                  context = list(app = "rstudio")
                )

                set_status("interpreting")
                followup_resp <- httr2::request(BACKEND_URL) |>
                  httr2::req_url_path_append("chat", "") |>
                  httr2::req_headers("Content-Type" = "application/json") |>
                  httr2::req_body_json(followup_body) |>
                  httr2::req_options(ssl_verifypeer = 0) |>
                  httr2::req_timeout(90) |>
                  httr2::req_perform()

                followup_data <- httr2::resp_body_json(followup_resp, simplifyVector = FALSE)
                if (!is.null(followup_data$reply) && nchar(followup_data$reply) > 5) {
                  # Attach plot preview if the listener saved one
                  plot_attachment <- NULL
                  tryCatch({
                    if (file.exists("/tmp/.tsifl_last_plot.json")) {
                      raw <- paste(readLines("/tmp/.tsifl_last_plot.json",
                                             warn = FALSE),
                                   collapse = "")
                      meta <- jsonlite::fromJSON(raw, simplifyVector = TRUE)
                      if (!is.null(meta$png) && file.exists(meta$png)) {
                        # Convert PNG to data URI so it embeds directly
                        png_bytes <- readBin(meta$png, "raw",
                                             n = file.info(meta$png)$size)
                        png_b64 <- base64enc::base64encode(png_bytes)
                        plot_attachment <- list(
                          png_data_uri = paste0("data:image/png;base64,", png_b64),
                          html_path = if (isTRUE(meta$interactive) &&
                                          !is.null(meta$html) &&
                                          file.exists(meta$html)) {
                            meta$html
                          } else NULL,
                          interactive = isTRUE(meta$interactive)
                        )
                        # Consume the metadata file so we don't re-attach
                        # the same plot to the next turn by accident.
                        try(unlink("/tmp/.tsifl_last_plot.json"), silent = TRUE)
                      }
                    }
                  }, error = function(e) {})
                  add_message("assistant", followup_data$reply,
                              plot = plot_attachment)
                }
              }
            }, error = function(e) {
              # Phase 2 is best-effort
            })
          }

          set_status("done")

        }, error = function(e) {
          add_message("assistant",
            paste0("Could not reach backend.\n", e$message))
          set_status("error", "Disconnected")
        })
        }, delay = 0.05)
      }, once = TRUE)
    })

    # ── Action executor ──────────────────────────────────────────────────────
    execute_r_action <- function(action, add_message, r_context = list(), user_wants_excel = FALSE) {
      type    <- action$type
      payload <- action$payload

      if (type == "run_r_code") {
        code <- payload$code

        # target controls where the code VISIBLY lands in the editor.
        # Values:
        #   "console" — run only, don't touch the editor (pure REPL)
        #   "new"     — open a fresh .R script tab with the code (default)
        #   "active"  — append the code to the currently-open editor tab
        target <- if (is.null(payload$target)) "new" else payload$target

        if (target != "console") {
          code_file <- "/tmp/.tsifl_insert_code.R"
          tryCatch(writeLines(code, code_file), error = function(e) {})

          if (target == "active") {
            # Append to the currently active document
            insert_script <- paste0(
              'local({\n',
              '  code <- paste(readLines("', code_file, '"), collapse = "\\n")\n',
              '  tryCatch({\n',
              '    ctx <- rstudioapi::getActiveDocumentContext()\n',
              '    last_line <- length(ctx$contents)\n',
              '    last_col  <- nchar(ctx$contents[last_line])\n',
              '    rstudioapi::insertText(\n',
              '      location = c(last_line, last_col + 1),\n',
              '      text = paste0("\\n\\n# tsifl — Generated Code\\n", code, "\\n"),\n',
              '      id = ctx$id\n',
              '    )\n',
              '  }, error = function(e) {\n',
              '    # Fallback: new tab if no active document\n',
              '    rstudioapi::documentNew(\n',
              '      text = paste0("# tsifl — Generated Code\\n\\n", code, "\\n"),\n',
              '      type = "r"\n',
              '    )\n',
              '  })\n',
              '})\n'
            )
          } else {
            # Default "new": open a fresh script tab
            insert_script <- paste0(
              'local({\n',
              '  code <- paste(readLines("', code_file, '"), collapse = "\\n")\n',
              '  rstudioapi::documentNew(\n',
              '    text = paste0("# tsifl — Generated Code\\n\\n", code, "\\n"),\n',
              '    type = "r"\n',
              '  )\n',
              '})\n'
            )
          }

          script_file <- "/tmp/.tsifl_insert_script.R"
          tryCatch(writeLines(insert_script, script_file), error = function(e) {})

          tryCatch({
            rstudioapi::sendToConsole(
              paste0('invisible(source("', script_file, '", local = TRUE))'),
              execute = TRUE, echo = FALSE, focus = FALSE
            )
          }, error = function(e) {})
          Sys.sleep(0.8)
        }

        # 2. FILE-BASED BRIDGE to the MAIN R session.
        #    We used to rely on rstudioapi::sendToConsole, but that IPC path is
        #    unreliable from background jobs — many users see "RStudio did not
        #    respond to rstudioapi IPC request" even when the architecture is
        #    correct. Instead, the main R session has a listener (installed by
        #    tsifulator_addin()) that polls /tmp/.tsifl_pending_code.R every
        #    0.5s via later::later(). We write code there and wait for a done
        #    marker at /tmp/.tsifl_done.marker.
        #
        #    The listener handles sink(), output capture, and plot snapshot —
        #    so no preamble commands needed.

        pending_file <- "/tmp/.tsifl_pending_code.R"
        done_file    <- "/tmp/.tsifl_done.marker"

        # Detect if code will produce a plot (used for wait timing below)
        plot_keywords <- c("plot(", "ggplot(", "boxplot(", "hist(", "barplot(",
                          "geom_", "abline(", "curve(", "pie(", "heatmap(",
                          "pairs(", "qqnorm(", "acf(", "pacf(", "stripchart(",
                          "par(mfrow")
        code_has_plot <- any(sapply(plot_keywords, function(kw) grepl(kw, code, fixed = TRUE)))

        code_to_run <- code

        # Clean previous done marker so we don't race with stale signals
        try(unlink(done_file), silent = TRUE)

        send_err <- ""
        sent <- tryCatch({
          writeLines(code_to_run, pending_file)
          TRUE
        }, error = function(e) {
          send_err <<- paste("Could not write bridge file:", conditionMessage(e))
          FALSE
        })

        if (sent) {
          add_message("action", "Running in console")

          # Wait for the listener to pick it up and signal done. Poll up to
          # ~60s (long code like models on large data can take a while). The
          # listener drops a timestamp into done_file when eval() returns.
          max_wait <- if (code_has_plot) 90 else 60
          done_ok <- FALSE
          for (.waited in seq_len(max_wait * 4)) {
            Sys.sleep(0.25)
            if (file.exists(done_file)) { done_ok <- TRUE; break }
          }

          if (!done_ok) {
            # The listener never fired — likely because it wasn't installed
            # (old tsifl version still running, or user never called
            # tsifulator_addin() in the main session). Fall back to the
            # legacy rstudioapi::sendToConsole path with retry.
            send_err <- paste0(
              "Listener did not respond within ", max_wait, "s — ",
              "the main R session may not have the tsifl bridge installed. ",
              "Fix: in the main R console run `tsifulator:::.install_tsifl_listener()` ",
              "or restart via Session → Restart R, then relaunch tsifl."
            )
            # Try the legacy path as a courtesy
            legacy_ok <- tryCatch({
              rstudioapi::sendToConsole(
                code_to_run, execute = TRUE, echo = TRUE, focus = FALSE
              )
              TRUE
            }, error = function(e) FALSE)
            if (legacy_ok) {
              Sys.sleep(if (code_has_plot) 6 else 3)
              done_ok <- TRUE
              send_err <- ""
            }
          }

          # Small settle delay after done marker for sink to flush etc.
          Sys.sleep(0.3)

          # ── Cross-app memory: snapshot data frames ──────────────────────
          tryCatch({
            # Read captured output
            output_text <- ""
            if (file.exists("/tmp/.tsifl_last_output.txt")) {
              output_text <- tryCatch(
                paste(readLines("/tmp/.tsifl_last_output.txt", warn = FALSE), collapse = "\n"),
                error = function(e) ""
              )
            }
            if (nchar(output_text) > 0) {
              first_line <- strsplit(output_text, "\n")[[1]][1]
              transfer_body <- list(
                from_app  = "rstudio",
                to_app    = "any",
                data_type = "r_output",
                data      = substr(output_text, 1, 5000),
                metadata  = list(code = substr(code, 1, 500), summary = substr(first_line, 1, 200))
              )
              tryCatch({
                httr2::request(BACKEND_URL) |>
                  httr2::req_url_path_append("transfer", "store") |>
                  httr2::req_headers("Content-Type" = "application/json") |>
                  httr2::req_body_json(transfer_body) |>
                  httr2::req_options(ssl_verifypeer = 0) |>
                  httr2::req_perform()
              }, error = function(e) {})
            }

            # POST data_snapshot for each data frame. We read from the env
            # snapshot that the main-session watcher maintains (every 3s) —
            # NOT from a non-existent /tmp/.tsifl_df_info.rds which never
            # existed and caused a noisy gzfile warning on every chat.
            .pick <- function(x, default = "") if (is.null(x) || length(x) == 0) default else x
            env_snap <- suppressWarnings(tryCatch(
              readRDS("/tmp/.tsifl_env_snapshot.rds"),
              error = function(e) list()
            ))
            env_objs <- if (is.list(env_snap) && !is.null(env_snap$env))
              env_snap$env else list()
            for (obj in env_objs) {
              # Only send data frames / tibbles, not random scalars
              cls <- .pick(obj$class)
              if (!grepl("data.frame|tbl", cls, ignore.case = TRUE)) next
              obj_name <- .pick(obj$name, "unknown")
              obj_dim  <- .pick(obj$dim, "?")
              obj_cols <- .pick(obj$col_names)
              transfer_body <- list(
                from_app  = "rstudio",
                to_app    = "any",
                data_type = "data_snapshot",
                data      = paste0("Data frame '", obj_name, "': ", obj_dim,
                                   " (Columns: ", obj_cols, ")"),
                metadata  = list(
                  name      = obj_name,
                  dim       = obj_dim,
                  col_names = obj_cols,
                  class     = cls
                )
              )
              tryCatch({
                httr2::request(BACKEND_URL) |>
                  httr2::req_url_path_append("transfer", "store") |>
                  httr2::req_headers("Content-Type" = "application/json") |>
                  httr2::req_body_json(transfer_body) |>
                  httr2::req_options(ssl_verifypeer = 0) |>
                  httr2::req_perform()
              }, error = function(e) {})
            }
          }, error = function(e) {
            # Cross-app capture is best-effort, don't break the main flow
          })

          if (code_has_plot && isTRUE(user_wants_excel)) {
            tryCatch({
              if (file.exists(plot_path) && file.info(plot_path)$size > 100) {
                img_b64 <- base64enc::base64encode(plot_path)
                unlink(plot_path)

                transfer_body <- list(
                  from_app  = "rstudio",
                  to_app    = "excel",
                  data_type = "image",
                  data      = img_b64,
                  metadata  = list(title = "R Plot")
                )

                httr2::request(BACKEND_URL) |>
                  httr2::req_url_path_append("transfer", "store") |>
                  httr2::req_headers("Content-Type" = "application/json") |>
                  httr2::req_body_json(transfer_body) |>
                  httr2::req_options(ssl_verifypeer = 0) |>
                  httr2::req_perform()

                add_message("action", "Plot exported to Excel")
              }
            }, error = function(e) {
              # Silently skip — plot capture is best-effort
            })
          }
        } else {
          add_message("action", paste0("Could not send to console: ", send_err))
        }

      } else if (type == "install_package") {
        pkg <- payload$package
        add_message("action", paste0("Installing ", pkg))
        rstudioapi::sendToConsole(
          paste0('install.packages("', pkg, '")'),
          execute = TRUE, echo = TRUE, focus = FALSE
        )
        add_message("action", paste0("Installing ", pkg, " in console"))

      } else if (type == "export_plot") {
        # Save current plot via main console, then upload
        tryCatch({
          plot_path <- "/tmp/.tsifl_last_plot.png"
          # First check if auto-capture already saved the plot
          if (!file.exists(plot_path) || file.info(plot_path)$size < 100) {
            # No auto-captured plot — try dev.copy from console
            save_cmd <- sprintf(
              'tryCatch({ grDevices::dev.copy(grDevices::png, "%s", width=800, height=600, res=150); grDevices::dev.off() }, error=function(e){})',
              plot_path
            )
            rstudioapi::sendToConsole(save_cmd, execute = TRUE, echo = FALSE, focus = FALSE)
            Sys.sleep(2)
          }

          if (file.exists(plot_path) && file.info(plot_path)$size > 0) {
            img_b64 <- base64enc::base64encode(plot_path)
            unlink(plot_path)

            target_app  <- if (!is.null(payload$to_app)) payload$to_app else "excel"
            target_cell <- if (!is.null(payload$cell)) payload$cell else "A1"
            target_sheet <- if (!is.null(payload$sheet)) payload$sheet else ""

            transfer_body <- list(
              from_app  = "rstudio",
              to_app    = target_app,
              data_type = "image",
              data      = img_b64,
              metadata  = list(cell = target_cell, sheet = target_sheet)
            )

            resp <- httr2::request(BACKEND_URL) |>
              httr2::req_url_path_append("transfer", "store") |>
              httr2::req_headers("Content-Type" = "application/json") |>
              httr2::req_body_json(transfer_body) |>
              httr2::req_options(ssl_verifypeer = 0) |>
              httr2::req_perform()

            result <- httr2::resp_body_json(resp)
            tid <- result$transfer_id
            add_message("action", "Plot exported to Excel")
          } else {
            add_message("action", "No plot found in Plots pane")
          }
        }, error = function(e) {
          add_message("action", paste0("Could not export plot: ", e$message))
        })

      } else if (type == "fill_rmd_chunks") {
        # Fill empty code chunks in the active Rmd file by exercise number.
        # Works DIRECTLY on the file on disk (no rstudioapi needed from background job).
        chunks  <- payload$chunks   # named list of exercise -> R code
        answers <- payload$answers  # named list of exercise -> text answer (optional)
        if (is.null(chunks)) chunks <- list()
        if (is.null(answers)) answers <- list()

        tryCatch({
          # Get the file path — background job can't use rstudioapi directly,
          # so use sendToConsole to ask the MAIN R session for the path
          file_path <- r_context$open_editor$active_file_path

          if (is.null(file_path) || !nzchar(file_path)) {
            # Ask the main R session for the active document path
            # Try multiple rstudioapi methods for reliability
            path_file <- "/tmp/.tsifl_rmd_path.txt"
            tryCatch({
              unlink(path_file)
              rstudioapi::sendToConsole(
                paste0('local({ ',
                  'p <- tryCatch(rstudioapi::getActiveDocumentContext()$path, error=function(e)""); ',
                  'if(!nzchar(p)) p <- tryCatch(rstudioapi::documentPath(), error=function(e)""); ',
                  'if(is.null(p)) p <- ""; ',
                  'writeLines(p, "', path_file, '") ',
                '})'),
                execute = TRUE, echo = FALSE, focus = FALSE
              )
              Sys.sleep(1)
              if (file.exists(path_file)) {
                file_path <- trimws(readLines(path_file, warn = FALSE)[1])
              }
            }, error = function(e) {})
          }

          if (is.null(file_path) || !nzchar(file_path) || !file.exists(file_path)) {
            # Still no path — search common directories and match by content
            home <- Sys.getenv("HOME", path.expand("~"))
            search_dirs <- c(
              r_context$working_dir,
              file.path(home, "Documents"),
              file.path(home, "Downloads"),
              file.path(home, "Desktop"),
              home,
              # Common course/project folders
              file.path(home, "Documents", "R"),
              file.path(home, "Documents", "Stats"),
              file.path(home, "Documents", "School"),
              file.path(home, "Documents", "Courses")
            )
            search_dirs <- unique(search_dirs[!is.na(search_dirs) & nzchar(search_dirs)])

            # Extract YAML title from open editor for matching
            preview_lines <- r_context$open_editor$active_preview
            preview_title <- ""
            if (!is.null(preview_lines)) {
              title_m <- regmatches(preview_lines, regexpr("title:\\s*['\"]?(.+?)(?:['\"]?\\s*$)", preview_lines, perl = TRUE))
              if (length(title_m) > 0) preview_title <- tolower(trimws(title_m))
            }

            # Phase 1: Search listed dirs for title match
            for (d in search_dirs) {
              if (!dir.exists(d)) next
              rmd_files <- list.files(d, pattern = "\\.Rmd$", full.names = TRUE, ignore.case = TRUE)
              if (length(rmd_files) == 0) next
              if (nzchar(preview_title)) {
                for (f in rmd_files) {
                  first_lines <- tryCatch(paste(readLines(f, n = 5, warn = FALSE), collapse = " "), error = function(e) "")
                  if (grepl(preview_title, tolower(first_lines), fixed = TRUE)) {
                    file_path <- f; break
                  }
                }
                if (!is.null(file_path) && nzchar(file_path) && file.exists(file_path)) break
              }
            }

            # Phase 2: If no title match, do a RECURSIVE search from home for title match
            if (is.null(file_path) || !nzchar(file_path) || !file.exists(file_path)) {
              if (nzchar(preview_title)) {
                all_rmd <- list.files(home, pattern = "\\.Rmd$", full.names = TRUE,
                                      ignore.case = TRUE, recursive = TRUE)
                # Skip hidden dirs and library paths
                all_rmd <- all_rmd[!grepl("/\\.|/Library/|/\\.Trash/", all_rmd)]
                # Limit to first 50 to avoid slowness
                all_rmd <- utils::head(all_rmd, 50)
                for (f in all_rmd) {
                  first_lines <- tryCatch(paste(readLines(f, n = 5, warn = FALSE), collapse = " "), error = function(e) "")
                  if (grepl(preview_title, tolower(first_lines), fixed = TRUE)) {
                    file_path <- f; break
                  }
                }
              }
            }

            # Phase 3: Last resort — first .Rmd in any search dir
            if (is.null(file_path) || !nzchar(file_path) || !file.exists(file_path)) {
              for (d in search_dirs) {
                if (!dir.exists(d)) next
                rmd_files <- list.files(d, pattern = "\\.Rmd$", full.names = TRUE, ignore.case = TRUE)
                if (length(rmd_files) > 0) { file_path <- rmd_files[1]; break }
              }
            }
          }

          if (is.null(file_path) || !nzchar(file_path) || !file.exists(file_path)) {
            add_message("action", "Could not find Rmd file to fill")
          } else {
            # Read the file directly from disk
            lines <- readLines(file_path, warn = FALSE)
            filled <- 0

            # Process exercises in REVERSE order so line insertions don't shift indices
            ex_nums <- sort(as.integer(gsub("[^0-9]", "", names(chunks))), decreasing = TRUE)

            for (ex_num in ex_nums) {
              ex_key <- paste0("Exercise ", ex_num)
              code <- chunks[[ex_key]]
              if (is.null(code)) next

              # Find #### Exercise N header
              header_pat <- paste0("^####\\s+Exercise\\s+", ex_num, "\\b")
              header_idx <- grep(header_pat, lines)
              if (length(header_idx) == 0) next
              header_idx <- header_idx[1]

              # Find next ```{r opening (within 10 lines)
              chunk_start <- NA
              for (i in (header_idx + 1):min(header_idx + 10, length(lines))) {
                if (grepl("^```\\{r", lines[i])) { chunk_start <- i; break }
              }
              if (is.na(chunk_start)) next

              # Find closing ``` (within 50 lines)
              chunk_end <- NA
              for (i in (chunk_start + 1):min(chunk_start + 50, length(lines))) {
                if (grepl("^```\\s*$", lines[i])) { chunk_end <- i; break }
              }
              if (is.na(chunk_end)) next

              # Splice: keep chunk_start (```{r...}), insert code, keep chunk_end (```)
              code_lines <- strsplit(code, "\n")[[1]]
              lines <- c(lines[1:chunk_start], code_lines, lines[chunk_end:length(lines)])
              filled <- filled + 1
            }

            # Fill text answers (process in reverse too)
            ans_nums <- sort(as.integer(gsub("[^0-9]", "", names(answers))), decreasing = TRUE)
            for (ex_num in ans_nums) {
              ex_key <- paste0("Exercise ", ex_num)
              answer <- answers[[ex_key]]
              if (is.null(answer)) next

              header_pat <- paste0("^####\\s+Exercise\\s+", ex_num, "\\b")
              header_idx <- grep(header_pat, lines)
              if (length(header_idx) == 0) next
              header_idx <- header_idx[1]

              # Insert after header, before ```{r} or next ####
              insert_at <- header_idx
              for (i in (header_idx + 1):min(header_idx + 10, length(lines))) {
                if (grepl("^```\\{r", lines[i]) || grepl("^####", lines[i])) break
                insert_at <- i
              }

              answer_lines <- strsplit(answer, "\n")[[1]]
              lines <- c(lines[1:insert_at], answer_lines, lines[(insert_at + 1):length(lines)])
            }

            # Write back to disk
            writeLines(lines, file_path)
            add_message("action", paste0("Filled ", filled, " exercises in ", basename(file_path)))

            # Tell RStudio to revert/reload the file (it detects the disk change)
            tryCatch({
              rstudioapi::sendToConsole(
                paste0('invisible(tryCatch(rstudioapi::navigateToFile("', file_path, '"), error=function(e){}))'),
                execute = TRUE, echo = FALSE, focus = FALSE
              )
            }, error = function(e) {})
            Sys.sleep(0.5)

            # Run all the code in the console
            all_code <- paste(unlist(chunks), collapse = "\n\n")
            if (nzchar(all_code)) {
              tryCatch({
                rstudioapi::sendToConsole(
                  all_code, execute = TRUE, echo = TRUE, focus = FALSE
                )
                Sys.sleep(3)
              }, error = function(e) {})
            }
          }
        }, error = function(e) {
          add_message("action", paste0("Could not fill Rmd: ", e$message))
        })

      } else if (type == "create_r_script") {
        code  <- payload$code
        title <- if (!is.null(payload$title)) payload$title else "tsifl Script"
        tryCatch({
          rstudioapi::documentNew(
            text = paste0("# ", title, "\n# Generated by tsifl\n\n", code, "\n"),
            type = "r"
          )
          add_message("action", paste0("Created script: ", title))
        }, error = function(e) {
          add_message("action", "Could not create script file")
        })
      }
    }
  }

  # ── Launch ─────────────────────────────────────────────────────────────────
  # Prevent idle timeout — keep alive indefinitely
  options(
    shiny.autoreload = FALSE,
    shiny.maxRequestSize = 50 * 1024^2,  # 50MB max upload
    shiny.idle.timeout = 0,              # Never kill idle sessions
    shiny.sanitize.errors = FALSE,       # Show real errors for debugging
    httpuv.timeout = 0                   # Never timeout httpuv connections
  )

  shiny_app <- shiny::shinyApp(
    ui = ui,
    server = server,
    options = list(
      sessionTimeout = 86400000,  # 24 hours in ms
      ws.max.idle = 86400         # WebSocket max idle: 24 hours (seconds)
    )
  )

  shiny::runApp(
    shiny_app,
    host          = "127.0.0.1",
    port          = port,
    launch.browser = FALSE,
    quiet          = FALSE
  )
}
