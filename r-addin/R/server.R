#' tsifl Background Server
#'
#' Runs the Shiny UI as a background job so the main R session stays free.
#' When the main session is free, sendToConsole() routes plots to the Plots pane.
#'
#' @keywords internal
run_tsifl_server <- function(port = 7444) {

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

    body {
      background: #FFFFFF;
      color: #1E293B;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
      font-size: 13px;
      height: 100vh;
      overflow: hidden;
    }

    #header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 9px 14px;
      background: #FFFFFF;
      border-bottom: 1px solid #E2E8F0;
      flex-shrink: 0;
    }

    #logo {
      font-weight: 700;
      font-size: 13px;
      color: #0D5EAF;
      letter-spacing: -0.3px;
    }

    #tasks_label {
      font-size: 10px;
      color: #64748B;
      background: #F1F5F9;
      padding: 2px 8px;
      border-radius: 10px;
      border: 1px solid #E2E8F0;
    }

    #chat_history {
      height: calc(100vh - 145px);
      overflow-y: auto;
      padding: 10px 12px;
      display: flex;
      flex-direction: column;
      gap: 7px;
      background: #F8FAFC;
    }

    #chat_history::-webkit-scrollbar { width: 3px; }
    #chat_history::-webkit-scrollbar-track { background: transparent; }
    #chat_history::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 3px; }

    .msg-user {
      background: #EBF3FB;
      border-left: 2px solid #0D5EAF;
      padding: 7px 10px;
      border-radius: 5px;
      line-height: 1.5;
      font-size: 13px;
      color: #1E293B;
    }

    .msg-assistant {
      background: #FFFFFF;
      border-left: 2px solid #86EFAC;
      padding: 7px 10px;
      border-radius: 5px;
      line-height: 1.5;
      font-size: 13px;
      color: #1E293B;
      box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    .msg-action {
      background: rgba(22, 163, 74, 0.07);
      border-left: 2px solid #16A34A;
      padding: 5px 10px;
      border-radius: 5px;
      font-family: 'SF Mono', 'Fira Code', Monaco, monospace;
      font-size: 11px;
      color: #16A34A;
      line-height: 1.4;
      white-space: pre-wrap;
    }

    #input_area {
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      background: #FFFFFF;
      border-top: 1px solid #E2E8F0;
      padding: 8px 12px;
      display: flex;
      flex-direction: column;
      gap: 5px;
      transition: background 0.15s, outline 0.15s;
    }

    #user_input {
      width: 100%;
      background: #F8FAFC;
      color: #1E293B;
      border: 1px solid #E2E8F0;
      border-radius: 5px;
      padding: 7px 10px;
      font-size: 13px;
      font-family: inherit;
      resize: none;
      outline: none;
      transition: border-color 0.15s, box-shadow 0.15s;
      line-height: 1.4;
    }

    #user_input:focus {
      border-color: #0D5EAF;
      box-shadow: 0 0 0 3px rgba(13, 94, 175, 0.08);
      background: #FFFFFF;
    }

    #user_input::placeholder { color: #94A3B8; }

    #input_actions {
      display: flex;
      gap: 5px;
    }

    #attach_btn {
      background: #F1F5F9;
      color: #64748B;
      border: 1px solid #E2E8F0;
      border-radius: 5px;
      padding: 7px 12px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s;
      line-height: 1;
      flex-shrink: 0;
    }

    #attach_btn:hover {
      background: #EBF3FB;
      color: #0D5EAF;
      border-color: #0D5EAF;
    }

    #send_btn {
      flex: 1;
      background: #0D5EAF;
      color: white;
      border: none;
      border-radius: 5px;
      padding: 7px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: background 0.15s;
      letter-spacing: 0.2px;
    }

    #send_btn:hover { background: #0A4896; }

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
      border-radius: 4px;
      border: 1px solid #E2E8F0;
    }

    .image-preview-item .remove-img {
      position: absolute;
      top: -4px;
      right: -4px;
      width: 16px;
      height: 16px;
      background: #DC2626;
      color: white;
      border: none;
      border-radius: 50%;
      font-size: 10px;
      line-height: 16px;
      text-align: center;
      cursor: pointer;
      padding: 0;
    }

    .chat-canvas {
      border-radius: 8px;
      border: 1px solid #E2E8F0;
      display: block;
      max-width: 100%;
      margin-top: 6px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }

    .image-badge {
      display: inline-block;
      background: rgba(13, 94, 175, 0.09);
      color: #0D5EAF;
      font-size: 10px;
      font-weight: 600;
      padding: 2px 8px;
      border-radius: 10px;
      margin-top: 4px;
      border: 1px solid rgba(13, 94, 175, 0.2);
    }

    #status_bar {
      font-size: 10px;
      color: #94A3B8;
      padding: 1px 0;
    }

    .typing-indicator { display: flex; gap: 4px; padding: 8px 10px; align-items: center; }
    .typing-indicator span { width: 7px; height: 7px; background: #94A3B8; border-radius: 50%; animation: pulse 1.4s infinite ease-in-out; }
    .typing-indicator span:nth-child(2) { animation-delay: 0.2s; }
    .typing-indicator span:nth-child(3) { animation-delay: 0.4s; }
    @keyframes pulse { 0%, 80%, 100% { transform: scale(0.6); opacity: 0.4; } 40% { transform: scale(1); opacity: 1; } }
    @keyframes fadeInUp { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
    .message { animation: fadeInUp 0.3s ease; }

    .message pre { position: relative; background: #1E293B; color: #E2E8F0; border-radius: 6px; padding: 10px 12px; margin: 6px 0; font-family: 'SF Mono', Consolas, monospace; font-size: 12px; line-height: 1.5; overflow-x: auto; white-space: pre-wrap; }
    .message code { background: #F1F5F9; padding: 1px 4px; border-radius: 3px; font-size: 11px; font-family: 'SF Mono', Consolas, monospace; }

    @media (prefers-color-scheme: dark) {
      body { background: #0F172A; color: #F1F5F9; }
      #header { background: #1E293B; border-color: #334155; }
      .message.user { background: #1E3A5F; border-color: #0D5EAF; }
      .message.assistant { background: #1E293B; border-color: #334155; color: #F1F5F9; }
      #input_area { background: #0F172A; border-color: #334155; }
      textarea { background: #1E293B !important; color: #F1F5F9 !important; border-color: #334155 !important; }
    }
  "

  # ── UI ─────────────────────────────────────────────────────────────────────
  ui <- shiny::fluidPage(
    shiny::tags$head(
      shiny::tags$style(shiny::HTML(CSS)),
      shiny::tags$script(shiny::HTML("
        // Keep-alive ping every 30 seconds to prevent WebSocket idle disconnect
        setInterval(function() {
          if (Shiny && Shiny.setInputValue) {
            Shiny.setInputValue('_keepalive', Date.now(), {priority: 'event'});
          }
        }, 30000);

        // Reconnect if WebSocket drops
        $(document).on('shiny:disconnected', function(event) {
          setTimeout(function() { location.reload(); }, 3000);
        });
      "))
    ),

    shiny::div(id = "header",
      shiny::span(id = "logo", "\u26a1 tsifl"),
      shiny::div(style = "display:flex;align-items:center;gap:6px;",
        shiny::tags$button(
          id = "notes_btn",
          onclick = paste0("window.open('", BACKEND_URL, "/notes-app','_blank')"),
          style = "background:#F1F5F9;color:#64748B;border:1px solid #E2E8F0;border-radius:4px;padding:2px 8px;font-size:10px;font-weight:600;cursor:pointer;",
          "Notes"
        ),
        shiny::uiOutput("tasks_label")
      )
    ),

    # Quick actions (Improvements 60, 62, 63, 68)
    shiny::div(id = "quick_actions", style = "display:flex;gap:4px;padding:6px 12px;flex-wrap:wrap;border-bottom:1px solid #E2E8F0;",
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 't_test', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "t-test"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'linear_reg', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "Linear Reg"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'anova', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "ANOVA"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'ggplot_scatter', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "ggplot"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'summary_stats', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "Summary"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'correlation', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "Correlation"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'which_test', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "Which test?"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'profile_data', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "Profile data"),
      shiny::tags$button(class = "quick-btn", onclick = "Shiny.setInputValue('quick_action', 'compare_models', {priority: 'event'})", style = "background:#F8FAFC;color:#64748B;border:1px solid #E2E8F0;border-radius:10px;padding:2px 8px;font-size:10px;cursor:pointer;", "Compare models")
    ),

    shiny::div(id = "chat_history",
      shiny::uiOutput("chat_messages")
    ),

    shiny::div(id = "input_area",
      shiny::div(id = "image_preview_bar"),
      shiny::textAreaInput("user_input", label = NULL,
        placeholder = "Ask me to run R code, plot data, update Excel...",
        rows = 2, width = "100%"
      ),
      shiny::tags$input(type = "file", id = "image_input",
        accept = "image/*,.pdf,.csv,.txt,.json,.xml,.r,.R,.py,.js,.ts,.sql,.md,.html,.yaml,.yml,.docx,.xlsx,.sas,.do,.log",
        multiple = "multiple",
        style = "display:none;"
      ),
      shiny::div(id = "input_actions",
        shiny::tags$button(id = "attach_btn", title = "Attach file", "+"),
        shiny::actionButton("send_btn", "Send", width = "100%")
      ),
      shiny::div(id = "status_bar", shiny::textOutput("status", inline = TRUE))
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

      // Paste from clipboard
      document.querySelector('#user_input').addEventListener('paste', function(e) {
        var items = Array.from(e.clipboardData ? e.clipboardData.items : []);
        items.forEach(function(item) {
          if (item.type.startsWith('image/') || item.kind === 'file') {
            var f = item.getAsFile();
            if (f) readFileAsBase64(f);
          }
        });
      });

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

      // Hook into send button — pass images to Shiny before submit
      document.getElementById('send_btn').addEventListener('click', function() {
        if (pendingImages.length > 0) {
          Shiny.setInputValue('pending_images', JSON.stringify(pendingImages));
          pendingImages = [];
          updatePreview();
        } else {
          Shiny.setInputValue('pending_images', '[]');
        }
      }, true);  // capture phase — runs BEFORE Shiny's handler

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
    "))

  )

  # ── Server ─────────────────────────────────────────────────────────────────
  server <- function(input, output, session) {

    messages    <- shiny::reactiveVal(list())
    tasks_left  <- shiny::reactiveVal(NA)
    status_text <- shiny::reactiveVal("Connected")

    output$chat_messages <- shiny::renderUI({
      msgs <- messages()
      if (length(msgs) == 0) {
        return(shiny::p("Ask me anything...",
          style = "color:#2a3f5f; font-style:italic; font-size:12px; padding:4px 0;"))
      }
      ui_list <- lapply(msgs, function(m) {
        children <- list(shiny::span(m$text))
        # Add image containers for messages with images
        if (!is.null(m$images) && length(m$images) > 0) {
          for (i in seq_along(m$images)) {
            img_id <- m$img_ids[[i]]
            children <- c(children, list(
              shiny::div(id = img_id, style = "margin-top:6px;")
            ))
            # Send message to JS to render canvas
            session$sendCustomMessage("show_chat_image", list(
              target_id  = img_id,
              data       = m$images[[i]]$data,
              media_type = m$images[[i]]$media_type
            ))
          }
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

    output$status <- shiny::renderText(status_text())

    img_counter <- shiny::reactiveVal(0)

    add_message <- function(role, text, images = NULL) {
      current <- messages()
      entry <- list(role = role, text = text)
      if (!is.null(images) && length(images) > 0) {
        # Assign unique IDs for each image so JS can find the DOM target
        img_ids <- lapply(seq_along(images), function(i) {
          n <- img_counter() + i
          paste0("chat_img_", n)
        })
        img_counter(img_counter() + length(images))
        entry$images <- images
        entry$img_ids <- img_ids
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

      # Wait for the snapshot, with retries
      snap <- NULL
      for (attempt in 1:4) {
        Sys.sleep(0.4)
        snap <- tryCatch(readRDS(ENV_SNAPSHOT_FILE), error = function(e) NULL)
        if (!is.null(snap)) break
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
        if (nchar(ctx$path) > 0) {
          doc_info$active_file <- basename(ctx$path)
        }
        if (length(ctx$contents) > 0) {
          doc_info$active_preview <- paste(utils::head(ctx$contents, 15), collapse = "\n")
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
        shiny::updateTextAreaInput(session, "user_input", value = prompt)
        # Trigger send
        shinyjs_msg <- paste0("$('#send_btn').click();")
      }
    }, ignoreInit = TRUE)

    shiny::observeEvent(input$send_btn, {
      msg <- trimws(input$user_input)
      if (nchar(msg) == 0) return()

      # Capture pending images from JS
      images_json <- input$pending_images
      images <- list()
      if (!is.null(images_json) && nchar(images_json) > 2) {
        images <- jsonlite::fromJSON(images_json, simplifyVector = FALSE)
      }

      shiny::updateTextAreaInput(session, "user_input", value = "")
      add_message("user", msg, images = if (length(images) > 0) images else NULL)
      status_text("Thinking...")

      # Build request body
      body <- list(
        user_id = USER_ID,
        message = msg,
        context = get_r_context()
      )
      if (length(images) > 0) {
        body$images <- images
      }

      tryCatch({
        resp <- httr2::request(BACKEND_URL) |>
          httr2::req_url_path_append("chat", "") |>
          httr2::req_headers("Content-Type" = "application/json") |>
          httr2::req_body_json(body) |>
          httr2::req_options(ssl_verifypeer = 0) |>
          httr2::req_perform()

        data <- httr2::resp_body_json(resp)
        add_message("assistant", data$reply)

        if (!is.null(data$tasks_remaining) && data$tasks_remaining >= 0)
          tasks_left(data$tasks_remaining)

        all_actions <- list()
        # Safely collect actions array
        if (!is.null(data$actions) && is.list(data$actions) && length(data$actions) > 0) {
          all_actions <- c(all_actions, data$actions)
        }
        # Safely collect single action (must be a list with $type)
        if (!is.null(data$action) && is.list(data$action) && !is.null(data$action$type) &&
            !identical(data$action$type, "none")) {
          all_actions <- c(all_actions, list(data$action))
        }

        for (action in all_actions) {
          tryCatch({
            execute_r_action(action, add_message)
          }, error = function(e) {
            add_message("action", paste0("\u26a0\ufe0f Action error: ", e$message))
          })
        }

        status_text("Done")

      }, error = function(e) {
        add_message("assistant",
          paste0("\u26a0\ufe0f Could not reach backend.\n", e$message))
        status_text("Disconnected")
      })
    })

    # ── Action executor ──────────────────────────────────────────────────────
    execute_r_action <- function(action, add_message) {
      type    <- action$type
      payload <- action$payload

      if (type == "run_r_code") {
        code <- payload$code

        # 1. Create a NEW R script document with the generated code
        #    so it appears in a fresh tab (not buried in existing file)
        tryCatch({
          # Create new untitled R script
          rstudioapi::documentNew(
            text = paste0("# tsifl — Generated Code\n\n", code, "\n"),
            type = "r"
          )
        }, error = function(e) {
          # Fallback: try inserting into current editor
          tryCatch({
            ctx <- rstudioapi::getSourceEditorContext()
            rstudioapi::insertText(
              location = ctx$selection[[1]]$range$end,
              text     = paste0("\n# tsifl\n", code, "\n"),
              id       = ctx$id
            )
          }, error = function(e2) {})
        })

        # 2. Send to the MAIN R console.
        #    Because this server runs in a background job, the main console
        #    is NOT blocked — sendToConsole goes straight there, graphics
        #    route normally to the Plots pane.
        #    Auto-install missing packages (Improvement 67)
        #    Wrap code with tryCatch that detects missing packages and installs them
        wrapped_code <- paste0(
          'tryCatch({ ', code, ' }, error = function(e) { ',
          '  msg <- conditionMessage(e); ',
          '  if (grepl("there is no package called", msg)) { ',
          '    pkg <- sub(".*there is no package called .(.+?).", "\\\\1", msg); ',
          '    message("Auto-installing package: ", pkg); ',
          '    install.packages(pkg, repos = "https://cran.r-project.org", quiet = TRUE); ',
          '    eval(parse(text = paste0("library(", pkg, ")"))); ',
          '    eval(parse(text = ', deparse(code), ')); ',
          '  } else { stop(e) } })'
        )
        sent <- tryCatch({
          rstudioapi::sendToConsole(
            code, execute = TRUE, echo = TRUE, focus = FALSE
          )
          TRUE
        }, error = function(e) FALSE)

        if (sent) {
          add_message("action", "\u2705 Running in console \u2014 check Plots pane for charts")

          # ── Cross-app memory: capture R output and data snapshots ──────────
          tryCatch({
            # Build a capture script that runs in the main session
            capture_script_path <- "/tmp/.tsifl_capture_output.R"
            writeLines(c(
              'tryCatch({',
              paste0('  .tsifl_code <- ', deparse(code), ';'),
              '  .tsifl_output <- paste(utils::capture.output(eval(parse(text = .tsifl_code))), collapse = "\\n");',
              '  .tsifl_output <- substr(.tsifl_output, 1, 5000);',
              '  writeLines(.tsifl_output, "/tmp/.tsifl_last_output.txt");',
              '  # Snapshot data frames',
              '  .tsifl_df_info <- list();',
              '  for (.tsifl_nm in ls(.GlobalEnv)) {',
              '    .tsifl_obj <- tryCatch(get(.tsifl_nm, envir = .GlobalEnv), error = function(e) NULL);',
              '    if (is.data.frame(.tsifl_obj) || inherits(.tsifl_obj, "tbl_df") || inherits(.tsifl_obj, "data.table")) {',
              '      .tsifl_csv_path <- paste0("/tmp/", .tsifl_nm, ".csv");',
              '      tryCatch(utils::write.csv(.tsifl_obj, .tsifl_csv_path, row.names = FALSE), error = function(e) {});',
              '      .tsifl_df_info[[.tsifl_nm]] <- list(',
              '        name = .tsifl_nm,',
              '        nrow = nrow(.tsifl_obj),',
              '        ncol = ncol(.tsifl_obj),',
              '        columns = paste(names(.tsifl_obj), collapse = ", "),',
              '        csv_path = .tsifl_csv_path',
              '      );',
              '    }',
              '  };',
              '  saveRDS(.tsifl_df_info, "/tmp/.tsifl_df_info.rds");',
              '  rm(list = grep("^\\\\.tsifl_", ls(), value = TRUE));',
              '}, error = function(e) {})'
            ), capture_script_path)

            # Run the capture script in the main session
            Sys.sleep(1)  # Wait for the main code to finish
            tryCatch(
              rstudioapi::sendToConsole(
                'tryCatch(source("/tmp/.tsifl_capture_output.R", local = TRUE, echo = FALSE), error = function(e) {})',
                execute = TRUE, echo = FALSE, focus = FALSE
              ),
              error = function(e) {}
            )

            Sys.sleep(1.5)  # Wait for capture to complete

            # POST r_output to transfer/store
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

            # POST data_snapshot for each data frame
            df_info <- tryCatch(readRDS("/tmp/.tsifl_df_info.rds"), error = function(e) list())
            for (df_meta in df_info) {
              transfer_body <- list(
                from_app  = "rstudio",
                to_app    = "any",
                data_type = "data_snapshot",
                data      = paste0("Data frame '", df_meta$name, "': ",
                                   df_meta$nrow, " rows x ", df_meta$ncol, " cols. ",
                                   "Columns: ", df_meta$columns),
                metadata  = list(
                  name     = df_meta$name,
                  nrow     = df_meta$nrow,
                  ncol     = df_meta$ncol,
                  columns  = df_meta$columns,
                  csv_path = df_meta$csv_path
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

          # Auto-export plot: inject a save command into the MAIN console
          # (dev.copy doesn't work from background job — no plot device here)
          plot_keywords <- c("plot(", "ggplot(", "boxplot(", "hist(", "barplot(",
                            "geom_", "abline(", "curve(", "pie(", "heatmap(",
                            "pairs(", "qqnorm(", "acf(", "pacf(", "stripchart(")
          has_plot <- any(sapply(plot_keywords, function(kw) grepl(kw, code, fixed = TRUE)))
          if (has_plot) {
            # Tell the main console to save the current plot to a temp file
            plot_path <- file.path(tempdir(), ".tsifl_last_plot.png")
            save_cmd <- sprintf(
              'tryCatch({ grDevices::dev.copy(grDevices::png, "%s", width=800, height=600, res=150); grDevices::dev.off() }, error=function(e){})',
              plot_path
            )
            Sys.sleep(2)  # Wait for plot to render in main session
            tryCatch({
              rstudioapi::sendToConsole(save_cmd, execute = TRUE, echo = FALSE, focus = FALSE)
            }, error = function(e) {})

            # Now wait a moment for the file to be written, then read it
            Sys.sleep(1)
            tryCatch({
              if (file.exists(plot_path) && file.info(plot_path)$size > 0) {
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

                add_message("action", "\U0001f4e4 Plot exported \u2014 say 'paste R plot' in Excel to insert it")
              }
            }, error = function(e) {
              # Silently skip — plot capture is best-effort
            })
          }
        } else {
          add_message("action", "\u26a0\ufe0f Could not send to console")
        }

      } else if (type == "install_package") {
        pkg <- payload$package
        add_message("action", paste0("Installing: ", pkg))
        rstudioapi::sendToConsole(
          paste0('install.packages("', pkg, '")'),
          execute = TRUE, echo = TRUE, focus = FALSE
        )
        add_message("action", paste0("\u2705 Installing ", pkg, " in console"))

      } else if (type == "export_plot") {
        # Save current plot via main console, then upload from background job
        tryCatch({
          plot_path <- file.path(tempdir(), ".tsifl_export_plot.png")
          save_cmd <- sprintf(
            'tryCatch({ grDevices::dev.copy(grDevices::png, "%s", width=800, height=600, res=150); grDevices::dev.off() }, error=function(e){})',
            plot_path
          )
          rstudioapi::sendToConsole(save_cmd, execute = TRUE, echo = FALSE, focus = FALSE)
          Sys.sleep(1.5)

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
            add_message("action", paste0("\u2705 Plot exported (ID: ", tid, ") \u2014 say 'paste R plot' in Excel"))
          } else {
            add_message("action", "\u26a0\ufe0f No plot found in Plots pane. Generate a plot first.")
          }
        }, error = function(e) {
          add_message("action", paste0("\u26a0\ufe0f Could not export plot: ", e$message))
        })

      } else if (type == "create_r_script") {
        code  <- payload$code
        title <- if (!is.null(payload$title)) payload$title else "tsifl Script"
        tryCatch({
          rstudioapi::documentNew(
            text = paste0("# ", title, "\n# Generated by tsifl\n\n", code, "\n"),
            type = "r"
          )
          add_message("action", paste0("\u2705 Created script: ", title))
        }, error = function(e) {
          add_message("action", "\u26a0\ufe0f Could not create script file")
        })
      }
    }
  }

  # ── Launch ─────────────────────────────────────────────────────────────────
  # Set options to prevent idle timeout — keep the Shiny app alive indefinitely
  options(
    shiny.autoreload = FALSE,
    shiny.maxRequestSize = 50 * 1024^2  # 50MB max upload
  )

  shiny_app <- shiny::shinyApp(
    ui = ui,
    server = server,
    options = list(
      # Disable session idle timeout (default is 0 which means disconnect on WebSocket close)
      # Setting to very large number keeps session alive
      sessionTimeout = 0
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
