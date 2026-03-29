#' tsifl Background Server
#'
#' Runs the Shiny UI as a background job so the main R session stays free.
#' When the main session is free, sendToConsole() routes plots to the Plots pane.
#'
#' @keywords internal
run_tsifl_server <- function(port = 7444) {

  BACKEND_URL <- "https://focused-solace-production-6839.up.railway.app"

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
  "

  # ── UI ─────────────────────────────────────────────────────────────────────
  ui <- shiny::fluidPage(
    shiny::tags$head(shiny::tags$style(shiny::HTML(CSS))),

    shiny::div(id = "header",
      shiny::span(id = "logo", "\u26a1 tsifl"),
      shiny::uiOutput("tasks_label")
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
        accept = "image/*", multiple = "multiple",
        style = "display:none;"
      ),
      shiny::div(id = "input_actions",
        shiny::tags$button(id = "attach_btn", title = "Attach image", "+"),
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

      // File picker change
      document.getElementById('image_input').addEventListener('change', function(e) {
        var files = Array.from(e.target.files);
        files.forEach(function(f) {
          if (f.type.startsWith('image/')) readImageFile(f);
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
          if (f.type.startsWith('image/')) readImageFile(f);
        });
      });

      // Paste from clipboard
      document.querySelector('#user_input').addEventListener('paste', function(e) {
        var items = Array.from(e.clipboardData ? e.clipboardData.items : []);
        items.forEach(function(item) {
          if (item.type.startsWith('image/')) {
            var f = item.getAsFile();
            if (f) readImageFile(f);
          }
        });
      });

      // Read image file → base64
      function readImageFile(file) {
        var reader = new FileReader();
        reader.onload = function() {
          var base64 = reader.result;
          var mediaType = file.type || 'image/png';
          var data = base64.split(',')[1];
          pendingImages.push({ media_type: mediaType, data: data });
          updatePreview();
        };
        reader.readAsDataURL(file);
      }

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
          btn.title = 'Attach image';
          return;
        }
        btn.textContent = pendingImages.length;
        btn.title = pendingImages.length + ' image(s) attached';
        bar.style.display = 'flex';
        pendingImages.forEach(function(img, i) {
          var wrapper = document.createElement('div');
          wrapper.className = 'image-preview-item';
          renderToCanvas(img.data, img.media_type, 48, 48).then(function(canvas) {
            if (canvas) wrapper.insertBefore(canvas, wrapper.firstChild);
          });
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

    get_r_context <- function() {
      list(
        app         = "rstudio",
        r_version   = R.version$version.string,
        working_dir = getwd()
      )
    }

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

        all_actions <- c(
          if (!is.null(data$actions) && length(data$actions) > 0) data$actions else list(),
          if (!is.null(data$action)  && length(data$action)  > 0 &&
              !identical(data$action$type, "none")) list(data$action) else list()
        )

        for (action in all_actions) {
          execute_r_action(action, add_message)
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

        # 1. Insert into editor so user can see the code
        tryCatch({
          ctx <- rstudioapi::getSourceEditorContext()
          rstudioapi::insertText(
            location = ctx$selection[[1]]$range$end,
            text     = paste0("\n# tsifl\n", code, "\n"),
            id       = ctx$id
          )
        }, error = function(e) {})

        # 2. Send to the MAIN R console.
        #    Because this server runs in a background job, the main console
        #    is NOT blocked — sendToConsole goes straight there, graphics
        #    route normally to the Plots pane.
        sent <- tryCatch({
          rstudioapi::sendToConsole(
            code, execute = TRUE, echo = TRUE, focus = FALSE
          )
          TRUE
        }, error = function(e) FALSE)

        if (sent) {
          add_message("action", "\u2705 Running in console \u2014 check Plots pane for charts")
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
      }
    }
  }

  # ── Launch ─────────────────────────────────────────────────────────────────
  shiny::runApp(
    shiny::shinyApp(ui, server),
    host          = "127.0.0.1",
    port          = port,
    launch.browser = FALSE,
    quiet          = FALSE
  )
}
