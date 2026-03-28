#' tsifl RStudio Addin
#'
#' Launches the Tsifulator chat panel inside RStudio.
#' Connects to the shared backend brain — same memory as Excel.
#'
#' @export
tsifulator_addin <- function() {

  BACKEND_URL <- "https://focused-solace-production-6839.up.railway.app"

  # Read user ID from shared config (written by Excel add-in after login)
  config_path <- path.expand("~/.tsifulator_user")
  USER_ID <- if (file.exists(config_path)) {
    trimws(readLines(config_path, n = 1, warn = FALSE))
  } else if (nchar(Sys.getenv("TSIFULATOR_USER_ID")) > 0) {
    Sys.getenv("TSIFULATOR_USER_ID")
  } else {
    "dev-user-001"
  }

  # ── Design tokens — light/white + Greek flag blue ─────────────────────────
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
      height: calc(100vh - 125px);
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

    #send_btn {
      width: 100%;
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

    #status_bar {
      font-size: 10px;
      color: #94A3B8;
      padding: 1px 0;
    }
  "

  # ── UI ──────────────────────────────────────────────────────────────────────
  ui <- shiny::fluidPage(
    shiny::tags$head(shiny::tags$style(shiny::HTML(CSS))),

    # Header
    shiny::div(id = "header",
      shiny::span(id = "logo", "\u26a1 tsifl"),
      shiny::uiOutput("tasks_label")
    ),

    # Chat history
    shiny::div(id = "chat_history",
      shiny::uiOutput("chat_messages")
    ),

    # Input area
    shiny::div(id = "input_area",
      shiny::textAreaInput("user_input", label = NULL,
        placeholder = "Ask me to run R code, plot data, update Excel...",
        rows = 2, width = "100%"
      ),
      shiny::actionButton("send_btn", "Send", width = "100%"),
      shiny::div(id = "status_bar", shiny::textOutput("status", inline = TRUE))
    )
  )

  # ── Server ──────────────────────────────────────────────────────────────────
  server <- function(input, output, session) {

    messages     <- shiny::reactiveVal(list())
    tasks_left   <- shiny::reactiveVal(NA)
    status_text  <- shiny::reactiveVal("Connected")

    output$chat_messages <- shiny::renderUI({
      msgs <- messages()
      if (length(msgs) == 0) {
        return(shiny::p("Ask me anything...",
          style = "color:#2a3f5f; font-style:italic; font-size:12px; padding:4px 0;"))
      }
      lapply(msgs, function(m) {
        shiny::div(class = paste0("msg-", m$role), m$text)
      })
    })

    output$tasks_label <- shiny::renderUI({
      t <- tasks_left()
      if (is.na(t)) return(shiny::span(id = "tasks_label", ""))
      shiny::span(id = "tasks_label", paste(t, "tasks left"))
    })

    output$status <- shiny::renderText(status_text())

    shiny::observeEvent(input$send_btn, {
      msg <- trimws(input$user_input)
      if (nchar(msg) == 0) return()

      shiny::updateTextAreaInput(session, "user_input", value = "")
      add_message("user", msg)
      status_text("Reading R environment...")

      r_context <- get_r_context()
      status_text("Thinking...")

      tryCatch({
        resp <- httr2::request(BACKEND_URL) |>
          httr2::req_url_path_append("chat", "") |>
          httr2::req_headers("Content-Type" = "application/json") |>
          httr2::req_body_json(list(
            user_id = USER_ID,
            message = msg,
            context = r_context
          )) |>
          httr2::req_options(ssl_verifypeer = 0) |>
          httr2::req_perform()

        data <- httr2::resp_body_json(resp)
        add_message("assistant", data$reply)

        if (!is.null(data$tasks_remaining) && data$tasks_remaining >= 0) {
          tasks_left(data$tasks_remaining)
        }

        all_actions <- c(
          if (!is.null(data$actions) && length(data$actions) > 0) data$actions else list(),
          if (!is.null(data$action) && length(data$action) > 0 &&
              !identical(data$action$type, "none")) list(data$action) else list()
        )

        for (action in all_actions) {
          execute_r_action(action, session, add_message)
        }

        status_text("Done")

      }, error = function(e) {
        add_message("assistant",
          paste0("\u26a0\ufe0f Could not reach backend.\n", e$message))
        status_text("Disconnected")
      })
    })

    add_message <- function(role, text) {
      current <- messages()
      messages(c(current, list(list(role = role, text = text))))
    }

    get_r_context <- function() {
      env_vars <- ls(envir = .GlobalEnv)
      var_info <- lapply(env_vars[seq_len(min(20, length(env_vars)))], function(v) {
        obj <- get(v, envir = .GlobalEnv)
        list(
          name    = v,
          class   = class(obj)[1],
          dim     = paste(dim(obj), collapse = "x"),
          preview = if (is.data.frame(obj)) {
            paste(names(obj), collapse = ", ")
          } else {
            tryCatch(as.character(obj)[1], error = function(e) "?")
          }
        )
      })

      list(
        app         = "rstudio",
        r_version   = R.version$version.string,
        loaded_pkgs = paste((.packages()), collapse = ", "),
        env_objects = var_info,
        working_dir = getwd()
      )
    }

    execute_r_action <- function(action, session, add_message) {
      type    <- action$type
      payload <- action$payload

      if (type == "run_r_code") {
        code <- payload$code

        tryCatch({
          ctx <- rstudioapi::getSourceEditorContext()
          insert_text <- paste0("\n# tsifl\n", code, "\n")
          rstudioapi::insertText(
            location = ctx$selection[[1]]$range$end,
            text     = insert_text,
            id       = ctx$id
          )
          add_message("action", "\u270f\ufe0f  Inserted into editor")
        }, error = function(e) { })

        add_message("action", paste0("Running:\n", code))
        result <- tryCatch({
          output <- capture.output(eval(parse(text = code), envir = .GlobalEnv))
          paste(output, collapse = "\n")
        }, error = function(e) paste0("Error: ", e$message))

        if (nchar(trimws(result)) > 0) {
          add_message("action", paste0("Output:\n", result))
        } else {
          add_message("action", "\u2705 Done (check Plots pane)")
        }

      } else if (type == "install_package") {
        pkg <- payload$package
        add_message("action", paste0("Installing: ", pkg))
        install.packages(pkg, quiet = TRUE)
        add_message("action", paste0("\u2705 ", pkg, " installed"))
      }
    }

  }

  shiny::runGadget(ui, server,
    viewer = shiny::paneViewer(minHeight = 400)
  )
}
