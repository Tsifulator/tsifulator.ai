#' Ollama (local model) brain for tsifl — dev mode
#'
#' Routes chat requests to a local Ollama server instead of the hosted
#' backend. Free. No API credits burned. Slower, less accurate, but
#' perfect for iterating on the R add-in itself without paying for every
#' test cycle.
#'
#' Enable: options(tsifulator.brain = "ollama")
#' Choose model: options(tsifulator.ollama_model = "llama3.1:8b")
#' Override URL: Sys.setenv(TSIFULATOR_OLLAMA_URL = "http://localhost:11434")
#'
#' Disable (back to Anthropic): options(tsifulator.brain = NULL)
#'
#' Recommended models for R coding (any of these work):
#'   - llama3.1:8b  (4.7GB, fast, decent at structured output)
#'   - qwen2.5:7b   (4.7GB, great at code, best for tool calling)
#'   - llama3.1:70b (40GB, much better but only on big machines)
#'
#' First-time setup:
#'   1. Install Ollama from https://ollama.com
#'   2. ollama pull llama3.1:8b   (or your preferred model)
#'   3. In R: options(tsifulator.brain = "ollama")
#'   4. Use tsifl as normal — no API credits burned
NULL


#' Check if the local Ollama server is reachable
#' @keywords internal
.tsifl_ollama_available <- function() {
  url <- Sys.getenv("TSIFULATOR_OLLAMA_URL", "http://localhost:11434")
  tryCatch({
    r <- httr2::request(url) |>
      httr2::req_timeout(2) |>
      httr2::req_error(is_error = function(r) FALSE) |>
      httr2::req_perform()
    httr2::resp_status(r) < 500
  }, error = function(e) FALSE)
}


#' Is the user currently in Ollama (dev) mode?
#'
#' Checks `options(tsifulator.brain)` first, then `TSIFULATOR_BRAIN` env var.
#' @keywords internal
.tsifl_using_ollama <- function() {
  opt <- getOption("tsifulator.brain", NULL)
  if (!is.null(opt) && identical(tolower(opt), "ollama")) return(TRUE)
  env <- Sys.getenv("TSIFULATOR_BRAIN", "")
  identical(tolower(env), "ollama")
}


#' Build the system prompt for Ollama's R code generation
#'
#' Smaller, more rigid than the hosted-backend prompt because local
#' models need explicit structure. Always asks for a JSON response so
#' parsing is reliable.
#' @keywords internal
.tsifl_ollama_system_prompt <- function(r_context = list()) {
  # Summarise what's loaded so the model uses exact column names
  env_summary <- ""
  env_obj <- r_context$env_objects
  if (!is.null(env_obj) && length(env_obj) > 0) {
    parts <- vapply(env_obj, function(o) {
      base <- paste0("- `", o$name, "` (", o$class, ")")
      if (!is.null(o$dim) && nzchar(o$dim)) {
        base <- paste0(base, " [", o$dim, "]")
      }
      if (!is.null(o$col_names) && nzchar(o$col_names)) {
        base <- paste0(base, " columns: ", o$col_names)
      }
      base
    }, character(1))
    env_summary <- paste0(
      "\nDATA CURRENTLY LOADED IN R:\n",
      paste(parts, collapse = "\n"),
      "\n\nUse the EXACT column names listed above. NEVER invent column names.\n"
    )
  }

  paste0(
    "You are tsifl, an R coding assistant inside RStudio.\n\n",
    "Respond ONLY with a single JSON object in this exact shape:\n",
    '{"reply": "one-sentence explanation", "code": "complete R code"}\n\n',
    "Rules for the code field:\n",
    "1. ALWAYS start with: suppressMessages({library(readr); library(dplyr); library(ggplot2); library(plotly); library(tidyr); library(tibble)})\n",
    "2. For CSVs use readr::read_delim(path, show_col_types = FALSE) — it auto-detects delimiter. Call glimpse(df) right after.\n",
    "3. Use exact column names from the DATA section below. Backtick-quote names with special chars (e.g. `Team&Contract`).\n",
    "4. For plots use ggplot2 (static) or plotly (interactive). For htmlwidgets ALWAYS pass selfcontained = FALSE to saveWidget.\n",
    "5. For 3D / interactive plots use plot_ly(...).\n",
    "6. NEVER reference a column you can't see in the DATA section. If you don't see it, ask the user instead.\n\n",
    "Rules for the reply field:\n",
    "- One sentence, plain English. NO markdown. NO code in the reply field.\n",
    "- Tell the user what the code does. Not how it does it.\n\n",
    "Do NOT wrap your response in markdown fences. Do NOT add commentary before or after the JSON. Just the JSON object.\n",
    env_summary
  )
}


#' Extract a JSON object from a possibly-noisy model response
#'
#' Local models often pad JSON with markdown fences, preambles, or
#' apology text. This pulls the first balanced {...} block and parses it.
#' @keywords internal
.tsifl_extract_json <- function(text) {
  if (!is.character(text) || !nzchar(text)) return(NULL)
  # First try direct parse (clean responses)
  result <- tryCatch(jsonlite::fromJSON(text, simplifyVector = FALSE),
                    error = function(e) NULL)
  if (!is.null(result)) return(result)
  # Strip markdown fences
  cleaned <- gsub("```json\\s*", "", text)
  cleaned <- gsub("```\\s*", "", cleaned)
  result <- tryCatch(jsonlite::fromJSON(cleaned, simplifyVector = FALSE),
                    error = function(e) NULL)
  if (!is.null(result)) return(result)
  # Find first balanced {...} block via brace counting
  chars <- strsplit(text, "", fixed = TRUE)[[1]]
  start <- NA_integer_
  depth <- 0L
  for (i in seq_along(chars)) {
    if (chars[i] == "{") {
      if (is.na(start)) start <- i
      depth <- depth + 1L
    } else if (chars[i] == "}") {
      depth <- depth - 1L
      if (depth == 0L && !is.na(start)) {
        candidate <- paste(chars[start:i], collapse = "")
        parsed <- tryCatch(jsonlite::fromJSON(candidate, simplifyVector = FALSE),
                           error = function(e) NULL)
        if (!is.null(parsed)) return(parsed)
        start <- NA_integer_  # reset, try next brace
      }
    }
  }
  NULL
}


#' Call the local Ollama server and return a response in the backend's shape
#'
#' Translates the backend's `{reply, action, actions}` contract so the
#' rest of the R add-in works unchanged. Skips the retry chain and
#' interpretation phase that the hosted path uses — those would cost
#' Ollama compute too and local models are slower per call.
#'
#' @keywords internal
.tsifl_call_ollama <- function(user_msg, r_context = list(), images = list()) {
  url <- Sys.getenv("TSIFULATOR_OLLAMA_URL", "http://localhost:11434")
  model <- getOption("tsifulator.ollama_model", "llama3.1:8b")

  system_prompt <- .tsifl_ollama_system_prompt(r_context)

  body <- list(
    model = model,
    stream = FALSE,
    options = list(
      temperature = 0.2,    # keep code generation tight
      num_ctx = 8192        # enough room for env_objects + history
    ),
    messages = list(
      list(role = "system", content = system_prompt),
      list(role = "user",   content = user_msg)
    )
  )

  resp <- tryCatch({
    httr2::request(url) |>
      httr2::req_url_path("/api/chat") |>
      httr2::req_body_json(body) |>
      httr2::req_timeout(180) |>
      httr2::req_perform()
  }, error = function(e) {
    return(list(
      reply = paste0(
        "Could not reach Ollama at ", url, " — is it running?\n",
        "Start with: ollama serve   (in a terminal)\n",
        "Then pull a model: ollama pull ", model, "\n",
        "Error: ", conditionMessage(e)
      ),
      actions = list()
    ))
  })

  if (is.list(resp) && !is.null(resp$reply)) return(resp)  # short-circuit error

  data <- tryCatch(httr2::resp_body_json(resp, simplifyVector = FALSE),
                   error = function(e) NULL)
  if (is.null(data) || is.null(data$message) || is.null(data$message$content)) {
    return(list(
      reply = "Ollama responded but the message body was empty.",
      actions = list()
    ))
  }

  raw <- data$message$content
  parsed <- .tsifl_extract_json(raw)

  if (is.null(parsed) || is.null(parsed$reply)) {
    return(list(
      reply = paste0(
        "Ollama returned text I couldn't parse as JSON. Try a smaller / ",
        "more focused prompt, or switch model with ",
        "options(tsifulator.ollama_model = \"qwen2.5:7b\").\n\n",
        "Raw response (first 400 chars):\n",
        substr(raw, 1, 400)
      ),
      actions = list()
    ))
  }

  # Build response in the backend's shape
  result <- list(reply = parsed$reply, actions = list())
  if (!is.null(parsed$code) && is.character(parsed$code) && nchar(parsed$code) > 0) {
    result$actions <- list(list(
      type = "run_r_code",
      payload = list(code = parsed$code)
    ))
  }
  # Surface the model used so the user can see what answered
  result$model_used <- paste0("ollama:", model)
  result$tasks_remaining <- -1L  # unlimited (it's free)
  result
}
