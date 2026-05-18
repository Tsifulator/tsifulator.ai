#' Cross-platform path for tsifl's scratch files.
#'
#' On macOS/Linux this lives under `/tmp/`; on Windows it lands in
#' `tempdir()` (typically `C:/Users/<you>/AppData/Local/Temp`). The dot-
#' prefix is kept for visual consistency on Unix but is harmless on Windows.
#'
#' @param name Suffix (e.g. "pending_code.R", "last_plot.png").
#' @return Absolute file path as a character scalar.
#' @keywords internal
.tsifl_tmp <- function(name) {
  base <- if (.Platform$OS.type == "unix" && dir.exists("/tmp")) {
    "/tmp"
  } else {
    tempdir()
  }
  file.path(base, paste0(".tsifl_", name))
}

#' Install a file-based code listener in the current (main) R session.
#'
#' rstudioapi::sendToConsole() from a background job to the main RStudio
#' console is fragile — in many environments it silently times out with
#' "RStudio did not respond to rstudioapi IPC request". Rather than fight
#' that IPC channel, we use a file-based bridge:
#'   - Shiny background job writes code to /tmp/.tsifl_pending_code.R
#'   - Main R session polls that file every 0.5s via later::later()
#'   - When found: read it, eval it in .GlobalEnv with sink() capture,
#'     then write a done marker so the background job knows it's finished.
#'
#' This function MUST be called from the MAIN R session (the one the user
#' is typing in), not from the background job. tsifulator_addin() takes
#' care of that since it runs in the main session when invoked from the
#' console.
#'
#' Safe to call multiple times — subsequent calls are no-ops.
#'
#' @keywords internal
.install_tsifl_listener <- function() {
  # Idempotent: mark installed in a global and bail if already done
  if (isTRUE(getOption("tsifulator.listener_installed"))) {
    return(invisible(FALSE))
  }

  pending_file <- .tsifl_tmp("pending_code.R")
  output_file  <- .tsifl_tmp("last_output.txt")
  done_file    <- .tsifl_tmp("done.marker")
  plot_file    <- .tsifl_tmp("last_plot.png")
  # Plots directory keeps timestamped copies so the chat UI can reference
  # historical plots, and we can have both PNG (inline preview) and HTML
  # (interactive, clickable) versions side by side.
  plots_dir    <- .tsifl_tmp("plots")
  plot_meta    <- .tsifl_tmp("last_plot.json")
  try(dir.create(plots_dir, showWarnings = FALSE, recursive = TRUE), silent = TRUE)

  # Clean up any stale single-shot files from previous sessions
  for (f in c(pending_file, output_file, done_file, plot_file, plot_meta)) {
    try(unlink(f), silent = TRUE)
  }

  # Helper: prune plots dir to most recent N files.
  # NEVER prunes pinned plots — those are listed in pinned.json (a JSON
  # array of basenames, written by the Shiny UI when the user clicks the
  # pin icon on a chip). Pinned plots stay forever and are loaded by the
  # plot_files reactive first.
  .prune_plots <- function(keep = 20) {
    tryCatch({
      pinned_path <- file.path(plots_dir, "pinned.json")
      pinned <- character(0)
      if (file.exists(pinned_path)) {
        pinned <- tryCatch(
          jsonlite::fromJSON(pinned_path),
          error = function(e) character(0)
        )
        if (!is.character(pinned)) pinned <- character(0)
      }
      all_files <- list.files(plots_dir, full.names = TRUE,
                              pattern = "\\.(html|png)$")
      if (length(all_files) == 0) return(invisible())
      info <- file.info(all_files)
      info <- info[order(info$mtime, decreasing = TRUE), , drop = FALSE]
      # Drop pinned files from the prune candidates entirely
      candidates <- rownames(info)[!(basename(rownames(info)) %in% pinned)]
      if (length(candidates) <= keep) return(invisible())
      old <- candidates[(keep + 1):length(candidates)]
      try(unlink(old), silent = TRUE)
    }, error = function(e) {})
  }

  poll_fn <- function() {
    tryCatch({
      if (file.exists(pending_file)) {
        # Read + remove the pending file atomically
        code <- tryCatch(
          paste(readLines(pending_file, warn = FALSE), collapse = "\n"),
          error = function(e) ""
        )
        try(unlink(pending_file), silent = TRUE)

        if (nchar(code) > 0) {
          # Start fresh output capture
          try(unlink(output_file), silent = TRUE)
          try(unlink(done_file), silent = TRUE)

          # sink(split=TRUE) echoes to console AND captures to file
          sink_ok <- tryCatch({
            sink(output_file, split = TRUE)
            TRUE
          }, error = function(e) FALSE)

          # Auto-load the common analysis packages BEFORE running user code.
          # Eliminates the entire "could not find function 'glimpse'" class
          # of failures regardless of what code the model emits. Each
          # library() is wrapped so a missing package doesn't abort —
          # users without one of these installed just don't get it loaded.
          # Suppressed messages so the load chatter doesn't clutter output.
          .tsifl_autoload <- c(
            "dplyr", "tidyr", "readr", "tibble",
            "ggplot2", "scales", "lubridate",
            "plotly", "DT"
          )
          suppressMessages(suppressWarnings({
            for (.pkg in .tsifl_autoload) {
              tryCatch(
                requireNamespace(.pkg, quietly = TRUE) &&
                  base::library(.pkg, character.only = TRUE,
                                quietly = TRUE, warn.conflicts = FALSE),
                error = function(e) NULL
              )
            }
          }))

          # Detect plot-producing code and open a png device around execution
          plot_keywords <- c("plot(", "ggplot(", "boxplot(", "hist(",
                             "barplot(", "geom_", "abline(", "curve(",
                             "pie(", "heatmap(", "pairs(", "qqnorm(",
                             "acf(", "pacf(", "stripchart(", "par(mfrow")
          has_plot <- any(sapply(plot_keywords, function(k) grepl(k, code, fixed = TRUE)))
          # htmlwidget detection (plotly, leaflet, DT, etc.) — these need
          # both PNG thumbnail AND HTML save for interactive viewing.
          htmlwidget_keywords <- c("plot_ly(", "plotly::", "leaflet(",
                                   "datatable(", "DT::", "dygraph(",
                                   "highchart(", "visNetwork(")
          has_htmlwidget <- any(sapply(htmlwidget_keywords,
                                       function(k) grepl(k, code, fixed = TRUE)))
          if (has_plot) try(unlink(plot_file), silent = TRUE)

          # htmlwidgets (plotly, leaflet, DT) auto-print by calling
          # getOption("viewer")(url) — which routes to RStudio's Viewer
          # pane, the same pane hosting tsifl. That displaces tsifl
          # entirely. Install a no-op viewer while we execute user code so
          # the widget doesn't hijack the pane; we capture the object from
          # .Last.value / .GlobalEnv below and save it to our plots dir,
          # which tsifl's Plot tab picks up via reactivePoll.
          orig_viewer <- getOption("viewer")
          if (has_htmlwidget) {
            options(viewer = function(url, height = NULL) invisible(NULL))
          }

          # Use source() with print.eval=TRUE instead of bare eval(parse(...))
          # so top-level expressions get auto-printed the same way they do at
          # the R prompt. Without this, `ggplot(...)` creates the object but
          # never draws — the plot code appears to "run" with no visible plot.
          tryCatch({
            # source() requires a file or connection — write code to a temp
            # file so we can source it. Also gives nicer error messages.
            # Force UTF-8 on both write and read: the model frequently emits
            # non-ASCII characters (degree signs, smart quotes, em-dashes,
            # accented column names from real datasets). Without explicit
            # UTF-8 encoding, R uses the system locale, which on macOS often
            # mis-parses multi-byte sequences and throws "unexpected ','"
            # syntax errors mid-string.
            tmp_src <- tempfile(pattern = "tsifl_", fileext = ".R")
            writeLines(enc2utf8(code), tmp_src, useBytes = TRUE)
            on.exit(try(unlink(tmp_src), silent = TRUE), add = TRUE)
            source(
              tmp_src,
              local = .GlobalEnv,
              echo = FALSE,
              print.eval = TRUE,  # auto-print top-level values (renders plots)
              max.deparse.length = 500,
              encoding = "UTF-8"
            )
          }, error = function(e) cat("Error:", conditionMessage(e), "\n"),
             finally = {
               if (has_htmlwidget) {
                 try(options(viewer = orig_viewer), silent = TRUE)
               }
             })

          # Best-effort plot capture
          plot_meta_data <- NULL
          if (has_plot) {
            tryCatch({
              grDevices::dev.copy(
                grDevices::png, filename = plot_file,
                width = 1000, height = 700, res = 150
              )
              grDevices::dev.off()
              if (!file.exists(plot_file) ||
                  file.info(plot_file)$size < 200) {
                try(ggplot2::ggsave(plot_file, width = 8, height = 6, dpi = 150),
                    silent = TRUE)
              }
            }, error = function(e) {})
          }

          # Resolve PNG and widget outputs independently. A ggplot/base
          # plot produces a PNG but no widget; a plot_ly() call produces a
          # widget but no PNG. Prior versions gated the entire save block
          # on PNG existence, which silently dropped plotly-only plots.
          png_ok <- file.exists(plot_file) && file.info(plot_file)$size > 200

          widget_obj <- NULL
          if (has_htmlwidget &&
              requireNamespace("htmlwidgets", quietly = TRUE)) {
            # Prefer .Last.value — the auto-printed result of the last
            # top-level expression (usually the final plot_ly(...) call).
            lv <- tryCatch(get(".Last.value", envir = baseenv()),
                           error = function(e) NULL)
            if (!is.null(lv) &&
                (inherits(lv, "htmlwidget") ||
                 inherits(lv, "plotly") ||
                 inherits(lv, "datatables"))) {
              widget_obj <- lv
            }
            # Fall back: scan user objects in .GlobalEnv for any
            # htmlwidget so names like p3d / interactive_plot still work.
            if (is.null(widget_obj)) {
              internal_skip <- c(".tsifl_watcher", ".tsifl_capture",
                                 ".tsifl_listener_capture",
                                 ".tsifl_listener_poll")
              for (nm in setdiff(ls(.GlobalEnv), internal_skip)) {
                obj <- tryCatch(get(nm, envir = .GlobalEnv),
                                error = function(e) NULL)
                if (!is.null(obj) &&
                    (inherits(obj, "htmlwidget") ||
                     inherits(obj, "plotly") ||
                     inherits(obj, "datatables"))) {
                  widget_obj <- obj
                  break
                }
              }
            }
          }

          # Save whatever we got — PNG, HTML, or both — into the
          # timestamped plots dir, but DEDUP against every existing
          # file of the same type (not just the latest). The model
          # re-emits the same plot code on most turns; without dedup
          # the Plots tab and chat accumulate identical entries.
          all_hashes <- function(pat) {
            f <- list.files(plots_dir, pattern = pat, full.names = TRUE)
            if (length(f) == 0) return(character(0))
            tryCatch(unname(tools::md5sum(f)),
                     error = function(e) character(0))
          }

          # Try to pull a descriptive title out of the code so the chip
          # label is "Math vs Reading Scores" instead of "Plot 4:32 AM".
          extract_title <- function(src) {
            patterns <- c(
              'ggtitle\\s*\\(\\s*["\\\']([^"\\\']+)["\\\']',
              'labs\\s*\\([^)]*title\\s*=\\s*["\\\']([^"\\\']+)["\\\']',
              'plot_ly\\s*\\([^)]*title\\s*=\\s*["\\\']([^"\\\']+)["\\\']',
              'layout\\s*\\([^)]*title\\s*=\\s*["\\\']([^"\\\']+)["\\\']',
              '\\bmain\\s*=\\s*["\\\']([^"\\\']+)["\\\']'
            )
            for (pat in patterns) {
              m <- tryCatch(
                regmatches(src, regexec(pat, src, perl = TRUE))[[1]],
                error = function(e) character(0)
              )
              if (length(m) >= 2 && nzchar(m[2])) return(m[2])
            }
            NA_character_
          }
          slugify <- function(s) {
            if (is.na(s) || !nzchar(s)) return("")
            s <- tolower(s)
            s <- gsub("[^a-z0-9]+", "-", s)
            s <- gsub("^-+|-+$", "", s)
            if (nchar(s) > 50) s <- substr(s, 1, 50)
            gsub("-+$", "", s)
          }

          if (png_ok || !is.null(widget_obj)) {
            tryCatch({
              ts_tag <- format(Sys.time(), "%Y%m%d_%H%M%S")
              slug   <- slugify(extract_title(code))
              suffix <- if (nzchar(slug)) paste0("_", slug) else ""
              png_path  <- NULL
              html_path <- NULL

              if (png_ok) {
                new_hash <- tryCatch(unname(tools::md5sum(plot_file)),
                                     error = function(e) NULL)
                prev <- all_hashes("\\.png$")
                if (!is.null(new_hash) && !(new_hash %in% prev)) {
                  png_path <- file.path(plots_dir,
                                        paste0("plot_", ts_tag, suffix, ".png"))
                  file.copy(plot_file, png_path, overwrite = TRUE)
                }
              }

              if (!is.null(widget_obj)) {
                # saveWidget needs a real path — write to a tmp first so
                # we can hash it before deciding whether to keep it.
                tmp_html <- tempfile(pattern = "tsifl_widget_",
                                     fileext = ".html")
                ok <- FALSE
                tryCatch({
                  htmlwidgets::saveWidget(widget_obj, tmp_html,
                                          selfcontained = TRUE)
                  ok <- file.exists(tmp_html)
                }, error = function(e) {})
                if (ok) {
                  new_hash <- tryCatch(unname(tools::md5sum(tmp_html)),
                                       error = function(e) NULL)
                  prev <- all_hashes("\\.html$")
                  if (!is.null(new_hash) && !(new_hash %in% prev)) {
                    html_path <- file.path(plots_dir,
                                           paste0("plot_", ts_tag, suffix, ".html"))
                    file.rename(tmp_html, html_path)
                    sz <- file.info(html_path)$size
                    if (sz < 10000) {
                      cat("tsifl warning: saved widget is only", sz,
                          "bytes — likely empty. Check data filtering.\n")
                    }
                  } else {
                    try(unlink(tmp_html), silent = TRUE)
                  }
                }
              }

              # Nothing survived dedup? Don't touch the metadata file —
              # the chat renderer would otherwise re-attach a stale plot.
              if (!is.null(png_path) || !is.null(html_path)) {
                meta <- list(
                  png  = png_path,
                  html = html_path,
                  ts   = as.numeric(Sys.time()),
                  interactive = !is.null(html_path)
                )
                writeLines(
                  jsonlite::toJSON(meta, auto_unbox = TRUE, null = "null"),
                  plot_meta
                )
                plot_meta_data <- meta
                .prune_plots(keep = 20)
              }
            }, error = function(e) {})
          }

          if (sink_ok) tryCatch(sink(), error = function(e) {})

          # Signal completion — background job polls for this marker
          writeLines(as.character(Sys.time()), done_file)
        }
      }
    }, error = function(e) {})

    # Reschedule ourselves (later will not fire if R is blocked, which is fine —
    # any code we've scheduled has finished by the time later can fire again).
    later::later(poll_fn, delay = 0.5)
  }

  # Kick off the polling loop
  later::later(poll_fn, delay = 0.5)

  # ── Environment + editor capture watcher ────────────────────────────────
  # The background Shiny job needs to know:
  #   (1) what data frames / packages the user has loaded in .GlobalEnv,
  #   (2) what file is currently open in the editor (active doc contents,
  #       path, id) — so the model can answer "what does this script do"
  #       without trying to call fake rstudioapi functions.
  # Both run in MAIN here (not via sendToConsole — that IPC is fragile from
  # background jobs and frequently times out silently). We write two files:
  #   /tmp/.tsifl_env_snapshot.rds    (env + pkgs + editor; legacy combined)
  #   /tmp/.tsifl_editor_snapshot.rds (editor only; faster path the Shiny
  #                                    side reads on every Send)
  env_snapshot_file    <- .tsifl_tmp("env_snapshot.rds")
  editor_snapshot_file <- .tsifl_tmp("editor_snapshot.rds")
  editor_log_file      <- .tsifl_tmp("editor_capture.log")

  # Capture editor doc from MAIN. rstudioapi calls work natively here, so we
  # get real contents — unlike from a background job where they return empty.
  capture_editor <- function() {
    tryCatch({
      ctx <- NULL
      if (requireNamespace("rstudioapi", quietly = TRUE) &&
          rstudioapi::isAvailable()) {
        # getSourceEditorContext returns the topmost source editor, sticky
        # to last-focused source pane even when focus is on the chat panel.
        ctx <- tryCatch(
          rstudioapi::getSourceEditorContext(),
          error = function(e) NULL
        )
        # Fallback: active document context (returns whatever has focus —
        # could be Console, but worth trying if SourceEditor is empty).
        if (is.null(ctx) || length(ctx$contents) == 0 ||
            (length(ctx$contents) == 1 && !nzchar(ctx$contents))) {
          ctx <- tryCatch(
            rstudioapi::getActiveDocumentContext(),
            error = function(e) NULL
          )
        }
      }
      # Resolve path more aggressively: getSourceEditorContext sometimes
      # returns path="" for saved docs that were opened via navigateToFile
      # rather than the file menu. documentPath(id) fills that gap.
      resolved_path <- ""
      if (!is.null(ctx)) {
        resolved_path <- if (is.null(ctx$path)) "" else as.character(ctx$path)
        if (!nzchar(resolved_path) && !is.null(ctx$id) && nzchar(ctx$id)) {
          resolved_path <- tryCatch(
            as.character(rstudioapi::documentPath(ctx$id)),
            error = function(e) ""
          )
          if (is.null(resolved_path) || is.na(resolved_path)) resolved_path <- ""
        }
      }
      payload <- if (!is.null(ctx)) {
        list(
          id       = if (is.null(ctx$id))   "" else as.character(ctx$id),
          path     = resolved_path,
          contents = if (is.null(ctx$contents)) "" else paste(ctx$contents, collapse = "\n"),
          ts       = Sys.time()
        )
      } else {
        list(id = "", path = "", contents = "", ts = Sys.time(), err = "no_ctx")
      }
      saveRDS(payload, editor_snapshot_file)

      # Debug log so we can diagnose "model can't see script" complaints
      tryCatch({
        cat(
          format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
          " path=", payload$path,
          " id=", payload$id,
          " chars=", nchar(payload$contents),
          " err=", (if (is.null(payload$err)) "<none>" else payload$err),
          "\n", sep = "", file = editor_log_file, append = TRUE
        )
      }, error = function(e) NULL)

      payload
    }, error = function(e) {
      tryCatch({
        saveRDS(
          list(id = "", path = "", contents = "", ts = Sys.time(),
               err = conditionMessage(e)),
          editor_snapshot_file
        )
      }, error = function(e2) NULL)
      list(id = "", path = "", contents = "", err = conditionMessage(e))
    })
  }

  capture_env <- function() {
    editor_info <- capture_editor()
    tryCatch({
      nms <- setdiff(
        ls(.GlobalEnv),
        c(".tsifl_watcher", ".tsifl_capture",
          ".tsifl_listener_capture", ".tsifl_listener_poll")
      )
      info <- lapply(nms, function(nm) {
        obj <- tryCatch(get(nm, envir = .GlobalEnv), error = function(e) NULL)
        if (is.null(obj)) return(list(name = nm, class = "unknown"))
        r <- list(name = nm, class = paste(class(obj), collapse = ", "))
        if (!is.null(dim(obj))) r$dim <- paste(dim(obj), collapse = "x")
        if (!is.null(names(obj))) {
          r$col_names <- paste(head(names(obj), 20), collapse = ", ")
        }
        tryCatch({
          r$preview <- paste(
            utils::capture.output(
              utils::str(obj, max.level = 0, give.attr = FALSE)
            )[1],
            collapse = ""
          )
        }, error = function(e) {})
        r
      })
      pkgs <- gsub("^package:", "", grep("^package:", search(), value = TRUE))
      saveRDS(
        list(env = info, pkgs = pkgs, editor = editor_info, ts = Sys.time()),
        env_snapshot_file
      )
    }, error = function(e) {
      tryCatch(
        saveRDS(
          list(
            env = list(), pkgs = character(0), editor = editor_info,
            ts = Sys.time(), err = conditionMessage(e)
          ),
          env_snapshot_file
        ),
        error = function(e2) {}
      )
    })
    # Reschedule env capture every 3 seconds — env state changes slowly
    later::later(capture_env, delay = 3)
  }

  # Faster dedicated editor cycle — every 1 second. Editor state changes
  # whenever the user types or switches tabs, and we want the snapshot
  # fresh for the next Send without depending on the slower env cycle.
  fast_editor_cycle <- function() {
    capture_editor()
    later::later(fast_editor_cycle, delay = 1)
  }

  # Fire an immediate capture and start both recurring loops
  capture_env()
  fast_editor_cycle()

  options(tsifulator.listener_installed = TRUE)
  invisible(TRUE)
}

#' tsifl RStudio Addin
#'
#' Launches the tsifl chat panel inside RStudio as a background job,
#' keeping the main R session free so plots appear in the Plots pane.
#'
#' @export
tsifulator_addin <- function() {

  # Install the file-based code bridge in THIS (main) session. This must
  # happen before the background job starts so the listener is ready to
  # receive code as soon as the user chats.
  tryCatch(.install_tsifl_listener(), error = function(e) {
    message("tsifl: warning — could not install console listener: ",
            conditionMessage(e))
  })

  PORT <- 7444

  # ── Check if a LIVE tsifl is already serving ──────────────────────────────
  # A raw open socket isn't enough (could be a zombie from a dead background
  # job). Actually GET the page and check for a 2xx HTTP response with a
  # short timeout. If anything looks off, kill the port and restart clean.
  is_live <- tryCatch({
    resp <- httr2::request(paste0("http://127.0.0.1:", PORT)) |>
      httr2::req_timeout(2) |>
      httr2::req_error(is_error = function(r) FALSE) |>
      httr2::req_perform()
    status <- httr2::resp_status(resp)
    status >= 200 && status < 500
  }, error = function(e) FALSE)

  if (is_live) {
    rstudioapi::viewer(paste0("http://127.0.0.1:", PORT))
    return(invisible(NULL))
  }

  # Anything holding the port but not responding is a zombie — nuke it.
  port_in_use <- tryCatch({
    con <- suppressWarnings(socketConnection(
      host = "127.0.0.1", port = PORT,
      open = "r+", blocking = TRUE, timeout = 1
    ))
    close(con); TRUE
  }, error = function(e) FALSE)
  if (port_in_use) {
    message("tsifl: stale process on port ", PORT, " — cleaning up...")
    if (.Platform$OS.type == "unix") {
      tryCatch(
        system(paste0("lsof -ti:", PORT, " | xargs kill -9 2>/dev/null"),
               intern = FALSE, ignore.stdout = TRUE, ignore.stderr = TRUE),
        error = function(e) {}
      )
      Sys.sleep(0.5)
    } else {
      # Windows: no portable way to kill by port without an extra dependency.
      # Tell the user what to do.
      message("tsifl: on Windows, please close RStudio fully and reopen, ",
              "or kill the R process holding port ", PORT, " via Task Manager.")
    }
  }

  # ── Write a tiny launcher script for the background job ───────────────────
  job_script <- tempfile(fileext = ".R")
  writeLines(
    paste0(
      'tsifulator:::run_tsifl_server(', PORT, ')'
    ),
    job_script
  )

  # ── Launch as RStudio background job ──────────────────────────────────────
  rstudioapi::jobRunScript(
    path       = job_script,
    name       = "\u26a1 tsifl",
    workingDir = getwd()
  )

  # ── Wait for the server to start (faster polling: 0.25s intervals, 5s max)
  message("tsifl: starting...")
  started <- FALSE
  for (i in seq_len(20)) {
    Sys.sleep(0.25)
    ok <- tryCatch({
      con <- url(paste0("http://127.0.0.1:", PORT), open = "r")
      close(con)
      TRUE
    }, error = function(e) FALSE)
    if (ok) { started <- TRUE; break }
  }

  if (!started) {
    message("tsifl: server took longer than expected \u2014 try again in a moment.")
    return(invisible(NULL))
  }

  # ── Open the UI in the RStudio Viewer pane ────────────────────────────────
  rstudioapi::viewer(paste0("http://127.0.0.1:", PORT))
  message("tsifl: ready \u2713")

  invisible(NULL)
}

#' Launch tsifl
#'
#' Convenience wrapper — same as using the Addins menu.
#'
#' @export
run_tsifl <- function() {
  tsifulator_addin()
}
