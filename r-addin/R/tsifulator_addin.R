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

  pending_file <- "/tmp/.tsifl_pending_code.R"
  output_file  <- "/tmp/.tsifl_last_output.txt"
  done_file    <- "/tmp/.tsifl_done.marker"
  plot_file    <- "/tmp/.tsifl_last_plot.png"
  # Plots directory keeps timestamped copies so the chat UI can reference
  # historical plots, and we can have both PNG (inline preview) and HTML
  # (interactive, clickable) versions side by side.
  plots_dir    <- "/tmp/.tsifl_plots"
  plot_meta    <- "/tmp/.tsifl_last_plot.json"
  try(dir.create(plots_dir, showWarnings = FALSE, recursive = TRUE), silent = TRUE)

  # Clean up any stale single-shot files from previous sessions
  for (f in c(pending_file, output_file, done_file, plot_file, plot_meta)) {
    try(unlink(f), silent = TRUE)
  }

  # Helper: prune plots dir to most recent N files
  .prune_plots <- function(keep = 20) {
    tryCatch({
      all_files <- list.files(plots_dir, full.names = TRUE)
      if (length(all_files) <= keep) return(invisible())
      info <- file.info(all_files)
      info <- info[order(info$mtime, decreasing = TRUE), , drop = FALSE]
      old <- rownames(info)[(keep + 1):nrow(info)]
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
            tmp_src <- tempfile(pattern = "tsifl_", fileext = ".R")
            writeLines(code, tmp_src)
            on.exit(try(unlink(tmp_src), silent = TRUE), add = TRUE)
            source(
              tmp_src,
              local = .GlobalEnv,
              echo = FALSE,
              print.eval = TRUE,  # auto-print top-level values (renders plots)
              max.deparse.length = 500
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

  # ── Environment capture watcher ──────────────────────────────────────────
  # The background Shiny job needs to know what data frames / packages the
  # user has loaded in their main R session's .GlobalEnv (so the model
  # doesn't hallucinate code against nonexistent objects). Prior to 0.3.0
  # this watcher was installed via rstudioapi::sendToConsole — but we moved
  # away from that API because it's unreliable. Install it here in the main
  # session directly, so env snapshots always flow regardless of IPC state.
  env_snapshot_file <- "/tmp/.tsifl_env_snapshot.rds"

  capture_env <- function() {
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
        list(env = info, pkgs = pkgs, ts = Sys.time()),
        env_snapshot_file
      )
    }, error = function(e) {
      tryCatch(
        saveRDS(
          list(
            env = list(), pkgs = character(0),
            ts = Sys.time(), err = conditionMessage(e)
          ),
          env_snapshot_file
        ),
        error = function(e2) {}
      )
    })
    # Reschedule every 3 seconds
    later::later(capture_env, delay = 3)
  }

  # Fire an immediate capture and start the recurring loop
  capture_env()

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
