#' tsifl RStudio Addin
#'
#' Launches the tsifl chat panel inside RStudio as a background job,
#' keeping the main R session free so plots appear in the Plots pane.
#'
#' @export
tsifulator_addin <- function() {

  PORT <- 7444

  # ── Check if already running ───────────────────────────────────────────────
  already_up <- tryCatch({
    con <- url(paste0("http://127.0.0.1:", PORT), open = "r")
    close(con)
    TRUE
  }, error = function(e) FALSE)

  if (already_up) {
    rstudioapi::viewer(paste0("http://127.0.0.1:", PORT))
    return(invisible(NULL))
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
  # The job runs the Shiny server in a SEPARATE R process.
  # The main R session stays completely free:
  #   - sendToConsole() from the job goes to the main console
  #   - plots in the main console go straight to the Plots pane
  rstudioapi::jobRunScript(
    path       = job_script,
    name       = "\u26a1 tsifl",
    workingDir = getwd()
  )

  # ── Wait for the server to start (up to 8 seconds) ────────────────────────
  message("tsifl: starting background server...")
  started <- FALSE
  for (i in seq_len(16)) {
    Sys.sleep(0.5)
    ok <- tryCatch({
      con <- url(paste0("http://127.0.0.1:", PORT), open = "r")
      close(con)
      TRUE
    }, error = function(e) FALSE)
    if (ok) { started <- TRUE; break }
  }

  if (!started) {
    message("tsifl: server took longer than expected \u2014 try running Addins > tsifl again in a moment.")
    return(invisible(NULL))
  }

  # ── Open the UI in the RStudio Viewer pane ────────────────────────────────
  rstudioapi::viewer(paste0("http://127.0.0.1:", PORT))
  message("tsifl: ready. Your R console is free \u2014 plots will appear in the Plots pane.")

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
