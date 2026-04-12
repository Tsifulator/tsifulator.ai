#' tsifl RStudio Addin
#'
#' Launches the tsifl chat panel inside RStudio as a background job,
#' keeping the main R session free so plots appear in the Plots pane.
#'
#' @export
tsifulator_addin <- function() {

  PORT <- 7444

  # ── Check if already running & responsive ──────────────────────────────────
  already_up <- tryCatch({
    con <- url(paste0("http://127.0.0.1:", PORT), open = "r")
    close(con)
    TRUE
  }, error = function(e) FALSE)

  if (already_up) {
    # Verify it's actually responsive (not a zombie port)
    responsive <- tryCatch({
      resp <- httr2::request(paste0("http://127.0.0.1:", PORT)) |>
        httr2::req_timeout(3) |>
        httr2::req_perform()
      TRUE
    }, error = function(e) FALSE)

    if (responsive) {
      rstudioapi::viewer(paste0("http://127.0.0.1:", PORT))
      return(invisible(NULL))
    } else {
      # Zombie process — kill it so we can restart fresh
      message("tsifl: stale session detected, restarting...")
      tryCatch(system(paste0("lsof -ti:", PORT, " | xargs kill -9"), intern = TRUE), error = function(e) {})
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
