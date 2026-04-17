#' One-command setup for tsifl
#'
#' Gets a new user from zero to running tsifl in their RStudio with a single
#' function call. Checks R version, installs missing packages, writes the
#' auto-launch hook to the user's ~/.Rprofile, kills any stale tsifl process,
#' verifies backend connectivity, and launches the chat panel.
#'
#' Safe to run multiple times — the .Rprofile hook is tagged with markers so
#' repeated calls update rather than duplicate.
#'
#' @param auto_launch Logical. If TRUE (default), launches the tsifl chat
#'   panel immediately after setup completes. Set to FALSE for headless setup
#'   (CI, scripting, etc.).
#' @param install_missing Logical. If TRUE (default), automatically installs
#'   missing required packages. If FALSE, only reports what's missing.
#' @param quiet Logical. If TRUE, suppresses progress messages. Default FALSE.
#'
#' @return Invisibly returns a list with the setup results.
#' @export
#' @examples
#' \dontrun{
#' # Standard setup: install deps, write hook, launch
#' tsifulator::setup()
#'
#' # Headless setup (e.g. for scripting)
#' tsifulator::setup(auto_launch = FALSE, quiet = TRUE)
#' }
setup <- function(auto_launch = TRUE, install_missing = TRUE, quiet = FALSE) {
  .say <- function(msg, mark = "  ") {
    if (!quiet) message(mark, msg)
  }
  .ok   <- function(msg) .say(msg, "\u2713 ")   # check mark
  .warn <- function(msg) .say(msg, "\u26a0 ")   # warning sign
  .fail <- function(msg) .say(msg, "\u2717 ")   # cross mark

  if (!quiet) message("\n\u26a1 tsifl setup starting...\n")

  results <- list(
    r_version_ok     = FALSE,
    in_rstudio       = FALSE,
    rstudio_api_ok   = FALSE,
    packages_ok      = FALSE,
    missing_packages = character(0),
    rprofile_ok      = FALSE,
    rprofile_path    = "",
    backend_ok       = FALSE,
    port_cleared     = FALSE,
    launched         = FALSE,
    warnings         = character(0),
    errors           = character(0)
  )

  # ── Step 1: R version check ────────────────────────────────────────────────
  r_ver <- getRversion()
  results$r_version_ok <- r_ver >= "4.0.0"
  if (results$r_version_ok) {
    .ok(sprintf("R %s detected", as.character(r_ver)))
  } else {
    .fail(sprintf("R %s detected \u2014 tsifl requires R >= 4.0.0", r_ver))
    results$errors <- c(results$errors, "R version below 4.0.0")
    return(invisible(results))
  }

  # ── Step 2: RStudio detection ──────────────────────────────────────────────
  in_rs <- tryCatch(
    isTRUE(requireNamespace("rstudioapi", quietly = TRUE)) &&
      rstudioapi::isAvailable(),
    error = function(e) FALSE
  )
  results$in_rstudio     <- in_rs
  results$rstudio_api_ok <- in_rs
  if (in_rs) {
    rs_ver <- tryCatch(
      as.character(rstudioapi::versionInfo()$version),
      error = function(e) "unknown"
    )
    .ok(sprintf("RStudio %s (rstudioapi available)", rs_ver))
  } else {
    .warn("Not running inside RStudio. tsifl is designed for RStudio; some features won't work in plain R.")
    results$warnings <- c(results$warnings, "not running in RStudio")
  }

  # ── Step 3: Required packages ──────────────────────────────────────────────
  required_pkgs <- c(
    "shiny", "httr2", "jsonlite", "rstudioapi",
    "base64enc", "later", "grDevices", "png", "tools"
  )
  missing <- required_pkgs[!vapply(
    required_pkgs,
    function(p) requireNamespace(p, quietly = TRUE),
    logical(1)
  )]
  results$missing_packages <- missing

  if (length(missing) == 0) {
    .ok("All required packages installed")
    results$packages_ok <- TRUE
  } else if (install_missing) {
    .say(sprintf("Installing %d missing package(s): %s",
                 length(missing), paste(missing, collapse = ", ")))
    install_ok <- tryCatch({
      utils::install.packages(missing, quiet = quiet)
      TRUE
    }, error = function(e) {
      .fail(paste("Package install failed:", conditionMessage(e)))
      results$errors <<- c(results$errors, paste("install:", conditionMessage(e)))
      FALSE
    })
    if (install_ok) {
      # Re-check
      still_missing <- missing[!vapply(
        missing,
        function(p) requireNamespace(p, quietly = TRUE),
        logical(1)
      )]
      if (length(still_missing) == 0) {
        .ok("All packages now installed")
        results$packages_ok <- TRUE
      } else {
        .fail(sprintf("Could not install: %s",
                      paste(still_missing, collapse = ", ")))
        results$missing_packages <- still_missing
        results$errors <- c(results$errors,
                            sprintf("still missing: %s",
                                    paste(still_missing, collapse = ", ")))
      }
    }
  } else {
    .warn(sprintf("Missing packages (install_missing=FALSE): %s",
                  paste(missing, collapse = ", ")))
    results$warnings <- c(results$warnings, "missing required packages")
  }

  # ── Step 4: Kill any stale tsifl on port 7444 ──────────────────────────────
  port_kill_ok <- tryCatch({
    port <- 7444
    # First check if anything is listening
    con <- suppressWarnings(try(
      socketConnection(host = "127.0.0.1", port = port,
                       open = "r+", blocking = TRUE, timeout = 1),
      silent = TRUE
    ))
    if (!inherits(con, "try-error")) {
      try(close(con), silent = TRUE)
      # Port is in use — try to kill the process (macOS/Linux only)
      if (.Platform$OS.type == "unix") {
        system(sprintf("lsof -ti:%d | xargs kill -9 2>/dev/null", port),
               intern = FALSE, ignore.stdout = TRUE, ignore.stderr = TRUE)
        Sys.sleep(0.5)
        .ok(sprintf("Stale process on port %d cleared", port))
      } else {
        .warn(sprintf("Port %d is in use but auto-kill only works on macOS/Linux", port))
      }
    }
    TRUE
  }, error = function(e) FALSE)
  results$port_cleared <- port_kill_ok

  # ── Step 5: Write the auto-launch hook to ~/.Rprofile ──────────────────────
  rprofile_path <- path.expand("~/.Rprofile")
  results$rprofile_path <- rprofile_path

  hook_content <- c(
    "# ── Tsifulator.ai auto-launch (managed by tsifulator::setup) ─ DO NOT EDIT ───",
    "# Remove by running tsifulator::teardown() or deleting everything between",
    "# these BEGIN/END markers.",
    "local({",
    "  if (!interactive()) return()",
    "  setHook(\"rstudio.sessionInit\", function(newSession) {",
    "    if (!newSession) return()",
    "    if (!requireNamespace(\"tsifulator\", quietly = TRUE)) return()",
    "    if (!requireNamespace(\"rstudioapi\", quietly = TRUE)) return()",
    "    if (!requireNamespace(\"later\",      quietly = TRUE)) return()",
    "    if (!rstudioapi::isAvailable()) return()",
    "    # Delay via later::later so the IDE is fully ready before launch",
    "    later::later(function() {",
    "      port_free <- tryCatch({",
    "        con <- suppressWarnings(socketConnection(",
    "          host = \"127.0.0.1\", port = 7444,",
    "          open = \"r+\", blocking = TRUE, timeout = 1",
    "        ))",
    "        close(con); FALSE",
    "      }, error = function(e) TRUE)",
    "      try({",
    "        if (port_free) {",
    "          tsifulator::tsifulator_addin()",
    "        } else {",
    "          rstudioapi::viewer(\"http://127.0.0.1:7444\")",
    "        }",
    "      }, silent = TRUE)",
    "    }, delay = 6)",
    "  }, action = \"append\")",
    "})",
    "# ── End tsifulator.ai auto-launch ───────────────────────────────────────────"
  )

  BEGIN_MARK <- "# ── Tsifulator.ai auto-launch (managed by tsifulator::setup)"
  END_MARK   <- "# ── End tsifulator.ai auto-launch"

  existing <- if (file.exists(rprofile_path)) {
    tryCatch(readLines(rprofile_path, warn = FALSE), error = function(e) character(0))
  } else {
    character(0)
  }

  begin_idx <- grep(BEGIN_MARK, existing, fixed = TRUE)
  end_idx   <- grep(END_MARK, existing, fixed = TRUE)

  new_lines <- if (length(begin_idx) == 1 && length(end_idx) == 1 && end_idx > begin_idx) {
    # Replace existing block
    c(
      existing[seq_len(begin_idx - 1)],
      hook_content,
      if (end_idx < length(existing)) existing[(end_idx + 1):length(existing)] else character(0)
    )
  } else {
    # Append new block (with a blank separator if file had content)
    c(existing,
      if (length(existing) > 0 && nzchar(tail(existing, 1))) "" else NULL,
      hook_content)
  }

  rprofile_ok <- tryCatch({
    writeLines(new_lines, rprofile_path)
    TRUE
  }, error = function(e) {
    .fail(paste("Could not write ~/.Rprofile:", conditionMessage(e)))
    results$errors <<- c(results$errors, paste("rprofile:", conditionMessage(e)))
    FALSE
  })
  results$rprofile_ok <- rprofile_ok
  if (rprofile_ok) {
    action_word <- if (length(begin_idx) == 1) "updated" else "installed"
    .ok(sprintf("Auto-launch hook %s in %s", action_word, rprofile_path))
  }

  # ── Step 6: Backend connectivity check ─────────────────────────────────────
  backend_url <- if (nchar(Sys.getenv("TSIFULATOR_BACKEND_URL")) > 0) {
    Sys.getenv("TSIFULATOR_BACKEND_URL")
  } else {
    "https://focused-solace-production-6839.up.railway.app"
  }
  backend_ok <- tryCatch({
    resp <- httr2::request(backend_url) |>
      httr2::req_url_path_append("health") |>
      httr2::req_timeout(5) |>
      httr2::req_error(is_error = function(r) FALSE) |>
      httr2::req_perform()
    status <- httr2::resp_status(resp)
    status >= 200 && status < 300
  }, error = function(e) FALSE)
  results$backend_ok <- backend_ok
  if (backend_ok) {
    .ok(sprintf("Backend reachable (%s)", backend_url))
  } else {
    .warn(sprintf("Backend not reachable at %s — you can still use tsifl when it comes back online",
                  backend_url))
    results$warnings <- c(results$warnings, "backend unreachable")
  }

  # ── Step 7: Launch ─────────────────────────────────────────────────────────
  if (auto_launch && results$in_rstudio && results$packages_ok) {
    launched <- tryCatch({
      tsifulator_addin()
      TRUE
    }, error = function(e) {
      .warn(sprintf("Could not auto-launch: %s", conditionMessage(e)))
      FALSE
    })
    results$launched <- isTRUE(launched)
    if (isTRUE(launched)) .ok("tsifl launched in Viewer pane")
  }

  # ── Summary ────────────────────────────────────────────────────────────────
  if (!quiet) {
    message("")
    if (length(results$errors) == 0) {
      message("\u2713 tsifl setup complete.")
      if (!results$launched && results$in_rstudio) {
        message("  Run `tsifulator::tsifulator_addin()` or use the Addins menu to launch.")
      }
      message("  On future RStudio starts, tsifl will auto-launch ~6 seconds after the session is ready.")
    } else {
      message("\u2717 tsifl setup finished with errors. See `tsifulator::status()` for details.")
    }
    message("")
  }

  invisible(results)
}


#' Remove tsifl's auto-launch hook and stop running instance
#'
#' Undoes what `setup()` did: removes the auto-launch block from
#' `~/.Rprofile` and kills any tsifl process listening on port 7444. The R
#' package itself is NOT uninstalled — use `remove.packages("tsifulator")`
#' for that.
#'
#' @param quiet Logical. If TRUE, suppresses progress messages.
#'
#' @return Invisibly returns a list of what was removed.
#' @export
teardown <- function(quiet = FALSE) {
  .say <- function(msg, mark = "  ") {
    if (!quiet) message(mark, msg)
  }
  .ok <- function(msg) .say(msg, "\u2713 ")

  if (!quiet) message("\n\u26a1 tsifl teardown...\n")

  rprofile_path <- path.expand("~/.Rprofile")
  results <- list(rprofile_cleaned = FALSE, port_killed = FALSE)

  # Remove the managed block from .Rprofile
  BEGIN_MARK <- "# ── Tsifulator.ai auto-launch (managed by tsifulator::setup)"
  END_MARK   <- "# ── End tsifulator.ai auto-launch"

  if (file.exists(rprofile_path)) {
    existing <- tryCatch(readLines(rprofile_path, warn = FALSE), error = function(e) character(0))
    begin_idx <- grep(BEGIN_MARK, existing, fixed = TRUE)
    end_idx   <- grep(END_MARK, existing, fixed = TRUE)
    if (length(begin_idx) == 1 && length(end_idx) == 1 && end_idx > begin_idx) {
      new_lines <- c(
        existing[seq_len(begin_idx - 1)],
        if (end_idx < length(existing)) existing[(end_idx + 1):length(existing)] else character(0)
      )
      # Trim trailing blank lines
      while (length(new_lines) > 0 && !nzchar(tail(new_lines, 1))) {
        new_lines <- new_lines[-length(new_lines)]
      }
      tryCatch({
        writeLines(new_lines, rprofile_path)
        results$rprofile_cleaned <- TRUE
        .ok("Auto-launch hook removed from ~/.Rprofile")
      }, error = function(e) {
        if (!quiet) message("\u2717 Could not write ~/.Rprofile: ", conditionMessage(e))
      })
    } else {
      .ok("No tsifl hook found in ~/.Rprofile (already clean)")
    }
  }

  # Kill anything on port 7444
  if (.Platform$OS.type == "unix") {
    tryCatch({
      system("lsof -ti:7444 | xargs kill -9 2>/dev/null",
             intern = FALSE, ignore.stdout = TRUE, ignore.stderr = TRUE)
      results$port_killed <- TRUE
      .ok("Killed any tsifl processes on port 7444")
    }, error = function(e) {})
  }

  if (!quiet) {
    message("")
    message("\u2713 tsifl teardown complete.")
    message("  To uninstall the package: remove.packages(\"tsifulator\")")
    message("")
  }

  invisible(results)
}


#' Print tsifl status / diagnostics
#'
#' Reports what's currently configured, what's running, what's missing.
#' Useful for debugging setup issues or answering "is tsifl working?".
#'
#' @return Invisibly returns a list of status fields.
#' @export
status <- function() {
  rprofile_path <- path.expand("~/.Rprofile")
  BEGIN_MARK <- "# ── Tsifulator.ai auto-launch (managed by tsifulator::setup)"

  rprofile_has_hook <- if (file.exists(rprofile_path)) {
    any(grepl(BEGIN_MARK,
              readLines(rprofile_path, warn = FALSE),
              fixed = TRUE))
  } else FALSE

  port_in_use <- tryCatch({
    con <- suppressWarnings(socketConnection(
      host = "127.0.0.1", port = 7444,
      open = "r+", blocking = TRUE, timeout = 1
    ))
    close(con); TRUE
  }, error = function(e) FALSE)

  backend_url <- if (nchar(Sys.getenv("TSIFULATOR_BACKEND_URL")) > 0) {
    Sys.getenv("TSIFULATOR_BACKEND_URL")
  } else {
    "https://focused-solace-production-6839.up.railway.app"
  }
  backend_ok <- tryCatch({
    resp <- httr2::request(backend_url) |>
      httr2::req_url_path_append("health") |>
      httr2::req_timeout(5) |>
      httr2::req_error(is_error = function(r) FALSE) |>
      httr2::req_perform()
    s <- httr2::resp_status(resp)
    s >= 200 && s < 300
  }, error = function(e) FALSE)

  listener_installed <- isTRUE(getOption("tsifulator.listener_installed"))

  required_pkgs <- c("shiny", "httr2", "jsonlite", "rstudioapi",
                     "base64enc", "later", "grDevices", "png", "tools")
  missing_pkgs <- required_pkgs[!vapply(
    required_pkgs,
    function(p) requireNamespace(p, quietly = TRUE),
    logical(1)
  )]

  in_rstudio <- tryCatch(rstudioapi::isAvailable(), error = function(e) FALSE)

  pkg_version <- tryCatch(as.character(utils::packageVersion("tsifulator")),
                          error = function(e) "unknown")

  message("\n\u26a1 tsifl status\n")
  message(sprintf("  Package version:       %s", pkg_version))
  message(sprintf("  R version:             %s", as.character(getRversion())))
  message(sprintf("  Running in RStudio:    %s", if (in_rstudio) "yes" else "no"))
  message(sprintf("  Required packages:     %s",
                  if (length(missing_pkgs) == 0) "all installed"
                  else paste("MISSING:", paste(missing_pkgs, collapse = ", "))))
  message(sprintf("  .Rprofile hook:        %s",
                  if (rprofile_has_hook) "installed" else "not installed"))
  message(sprintf("  Listener in session:   %s",
                  if (listener_installed) "yes" else "no"))
  message(sprintf("  tsifl running (7444):  %s",
                  if (port_in_use) "yes" else "no"))
  message(sprintf("  Backend reachable:     %s",
                  if (backend_ok) sprintf("yes (%s)", backend_url) else "no"))
  message("")

  invisible(list(
    package_version    = pkg_version,
    r_version          = as.character(getRversion()),
    in_rstudio         = in_rstudio,
    missing_packages   = missing_pkgs,
    rprofile_has_hook  = rprofile_has_hook,
    listener_installed = listener_installed,
    port_in_use        = port_in_use,
    backend_ok         = backend_ok,
    backend_url        = backend_url
  ))
}
