#' Get the user's stored Anthropic API key, if any.
#'
#' Stored at `~/.tsifulator_api_key` (created with mode 600 on Unix).
#' Falls back to the `ANTHROPIC_API_KEY` env var if present.
#'
#' @return Character scalar (the key) or NA if not configured.
#' @export
get_api_key <- function() {
  path <- path.expand("~/.tsifulator_api_key")
  if (file.exists(path)) {
    key <- tryCatch(trimws(readLines(path, n = 1, warn = FALSE)),
                    error = function(e) "")
    if (nzchar(key)) return(key)
  }
  env_key <- Sys.getenv("ANTHROPIC_API_KEY", unset = "")
  if (nzchar(env_key)) return(env_key)
  NA_character_
}


#' Set the Anthropic API key tsifl will use.
#'
#' tsifl is bring-your-own-key (BYOK): every chat goes through your
#' Anthropic account, not ours. This means you pay Anthropic directly
#' (~$1-3/M input tokens depending on model), tsifl never sees your
#' billing, and your data goes only to Anthropic — not to our servers.
#'
#' Get a key at https://console.anthropic.com/settings/keys
#'
#' @param key Character scalar. If NULL, prompts interactively.
#' @return Invisibly TRUE on success.
#' @export
set_api_key <- function(key = NULL) {
  path <- path.expand("~/.tsifulator_api_key")
  if (is.null(key)) {
    if (!interactive()) {
      stop("set_api_key() needs to be called interactively, or pass `key=` directly.")
    }
    message("\nGet a key at https://console.anthropic.com/settings/keys")
    message("It starts with 'sk-ant-' and looks like a long random string.\n")
    key <- readline("Paste your Anthropic API key: ")
  }
  key <- trimws(key)
  if (!nzchar(key)) {
    message("No key provided. Skipping.")
    return(invisible(FALSE))
  }
  if (!startsWith(key, "sk-ant-")) {
    warning("This doesn't look like an Anthropic API key (expected to start with 'sk-ant-'). Saving anyway.")
  }
  tryCatch({
    writeLines(key, path)
    if (.Platform$OS.type == "unix") {
      Sys.chmod(path, mode = "0600")
    }
    message("✓ Saved to ", path)
  }, error = function(e) {
    stop("Could not write key file: ", conditionMessage(e))
  })

  # Auto-launch the tsifl panel if we're in RStudio and nothing's running
  # on port 7444. Saves users from having to remember a second command
  # after setting the key — the most common point of friction in setup.
  in_rs <- tryCatch(
    requireNamespace("rstudioapi", quietly = TRUE) &&
      rstudioapi::isAvailable(),
    error = function(e) FALSE
  )
  if (interactive() && in_rs) {
    port_busy <- tryCatch({
      con <- suppressWarnings(socketConnection(
        host = "127.0.0.1", port = 7444,
        open = "r+", blocking = TRUE, timeout = 1
      ))
      close(con); TRUE
    }, error = function(e) FALSE)
    if (!port_busy) {
      message("Starting tsifl panel...")
      try(tsifulator_addin(), silent = TRUE)
    } else {
      message("tsifl is already running — your new key will be used on the next message.")
    }
  }

  invisible(TRUE)
}


#' Remove the saved Anthropic API key.
#' @export
unset_api_key <- function() {
  path <- path.expand("~/.tsifulator_api_key")
  if (file.exists(path)) {
    try(unlink(path), silent = TRUE)
    message("✓ Removed ", path)
  } else {
    message("No key file to remove.")
  }
  invisible(TRUE)
}


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
    api_key_set      = FALSE,
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

  # ── Step 3b: Anthropic API key (BYOK) ──────────────────────────────────────
  api_key <- get_api_key()
  if (is.na(api_key)) {
    if (interactive() && install_missing) {
      message("")
      message("tsifl is bring-your-own-key — calls go through YOUR Anthropic account.")
      message("Get a key at https://console.anthropic.com/settings/keys (free signup,")
      message("you'll need to add a small credit balance to actually use it).")
      message("")
      ans <- tolower(trimws(readline("Set your API key now? [Y/n]: ")))
      if (ans %in% c("", "y", "yes")) {
        tryCatch(set_api_key(), error = function(e) {
          .warn(paste("Could not save key:", conditionMessage(e)))
          results$warnings <<- c(results$warnings, "api_key not saved")
        })
        api_key <- get_api_key()
      }
    }
    if (is.na(api_key)) {
      .warn("No Anthropic key set — tsifl chat won't work until you run tsifulator::set_api_key()")
      results$warnings <- c(results$warnings, "no api_key configured")
    }
  } else {
    .ok("Anthropic API key configured")
  }
  results$api_key_set <- !is.na(api_key)

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
    "    # Delay via later::later so the IDE is fully ready before launch.",
    "    # ALWAYS call tsifulator_addin() — it handles dedup internally AND",
    "    # installs the main-session listener (without which code can't run).",
    "    # Previously we short-circuited to viewer() when port 7444 looked",
    "    # busy, which skipped listener install and left users stuck.",
    "    later::later(function() {",
    "      try(tsifulator::tsifulator_addin(), silent = TRUE)",
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


#' Cleanly reinstall tsifl from a local source directory
#'
#' Works around the common failure mode where `devtools::install()` silently
#' fails because the Shiny background job has the package files locked. This
#' function does the whole dance in order:
#'   1. Stop any running tsifl server (kill port 7444)
#'   2. Unload the tsifulator namespace so R releases its file handles
#'   3. `remove.packages("tsifulator")` to wipe the stale install
#'   4. `install.packages()` or `devtools::install()` from the source dir
#'   5. Verify the installed version matches what's on disk
#'   6. Warn loudly if the version didn't actually change
#'
#' Still safer to run from a fresh Rscript --vanilla session if anything
#' persists — the .Rprofile auto-launch hook can re-lock the package if it
#' fires during the install.
#'
#' @param source_dir Path to the r-addin source directory.
#' @param quiet Logical. If TRUE, suppresses progress messages.
#'
#' @return Invisibly returns the new installed version string.
#' @export
reinstall <- function(source_dir = NULL, quiet = FALSE) {
  .say <- function(msg, mark = "  ") if (!quiet) message(mark, msg)
  .ok <- function(msg) .say(msg, "\u2713 ")
  .warn <- function(msg) .say(msg, "\u26a0 ")

  if (!quiet) message("\n\u26a1 tsifl reinstall...\n")

  # 1. Kill any running tsifl
  if (.Platform$OS.type == "unix") {
    tryCatch({
      system("lsof -ti:7444 | xargs kill -9 2>/dev/null",
             intern = FALSE, ignore.stdout = TRUE, ignore.stderr = TRUE)
      .ok("Stopped tsifl on port 7444")
    }, error = function(e) {})
  }

  # 2. Unload namespace to release file handles
  tryCatch({
    if ("tsifulator" %in% loadedNamespaces()) {
      unloadNamespace("tsifulator")
      .ok("Unloaded tsifulator namespace")
    }
  }, error = function(e) {
    .warn(sprintf("Could not unload namespace: %s", conditionMessage(e)))
  })

  # 3. Remove the installed package
  old_ver <- tryCatch(as.character(utils::packageVersion("tsifulator")),
                      error = function(e) "not installed")
  tryCatch({
    suppressWarnings(utils::remove.packages("tsifulator"))
    .ok(sprintf("Removed v%s", old_ver))
  }, error = function(e) {
    .warn(sprintf("remove.packages failed: %s", conditionMessage(e)))
  })

  # 4. Install from source
  if (is.null(source_dir)) {
    # Try common locations relative to the user's working tree. No
    # hardcoded paths \u2014 those only work on the maintainer's machine.
    candidates <- c(
      file.path(getwd(), "r-addin"),
      getwd(),
      file.path(getwd(), "..", "r-addin"),
      file.path(dirname(getwd()), "r-addin")
    )
    candidates <- normalizePath(candidates, mustWork = FALSE)
    source_dir <- candidates[vapply(candidates, function(d) {
      file.exists(file.path(d, "DESCRIPTION"))
    }, logical(1))][1]
    if (is.na(source_dir)) {
      message("\u2717 Could not find r-addin source directory automatically.")
      message("  Pass source_dir explicitly, e.g.:")
      message("    tsifulator::reinstall(source_dir = \"~/path/to/r-addin\")")
      return(invisible(NA_character_))
    }
  }
  if (!file.exists(file.path(source_dir, "DESCRIPTION"))) {
    message("\u2717 No DESCRIPTION found in ", source_dir)
    return(invisible(NA_character_))
  }

  .say(sprintf("Installing from %s", source_dir))
  install_ok <- tryCatch({
    if (requireNamespace("devtools", quietly = TRUE)) {
      devtools::install(source_dir, quiet = quiet, upgrade = "never")
    } else {
      utils::install.packages(source_dir, repos = NULL, type = "source",
                              quiet = quiet)
    }
    TRUE
  }, error = function(e) {
    message("\u2717 Install failed: ", conditionMessage(e))
    FALSE
  })

  if (!install_ok) return(invisible(NA_character_))

  # 5. Verify version actually updated
  source_ver <- tryCatch(
    read.dcf(file.path(source_dir, "DESCRIPTION"))[, "Version"],
    error = function(e) "?"
  )
  new_ver <- tryCatch(as.character(utils::packageVersion("tsifulator")),
                      error = function(e) "?")
  if (identical(new_ver, old_ver) || !identical(new_ver, unname(source_ver))) {
    .warn(sprintf(
      "Installed version is %s but source is %s. The package file lock ",
      new_ver, source_ver
    ))
    .warn("may still be held. Quit RStudio fully and run this in a Terminal:")
    .warn(sprintf("  Rscript --vanilla -e 'remove.packages(\"tsifulator\")'"))
    .warn(sprintf("  Rscript --vanilla -e 'devtools::install(\"%s\")'", source_dir))
  } else {
    .ok(sprintf("tsifulator v%s installed successfully", new_ver))
  }

  if (!quiet) message("")
  invisible(new_ver)
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

  # Detect stale installs: compare loaded version vs source on disk if we
  # can find the repo. Helps users notice when devtools::install() silently
  # fails due to a file lock.
  stale_warning <- ""
  for (cand in c(file.path(getwd(), "r-addin"),
                 getwd(),
                 file.path(dirname(getwd()), "r-addin"))) {
    desc_file <- file.path(cand, "DESCRIPTION")
    if (file.exists(desc_file)) {
      src_ver <- tryCatch(
        unname(read.dcf(desc_file)[, "Version"]),
        error = function(e) NA_character_
      )
      if (!is.na(src_ver) && !identical(src_ver, pkg_version)) {
        stale_warning <- sprintf(
          "  \u26a0 STALE INSTALL: loaded v%s but source on disk is v%s.\n" ,
          pkg_version, src_ver
        )
        stale_warning <- paste0(
          stale_warning,
          "    Run tsifulator::reinstall() or see status docs for a clean install.\n"
        )
      }
      break
    }
  }

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
  api_key_set <- !is.na(get_api_key())
  message(sprintf("  Anthropic API key:     %s",
                  if (api_key_set) "configured (BYOK)" else "NOT SET — run set_api_key()"))
  if (nzchar(stale_warning)) {
    message("")
    message(stale_warning)
  }
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
