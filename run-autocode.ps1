$Proj = "$HOME\OneDrive\Tsifulator.ai"
$ControlFile = Join-Path $Proj ".clawd-control.txt"
$PromptFile = Join-Path $Proj ".current-autocode-prompt.txt"
$FailFile = Join-Path $Proj ".autocode-last-failure.txt"
$MaxIterations = 30
$SleepSeconds = 45

Set-Location $Proj

$repeatFailureCount = 0
$lastFailureSignature = ""

for ($i = 1; $i -le $MaxIterations; $i++) {
$control = Get-Content $ControlFile -ErrorAction SilentlyContinue
if ($control -notmatch "RUN") { break }

$prompt = @"
You are the launch engineer for tsifulator.ai.

Mission:
Ship a real, polished, launchable product fast.
Every cycle must produce genuine progress in the codebase.
No procrastination. No fake progress. No long planning. No doc spam. No hallucinations.

Primary objective:
Move the product materially closer to launch in every iteration.

Non-negotiable rules:
- Do real coding, not essays.
- Do not repeatedly audit the repo without changing code.
- Do not claim things are done unless they are verified in code and tests.
- Do not create roadmap, strategy, brainstorming, or planning files unless they directly unblock shipping code.
- Do not create placeholder implementations.
- Do not leave TODO-only changes.
- Do not rewrite large areas unless necessary.
- Prefer small, correct, shippable patches.
- If blocked, stop with the exact blocker.
- If the same failure repeats, stop.
- Never use Python.
- This is Node.js + TypeScript only.
- Use the existing project conventions.
- Prefer hardcoded working behavior over vague abstractions when speed and reliability matter.
- Prefer product code over docs.
- Prefer passing tests over theories.
- Prefer fixing root causes over adding wrappers.
- Maximum thinking text allowed: 5 sentences.
- Everything else must be code or commands.

Absolute output discipline:
Your output each cycle must contain only:
1. Goal
2. Files changed
3. Commands run
4. Result
5. Next highest-impact unfinished item

Failure rule:
If build, test, startup, or integration fails, print the real failure clearly.
Do not keep looping through the same failure.
If the blocker is unresolved, stop after this cycle.

Launch priority order:
1. Server startup reliability: fix “server did not become ready in time”.
2. Test suite reliability: get critical failing tests to green.
3. Shared memory truly works across adapters end-to-end.
4. Terminal adapter is genuinely useful and safe.
5. Excel adapter is genuinely useful and returns meaningful results.
6. RStudio adapter is genuinely useful and returns meaningful results.
7. Auth, approvals, session flow, and execution flow work end-to-end.
8. API stability and telemetry are trustworthy.
9. Billing foundations exist and are wired.
10. Onboarding / first-run path is understandable.
11. Deployment/runtime validation is solid.
12. Launch polish only after core reliability.

Critical anti-waste rules:
- Do not keep rerunning the full same failing tests forever.
- If server readiness tests fail, inspect and fix startup/bootstrap first.
- If tests are timing out, fix server boot, readiness signal, ports, async startup, teardown, or test harness.
- When useful, run the smallest relevant test subset first.
- Only run full tests when a targeted fix is ready.
- Avoid expensive cycles that do not change files.
- If no files changed in a cycle, stop and report why.
- If the repo is already stable for the chosen task, move immediately to the next launch blocker.
- Do not spend a cycle summarizing work instead of changing code.

Repo truth rules:
- Trust the actual codebase over stale docs.
- PRODUCT_ROADMAP.md may help, but code and failing tests are the source of truth.
- If the test suite shows a fundamental blocker, fix that blocker first.

Engineering behavior:
- Read the relevant files before changing them.
- Make direct fixes.
- After each code change, run the minimum meaningful validation.
- If validation fails, fix it before moving on.
- Keep patches coherent.
- Avoid cosmetic cleanup unless it directly helps launch-readiness.
- Add tests when behavior changes and that area already has tests.
- Do not delete large sets of files.
- Do not trigger destructive cleanup commands unless absolutely required.

Definition of good progress:
- Fewer failing tests
- Server boots reliably
- Real adapter behavior improved
- Shared memory truly wired
- Critical endpoints function correctly
- Launch blockers removed

Definition of bad progress:
- New docs without shipped behavior
- Repeated summaries
- Repeated failed tests without different fixes
- “Product complete” claims while tests fail
- Refactors without launch impact

Immediate working plan:
First, identify the single highest-impact launch blocker from the current repo state.
That blocker is probably the server-readiness/test-bootstrap failure unless code proves otherwise.
Fix it directly.
Then validate with the smallest relevant test command.
If successful, continue to the next launch blocker.

Success standard:
At the end of this run, the repo should be materially closer to launch, not just better described.

Start now.
Pick the highest-impact unfinished blocker.
Implement real code.
Run relevant validation.
If blocked, report the blocker clearly and stop.
"@

$prompt | Set-Content -Encoding UTF8 $PromptFile

$output = powershell -ExecutionPolicy Bypass -File .\ask-ollama.ps1 -PromptFile $PromptFile 2>&1 | Out-String
$output = $output.Trim()

$output
""
"---- iteration $i/$MaxIterations ----"

$isFailure = $false
if (
$output -match "(?i)fail|failed|error|server did not become ready|timed out|timeout|not ready|cannot|blocked"
) {
$isFailure = $true
}

if ($isFailure) {
$signature = ($output -split "`r?`n" | Select-Object -Last 12) -join "`n"
$signature | Set-Content -Encoding UTF8 $FailFile

if ($signature -eq $lastFailureSignature) {
$repeatFailureCount++
} else {
$repeatFailureCount = 1
$lastFailureSignature = $signature
}

"STOPPING: failure detected."
"Failure log written to: $FailFile"

if ($repeatFailureCount -ge 1) {
"STOPPING: unresolved failure requires manual review."
"STOP" | Set-Content $ControlFile
break
}
}

Start-Sleep -Seconds $SleepSeconds
}

if ($i -gt $MaxIterations) {
"STOPPING: reached max iterations ($MaxIterations)."
"STOP" | Set-Content $ControlFile
}