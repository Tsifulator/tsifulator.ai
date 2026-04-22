# Installing tsifl — for analysts

This guide is written for someone who uses Excel all day but doesn't write
code. If you can run one Terminal command, you can install tsifl.

There are **three pieces**, installed in this order:

1. **Excel add-in** — the side panel that talks to you inside Excel.
2. **Desktop agent** *(optional)* — a small helper that handles things
   the Excel add-in can't do on its own (Solver, Scenario Manager,
   Analysis ToolPak, Install/Uninstall add-ins).
3. **Your account** — sign in inside the add-in so your project memory
   sticks around between sessions.

You'll need about **15 minutes**.

---

## What you need before you start

- A **Mac** (Windows instructions coming later).
- **Excel** (Microsoft 365 subscription — needs a modern version that
  supports add-ins). Excel 2016 and newer work.
- **Terminal** (pre-installed on every Mac: Cmd+Space → type "Terminal").
- **Node.js** (free download at https://nodejs.org — grab the LTS version).
- A **tsifl account** — ask the tsifl team for an invite.

## Step 1 — Install the Excel add-in (10 min)

Open Terminal. Copy-paste each block below one at a time.

### 1a. Get the code

```bash
cd ~
git clone https://github.com/Tsifulator/tsifulator.ai.git
cd tsifulator.ai/excel-addin
```

*(If you don't have `git`, install Xcode Command Line Tools:
`xcode-select --install` — Mac will prompt you.)*

### 1b. Install the add-in's dependencies

```bash
npm install
```

This takes 1–2 minutes. Ignore warnings about peer dependencies.

### 1c. Generate the local HTTPS certificates (one-time)

```bash
npx office-addin-dev-certs install
```

You'll get a keychain prompt — type your Mac password to allow the cert.

### 1d. Start the dev server

```bash
npm start
```

This opens Excel with tsifl sideloaded. Leave this Terminal window running —
if you close it, the add-in stops working. The add-in appears as a side
panel on the right; if you don't see it, go to **Home → Add-ins → My
Add-ins → SHARED FOLDER → tsifl**.

## Step 2 — Install the desktop agent *(optional)*

You only need this if you want tsifl to use features like **Solver**,
**Scenario Manager**, **Analysis ToolPak**, or **install/uninstall
add-ins**. Skip this step if you only need basic Excel work.

### 2a. Install the agent

Open a **new** Terminal window (leave the one from Step 1 running):

```bash
cd ~/tsifulator.ai/desktop-agent
pip3 install -r requirements.txt
```

### 2b. Set your Anthropic API key

The desktop agent uses Claude directly for some tasks. Get a key from
https://console.anthropic.com, then:

```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

*(Add this line to `~/.zshrc` to make it permanent.)*

### 2c. Start the agent

```bash
python3 agent.py
```

Leave this Terminal running too. The agent will sit and wait for tsifl to
ask for help.

## Step 3 — Sign in

In Excel, click the tsifl side panel. Use the email + password the tsifl
team gave you. Your account is stored in Supabase so your **memory and
history persist** across Excel restarts.

## Step 4 — Your first real task

Try this in the tsifl chat:

> Add the formula `=C14*D7` to Calculator C18. Nothing else.

You should see tsifl write exactly that one formula, then the Memory
panel at the top of the taskpane should flip to `1 cell remembered`.

Then try something harder:

> Now create named ranges on Price Solver: Selling_Price from B12:C12,
> Total_Commission from B14:C14, Net_Commission from B17:C17.
> Nothing else.

Tsifl should only add those three named ranges and **not redo C18** —
that's the memory working.

---

## Troubleshooting

**Add-in shows a blank screen.** The dev server (Step 1d) isn't running.
Re-run `npm start` from `~/tsifulator.ai/excel-addin`.

**Taskpane says "Can't reach tsifl backend".** The Railway backend might
be down or your network is blocked. Check
https://focused-solace-production-6839.up.railway.app/chat/debug/guards
in your browser — you should see a JSON response. If not, ping the
tsifl team.

**Header shows an old version number.** Right-click inside the taskpane
→ Reload. Still stuck? Close and re-insert the add-in from Home → My
Add-ins.

**Solver / Scenario Manager / Analysis ToolPak steps fail with
"No desktop agent is running".** You skipped Step 2. Start the agent
in a separate Terminal with `cd ~/tsifulator.ai/desktop-agent && python3 agent.py`.

**Memory panel shows entries from a different workbook.** Click `Reset`
in the Memory panel. State is keyed by workbook filename — if you
renamed a file, old memory won't move over.

**"The requested resource doesn't exist" errors.** You saw a phantom
sheet get created (Office.js rejected it because it already existed in
a different form). Click Reset on the Memory panel, then re-send your
request.

---

## What to do if you're stuck

Screenshot the taskpane error + the Terminal window running `npm start`
and paste them to the tsifl team. Include the **build version** showing
at the top of the taskpane (looks like `tsiflik@bc.edu · v58`).
