# Tsifulator.ai — Boutique Firm Demo Script
**Target:** Associate / VP at a boutique M&A or PE firm
**Format:** Zoom call or in-person, screen share
**Duration:** 5 minutes
**Goal:** Get them to say "how do I sign up"

---

## Pre-Demo Checklist (do this 5 min before the call)

- [ ] Run `bash start.sh` in terminal — wait for "Excel add-in ready"
- [ ] Open Excel with a blank workbook
- [ ] Open the Tsifulator.ai sidebar (Home → Add-ins → Tsifulator.ai)
- [ ] Sign in to your account
- [ ] Open RStudio with a blank script tab
- [ ] Open Tsifulator.ai in RStudio (Addins → Tsifulator.ai)
- [ ] Zoom in your Excel sidebar so it's legible on screen share
- [ ] Close Slack, email, everything — clean screen

---

## The Script

### Opening Line (say this while they're settling in)

> "I'm going to show you something your analysts are going to love.
> No slides. Just a live demo. Watch what happens when I type this."

---

### ACT 1 — The Model Builder (Excel, 90 seconds)

**Type this into the Tsifulator.ai Excel sidebar:**
```
Build me a 3-year LBO model for a $50M EBITDA business
at 8x entry multiple, 60% debt financing, 5% interest rate
```

*[Watch the model appear — don't say anything. Let the silence work.]*

**Then say:**
> "That took 8 seconds. Your junior analyst takes 3 hours to build
> that from scratch, and it has errors. This doesn't."

**Type this second command:**
```
Add revenue of 400M in Year 1 growing at 12% annually,
COGS at 65% of revenue
```

*[Watch it fill in the Income Statement in real time]*

**Say:**
> "Notice it read what was already in the sheet and added to it.
> It didn't overwrite anything. It understood the structure."

---

### ACT 2 — The R Analysis (RStudio, 60 seconds)

*[Switch screen share to RStudio — keep it smooth, no fumbling]*

**Type this into the Tsifulator.ai RStudio panel:**
```
I just built an LBO model in Excel with 400M revenue
growing 12% annually and 50M EBITDA. Plot a 3-year
revenue and EBITDA projection using ggplot2
```

*[Watch the code appear in the script editor AND the plot render in the Plots pane]*

**Say:**
> "Same brain. It knows what we built in Excel. The code
> is right there in the editor — your analyst can see it,
> edit it, run it again. Nothing is a black box."

---

### ACT 3 — The Memory (30 seconds)

**Close the Excel sidebar and reopen it. Type:**
```
What have we been working on?
```

*[Claude recalls the LBO model from memory]*

**Say:**
> "It remembered. Close it, come back tomorrow, it still
> knows your deal. Every assumption, every structure.
> That's shared memory across your entire workflow."

---

### The Close (say this, then stop talking)

> "Your analysts spend 4-6 hours building models like this.
> They spend another 2 hours QA'ing them for errors.
>
> Tsifulator does the structure in 30 seconds, the analyst
> reviews and refines — that's the workflow.
>
> We're in alpha. Founding seats are $49 a month.
> That's less than one hour of your analyst's billable time.
>
> I have 3 spots left for founding clients."

*[Stop. Don't say anything else. Let them respond.]*

---

## Handling Objections

**"How is this different from Copilot for Excel?"**
> "Copilot suggests. Tsifulator executes. It also works across
> Excel AND your R environment with shared memory — Copilot
> is trapped inside one app. We're the operating layer between them."

**"Is our data secure?"**
> "Your data never leaves your machine during a session.
> The AI processes your request — not your raw spreadsheet data.
> We're also building an on-premise version for firms with stricter compliance needs."

**"Can it do [specific model type]?"**
> "Yes — type it in right now, let's see."
> *[Let them drive. Whatever they type, it will try to build it.
> This is the most powerful moment — they become the demo.]*

**"What happens when it makes a mistake?"**
> "Same as when your analyst makes a mistake — you catch it and fix it.
> The difference is Tsifulator gives you a starting point in 30 seconds
> instead of 3 hours. The analyst is still in the loop."

**"We'd need to test it first."**
> "That's exactly what the alpha is for. Sign up, use it for 30 days,
> tell me what your team needs. Founding clients shape the roadmap."

---

## The Follow-Up Email (send within 1 hour)

Subject: Tsifulator.ai — Founding Access

Hi [Name],

Great speaking today. As promised — founding access details:

**$49/month per seat** (will be $149 at general release)
**Includes:** Excel add-in, RStudio integration, shared memory, unlimited model builds

To get started, reply to this email and I'll send you the setup link.

The 3 founding spots are first-come, first-served.

Nicholas
Founder, Tsifulator.ai

---

## The 3 Commands That Always Land

If you have less than 2 minutes, just run these three:

1. `Build me an LBO model structure for a 50M EBITDA business`
2. `Add a returns analysis showing IRR at 3x, 4x and 5x MOIC`
3. `Summarize this model in 3 bullet points I can put in an email`

Command 3 is the one that makes people lean forward.
A financial analyst typing "summarize this for my email"
and watching a draft appear is the moment they get it.
