
# Tsifulator.ai
# Project Excel revenue model with 10% annual growth
year1 <- data.frame(Year = 1, Revenue = 100, COGS = 60, Gross_Profit = 40)
year2 <- data.frame(Year = 2, Revenue = 100 * 1.1, COGS = 60 * 1.1, Gross_Profit = 40 * 1.1)
year3 <- data.frame(Year = 3, Revenue = 100 * 1.1^2, COGS = 60 * 1.1^2, Gross_Profit = 40 * 1.1^2)

revenue_projection <- rbind(year1, year2, year3)
print(revenue_projection)

# Tsifulator.ai

# Tsifulator.ai

**Agentic Sandbox for Financial Analysts.**
A cross-app AI operating layer for Excel, RStudio, and Terminal.

## Architecture

```
tsifulator.ai/
├── backend/              # Python (FastAPI) — the AI brain
│   ├── main.py           # Server entry point
│   ├── routes/
│   │   └── chat.py       # Chat endpoint
│   ├── services/
│   │   ├── claude.py     # Claude (Anthropic) integration
│   │   └── usage.py      # Task limit monitor
│   └── requirements.txt
├── excel-addin/          # Office JS — sidebar in Excel
│   ├── manifest.xml      # Excel add-in registration
│   └── src/
│       ├── taskpane.html
│       ├── taskpane.js
│       └── taskpane.css
├── .env.example          # Copy to .env, fill in your keys
└── .gitignore
```

## Setup

### 1. Environment
```bash
cp .env.example .env
# Fill in your ANTHROPIC_API_KEY and SUPABASE keys
```

### 2. Backend
```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
```

### 3. Excel Add-in (development)
```bash
cd excel-addin
npm install
npm start
```

## Pricing Tiers
- **Starter** — $20/month — 40 tasks/month
- **Pro** — $49/month — Unlimited tasks
