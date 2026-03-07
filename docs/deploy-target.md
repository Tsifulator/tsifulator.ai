# Deploy Target Decision

## Chosen: Railway

**Why Railway for beta/launch:**

- **Zero-ops deployment** — push to git, auto-deploys
- **Dockerfile support** — our Dockerfile works as-is
- **Persistent volumes** — SQLite DB survives redeploys via Railway volumes
- **Free tier available** — $5/month hobby plan covers beta usage
- **HTTPS built-in** — automatic TLS certificates
- **Custom domains** — easy to add `api.tsifulator.ai` later
- **Logs & metrics** — built-in log viewer for debugging

## Setup steps

1. Create Railway account at https://railway.app
2. Connect GitHub repo
3. Railway auto-detects Dockerfile
4. Add environment variables:
   - `NODE_ENV=production`
   - `JWT_DEV_SECRET=<generate-strong-secret>`
   - `OPENAI_API_KEY=<your-key>`
   - `CORS_ORIGIN=<frontend-origin>`
   - `DB_PATH=/app/data/tsifulator.db`
5. Add persistent volume mounted at `/app/data`
6. Deploy

## Verification

After deploy, verify per `docs/production-deploy-checklist.md`:
- `GET /health` returns 200
- `POST /auth/dev-login` works
- `POST /chat` returns response
- `GET /chat/stream` streams SSE

## Cost estimate (beta)

| Component | Cost |
|---|---|
| Railway Hobby | $5/month |
| OpenAI API | Variable (usage-based) |
| Domain (optional) | ~$12/year |
| **Total** | **~$7/month** |

## Migration path

If Railway becomes limiting (scale, cost, control):
- **VPS (Hetzner/DigitalOcean)**: $4-6/month, full control, manual setup
- **Fly.io**: Similar to Railway, better edge distribution
- **Self-hosted**: Free, requires uptime management

The Dockerfile is portable — switching hosts means changing deploy commands, not code.

## Alternatives considered

| Option | Verdict |
|---|---|
| Railway | ✅ Chosen — best balance of simplicity and features |
| Render | Good but slower deploys, volume support less mature |
| Fly.io | Great but more CLI setup, overkill for single-region beta |
| VPS | Too much ops work for beta phase |
| Vercel/Netlify | Not suitable — they're for frontend/serverless, not persistent servers |
