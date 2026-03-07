FROM node:22-slim AS base
WORKDIR /app

# Install pnpm
RUN corepack enable && corepack prepare pnpm@latest --activate

# Dependencies
COPY package.json pnpm-lock.yaml* pnpm-workspace.yaml* ./
RUN pnpm install --frozen-lockfile --prod=false && \
    pnpm rebuild better-sqlite3 esbuild

# Source
COPY tsconfig.json ./
COPY server/ ./server/
COPY clients/ ./clients/

# Data directory
RUN mkdir -p /app/data

ENV NODE_ENV=production
ENV PORT=4000
ENV HOST=0.0.0.0
ENV DB_PATH=/app/data/tsifulator.db

EXPOSE 4000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://localhost:4000/health').then(r=>{if(!r.ok)throw r.status}).catch(()=>process.exit(1))"

CMD ["node", "--import", "tsx", "server/src/index.ts"]
