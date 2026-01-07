# ---------- Builder ----------
FROM node:20-bullseye AS builder

RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# --- Server deps (install first for caching) ---
COPY server/package*.json ./server/
RUN cd server && npm ci

# --- Client deps + build ---
COPY client/package*.json ./client/
RUN cd client && npm ci
COPY client ./client
RUN cd client && npm run build

# --- Copy server source (after deps) ---
COPY server ./server

# Sanity check: express must exist
RUN test -f /app/server/node_modules/express/package.json


# ---------- Runtime ----------
FROM node:20-bullseye

RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy server (with node_modules) + built client
COPY --from=builder /app/server ./server
COPY --from=builder /app/client/dist ./client/dist

ENV NODE_ENV=production
ENV PORT=8787
ENV SOFFICE_PATH=soffice

EXPOSE 8787
CMD ["node", "server/index.js"]
