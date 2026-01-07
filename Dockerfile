# ---------- Builder (build client only) ----------
FROM node:20-bullseye AS builder
WORKDIR /app

# Install client deps + build (use npm install, not npm ci)
COPY client/package*.json ./client/
RUN cd client && npm install

COPY client ./client
RUN cd client && npm run build


# ---------- Runtime (server + libreoffice) ----------
FROM node:20-bullseye

# LibreOffice for PPTX -> PNG thumbnails + common fonts
RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy server manifests and install server deps (use npm install)
COPY server/package*.json ./server/
RUN cd server && npm install --omit=dev

# Copy server source
COPY server ./server

# Copy built client dist into the place your server serves from
COPY --from=builder /app/client/dist ./client/dist

# Optional: sanity check express exists (keep while debugging)
RUN test -f /app/server/node_modules/express/package.json

ENV NODE_ENV=production
ENV SOFFICE_PATH=soffice

EXPOSE 8787
CMD ["node", "server/index.js"]
