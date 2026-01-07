FROM node:20-bullseye

RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy manifests first (better caching)
COPY package*.json ./
COPY server/package*.json server/
COPY client/package*.json client/

# Install deps explicitly
RUN npm install
RUN npm --prefix server install
RUN npm --prefix client install

# 🔎 Verify express exists (this will FAIL the build if it doesn't)
RUN test -f /app/server/node_modules/express/package.json

# Copy rest of code
COPY . .

# Build client
RUN npm --prefix client run build

ENV NODE_ENV=production
ENV SOFFICE_PATH=soffice
EXPOSE 8787

CMD ["node", "/app/server/index.js"]
