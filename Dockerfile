FROM node:20-bullseye

# LibreOffice for PPTX -> PNG thumbnails + common fonts
RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy manifests first for better Docker cache
COPY package*.json ./
COPY server/package*.json server/
COPY client/package*.json client/

# Install root deps (if any), then server + client deps explicitly
RUN npm install
RUN npm --prefix server install
RUN npm --prefix client install

# Copy the rest of the code
COPY . .

# Build client for production
RUN npm --prefix client run build

ENV NODE_ENV=production
ENV SOFFICE_PATH=soffice
EXPOSE 8787

CMD ["npm", "--prefix", "server", "start"]
