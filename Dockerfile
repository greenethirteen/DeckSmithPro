FROM node:20-bullseye

# LibreOffice for PPTX -> PNG thumbnails + common fonts
RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

# Install deps (root postinstall installs server+client deps too)
RUN npm install

# Build client for production
RUN npm --prefix client run build

ENV NODE_ENV=production
ENV PORT=8787
ENV SOFFICE_PATH=soffice
# You can mount a Railway volume here and set TMP_DIR to it
# ENV TMP_DIR=/app/server/.tmp

EXPOSE 8787
CMD ["npm", "--prefix", "server", "start"]
