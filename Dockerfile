FROM node:20-bullseye

RUN apt-get update && apt-get install -y \
  libreoffice \
  libreoffice-impress \
  fonts-dejavu \
  fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
COPY server/package*.json server/
COPY client/package*.json client/

RUN npm install
RUN npm --prefix server install
RUN npm --prefix client install

COPY . .

RUN npm --prefix client run build

ENV NODE_ENV=production
ENV SOFFICE_PATH=soffice

EXPOSE 8787
CMD ["npm", "--prefix", "server", "start"]
