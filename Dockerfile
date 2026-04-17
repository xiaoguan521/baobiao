FROM node:22-bookworm-slim

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci --omit=dev || npm install --omit=dev

COPY src ./src
COPY public ./public
COPY config ./config
COPY 模板.xlsx ./模板.xlsx

ENV PORT=3000
ENV ORACLE_USER=damoxing
ENV ORACLE_PASSWORD=Damoxing123!
ENV ORACLE_DSN=127.0.0.1:51521/FREEPDB1
ENV TEMPLATE_PATH=/app/模板.xlsx
ENV OUTPUT_DIR=/app/generated
ENV DOWNLOAD_ROOT=/app/generated
ENV REPORT_RULES_PATH=/app/config/report-rules.json
ENV REPORT_PUBLIC_BASE_URL=

RUN mkdir -p /app/generated

EXPOSE 3000

CMD ["npm", "start"]
