FROM node:18-slim
WORKDIR /app
COPY package.json .
RUN npm install --omit=dev
COPY server.js .
COPY src/ src/
COPY public/ public/
ENV PORT=8080
CMD ["node", "server.js"]
