FROM node:22-slim
WORKDIR /app

# Install from the lockfile for reproducible builds
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY server.js .
COPY src/ src/
COPY public/ public/

ENV PORT=8080
# /app is not writable for the unprivileged user – keep the geocoding cache in /tmp
ENV KNOWN_CITIES_FILE=/tmp/known_cities.json
USER node

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s \
  CMD node -e "fetch('http://127.0.0.1:'+(process.env.PORT||8080)+'/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

CMD ["node", "server.js"]
