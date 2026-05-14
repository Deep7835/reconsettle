# syntax=docker/dockerfile:1.7
# ---------- Stage 1: build the SPA ----------
FROM node:20-alpine AS build

WORKDIR /app

# Same-origin API calls in production — nginx proxies /api/* to backend
ARG VITE_API_BASE=""
ENV VITE_API_BASE=${VITE_API_BASE}

# Dependency layer (cached as long as package-lock.json doesn't change)
COPY package.json package-lock.json* ./
RUN npm ci --no-audit --no-fund

# App source (tighten via .dockerignore)
COPY . .

RUN npm run build


# ---------- Stage 2: nginx serves static files + proxies /api/ ----------
FROM nginx:1.27-alpine AS runtime

# Replace the default site
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Static files
COPY --from=build /app/dist /usr/share/nginx/html

# Convenience: a tiny entrypoint that lets you override the backend hostname
# at run-time via BACKEND_HOST env var without rebuilding the image.
RUN echo $'#!/bin/sh\n\
if [ -n "$BACKEND_HOST" ]; then\n\
  sed -i "s|proxy_pass http://backend:8000;|proxy_pass http://${BACKEND_HOST};|" /etc/nginx/conf.d/default.conf\n\
fi\n\
exec nginx -g "daemon off;"' > /docker-entrypoint-runtime.sh \
 && chmod +x /docker-entrypoint-runtime.sh

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
    CMD wget -qO- http://localhost/ >/dev/null 2>&1 || exit 1

CMD ["/docker-entrypoint-runtime.sh"]
