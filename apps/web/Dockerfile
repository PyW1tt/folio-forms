FROM oven/bun:1.4.0@sha256:5ff609364c049b54eb0ff560ec96319729a972078ef2c755d758f0c6ef89c2d6 AS builder

WORKDIR /app
ARG VITE_API_ORIGIN
ENV VITE_API_ORIGIN=${VITE_API_ORIGIN}
COPY . .
RUN bun install --frozen-lockfile
RUN bun run --cwd apps/web build

FROM joseluisq/static-web-server:2.44.0-alpine@sha256:c6704bc8f1fe05378d91c3288495ce2e0131cf31cf8956f117cb1c7d83c49c31

ENV SERVER_CONFIG_FILE=/etc/sws.toml

COPY apps/web/sws.toml /etc/sws.toml
COPY apps/web/sws-assets.toml /etc/sws-assets.toml
COPY --from=builder /app/apps/web/dist /home/sws/public

EXPOSE 80
