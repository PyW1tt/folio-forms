dcup:
  docker compose --env-file .env -f compose.yaml up -d
  sleep 10
  curl -f "http://localhost:$(docker compose --env-file .env -f compose.yaml port gateway 80 | sed 's/.*://')/ready"

dcup-build:
  docker compose --env-file .env -f compose.yaml up -d --build
  docker compose --env-file .env -f compose.yaml logs --tail=100 server
  sleep 10
  curl -f "http://localhost:$(docker compose --env-file .env -f compose.yaml port gateway 80 | sed 's/.*://')/ready"

dcdown:
  docker compose --env-file .env -f compose.yaml down

dev:
  docker compose --env-file .env -p folio-forms-dev -f compose.yaml -f compose.dev.yaml up -d --wait postgres rustfs onlyoffice dev-gateway
  docker compose --env-file .env -p folio-forms-dev -f compose.yaml -f compose.dev.yaml up --no-recreate rustfs-init
  bun run --cwd packages/db --env-file ../../apps/server/.env db:migrate
  bun run dev

dev-down:
  docker compose --env-file .env -p folio-forms-dev -f compose.yaml -f compose.dev.yaml down

dev-down-v:
  docker compose --env-file .env -p folio-forms-dev -f compose.yaml -f compose.dev.yaml down -v
